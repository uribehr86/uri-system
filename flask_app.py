from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import psycopg2
import sqlite3

from psycopg2 import pool
from psycopg2.extras import RealDictCursor
import os
import threading
from dotenv import load_dotenv
from functools import wraps
from datetime import datetime
import json
import io
import base64
import qrcode
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import google.generativeai as genai
from google_sheets_sync import sync_inventory_to_sheets
from utils import format_history, summarize_history

# טעינת הגדרות
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'uri_system_2026')

from datetime import timedelta
app.permanent_session_lifetime = timedelta(days=30)

# Initialize Google AI (Gemini)
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# Initialize Connection Pool
db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
try:
    db_pool = psycopg2.pool.SimpleConnectionPool(2, 20, db_url)
    print("[OK] Database connection pool created successfully")
except Exception as e:
    print(f"[ERROR] Error creating connection pool: {e}")
    db_pool = None

# פונקציה לחיבור לענן עם "הגנת תקיעה" ושימוש ב-Pool
IS_LOCAL_MODE = (db_pool is None)


class SafeCursor:
    """Wrapper for cursor to handle %s -> ? translation for SQLite"""
    def __init__(self, cursor, is_sqlite=False):
        self.cursor = cursor
        self.is_sqlite = is_sqlite

    def execute(self, query, params=None):
        if self.is_sqlite and params:
            # Handle list for IN clauses and basic %s replacements
            if "IN (" in query and isinstance(params, (list, tuple)):
                # This is a bit tricky, but common in this app
                pass # Already handled by placeholders in most cases
            query = query.replace('%s', '?')
            # Handle ILIKE -> LIKE for SQLite (SQLite LIKE is case-insensitive usually, but ILIKE is Postgres specific)
            query = query.replace('ILIKE', 'LIKE')
            # Handle NOW() -> datetime('now')
            query = query.replace('NOW()', "datetime('now', 'localtime')")
            # Handle NULLS LAST (SQLite supports it in newer versions, but let's be safe)
            # query = query.replace('NULLS LAST', '') 
        
        try:
            if params:
                return self.cursor.execute(query, params)
            else:
                return self.cursor.execute(query)
        except Exception as e:
            print(f"[DB ERROR] Query: {query}")
            print(f"[DB ERROR] Params: {params}")
            raise e

    def fetchone(self): return self.cursor.fetchone()
    def fetchall(self): return self.cursor.fetchall()
    def close(self): return self.cursor.close()
    def __getattr__(self, name): return getattr(self.cursor, name)

def get_db_connection():
    global IS_LOCAL_MODE
    if not db_pool:
        # Fallback to SQLite if pool failed
        try:
            conn = sqlite3.connect('system_data.db', check_same_thread=False)
            conn.row_factory = sqlite3.Row
            IS_LOCAL_MODE = True
            return conn
        except Exception as e:
            print(f"[CRITICAL] SQLite connection failed: {e}")
            return None
    
    try:
        # Attempt Postgres from pool
        conn = db_pool.getconn()
        IS_LOCAL_MODE = False
        return conn
    except Exception as e:
        print(f"[FALLBACK] Cloud DB Error: {e}. Switching to Local SQLite.")
        try:
            conn = sqlite3.connect('system_data.db', check_same_thread=False)
            conn.row_factory = sqlite3.Row
            IS_LOCAL_MODE = True
            return conn
        except Exception as e2:
            print(f"[CRITICAL] All DB connections failed: {e2}")
            return None

def get_safe_cursor(conn):
    if IS_LOCAL_MODE:
        return SafeCursor(conn.cursor(), is_sqlite=True)
    else:
        return conn.cursor(cursor_factory=RealDictCursor)

def release_db_connection(conn):
    if IS_LOCAL_MODE:
        if conn: conn.close()
    elif db_pool and conn:
        db_pool.putconn(conn)


@app.teardown_appcontext
def close_db(error):
    # This ensures connections are released if forgotten, 
    # though it's better to do it manually in routes.
    pass

@app.context_processor
def utility_processor():
    def get_cage_color(cage):
        if not cage: return "inherit"
        val = sum(ord(c) for c in str(cage))
        hue = (val * 137) % 360
        return f"hsl({hue}, 70%, 65%)"
    return dict(get_cage_color=get_cage_color, IS_LOCAL_MODE=IS_LOCAL_MODE)


@app.template_filter('format_history')
def format_history_filter(val_str):
    return format_history(val_str)

@app.template_filter('summarize_history')
def summarize_history_filter(entry):
    return summarize_history(entry)

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session: return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/')
def index():
    return redirect(url_for('portal')) if 'user' in session else redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        print(f"DEBUG: Login attempt - username: '{username}', password: '{password}'")
        
        # Hardcoded super-admin fallback
        if (username.lower() == "uri" and password == "1234") or (username.lower() == "admin_uri" and password == "uri*"):
            session.update({
                'user': username,
                'user_id': 1,
                'username': "אורי מנהל מערכת",
                'role': 'admin'
            })
            session.permanent = True
            print(f"[OK] User {username} logged in (hardcoded fallback)")
            return redirect(url_for('portal'))
            
        # Check database
        conn = get_db_connection()
        if conn:
            try:
                cur = get_safe_cursor(conn)
                cur.execute("SELECT * FROM users WHERE username = %s AND password = %s", (username, password))
                user = cur.fetchone()
                
                if user:
                    # Update last active time
                    cur.execute("UPDATE users SET timestamp = NOW() WHERE id = %s", (user['id'],))
                    conn.commit()
                    
                    session.update({
                        'user': user['username'],
                        'user_id': user.get('id', 999), # In case id is missing
                        'username': user['username'],
                        'role': user['role']
                    })
                    session.permanent = True
                    print(f"[OK] User {username} logged in via DB")
                    return redirect(url_for('portal'))
                else:
                    flash("שם משתמש או סיסמה שגויים", "danger")
            except Exception as e:
                print(f"DB Login Error: {e}")
                flash("שגיאה בהתחברות למסד הנתונים", "danger")
            finally:
                release_db_connection(conn)
        else:
            flash("שגיאת חיבור למסד הנתונים", "danger")
            
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        if not username or not password:
            flash("יש להזין שם משתמש וסיסמה", "warning")
            return redirect(url_for('register'))
            
        conn = get_db_connection()
        if conn:
            try:
                cur = get_safe_cursor(conn)
                # Check if exists
                cur.execute("SELECT * FROM users WHERE username = %s", (username,))
                if cur.fetchone():
                    flash("שם המשתמש כבר קיים במערכת, בחר שם אחר או התחבר", "warning")
                else:
                    cur.execute("INSERT INTO users (username, password, role) VALUES (%s, %s, 'technician')", (username, password))
                    conn.commit()
                    flash(f"משתמש {username} נוצר בהצלחה! כעת ניתן להתחבר.", "success")
                    cur.close()
                    release_db_connection(conn)
                    return redirect(url_for('login'))
                cur.close()
            except Exception as e:
                print(f"DB Register Error: {e}")
                flash("שגיאה ביצירת המשתמש", "danger")
            finally:
                release_db_connection(conn)
        else:
            flash("שגיאת חיבור למסד הנתונים", "danger")
            
    return render_template('register.html')

@app.route('/portal')
@login_required
def portal():
    return render_template('portal.html')

@app.route('/dashboard')
@app.route('/manage-computers')
@app.route('/computers') # תמיכה בשני השמות
@login_required
def computers():
    search = request.args.get('q', '').strip()
    cage_search = request.args.get('cage_q', '').strip()
    status_filter = request.args.get('status', '').strip()
    page = request.args.get('page', 1, type=int)
    sort = request.args.get('sort', 'scan_time')
    direction = request.args.get('dir', 'desc').lower()
    
    per_page = 100
    offset = (page - 1) * per_page
    
    # Whitelist of allowed sort columns
    allowed_sorts = {
        'barcode': 'barcode',
        'case_number': 'case_number',
        'cage_number': 'cage_number',
        'status': 'status',
        'location': 'location',
        'scan_time': 'scan_time',
        'scan_time': 'scan_time',
        'exam_appeal': 'exam_appeal',
        'specs': 'specs',
        'project': 'project'
    }
    sort_col = allowed_sorts.get(sort, 'scan_time')
    sort_dir = 'ASC' if direction == 'asc' else 'DESC'
    
    conn = get_db_connection()
    if not conn: return "<h1>⚠️ המערכת לא מצליחה להתחבר לענן. בדוק חיבור אינטרנט.</h1>"
    try:
        cur = get_safe_cursor(conn)
        
        # Dashboard Stats
        cur.execute("SELECT COUNT(*) as total FROM computers")
        total_in_db = cur.fetchone()['total']
        
        cur.execute("SELECT status, COUNT(*) as count FROM computers GROUP BY status")
        stats = cur.fetchall()
        stats_dict = {row['status']: row['count'] for row in stats}
        faulty_count = stats_dict.get('תקול', 0)
        # Count computers with no cage assigned
        cur.execute("SELECT COUNT(*) as count FROM computers WHERE (cage_number IS NULL OR TRIM(cage_number) = '') AND (cage_name IS NULL OR TRIM(cage_name) = '')")
        not_in_cage_count = cur.fetchone()['count']
        
        # Base query for computers and count of total matching records
        base_where = " WHERE 1=1"
        params = []
        
        # Free search across multiple fields
        if search:
            # Normalize search term if it's purely digits (barcode-like)
            norm_search = re.sub(r'^0+(?=\d)', '', search)
            base_where += " AND (barcode ILIKE %s OR barcode ILIKE %s OR case_number ILIKE %s OR location ILIKE %s OR notes ILIKE %s OR exam_appeal ILIKE %s)"
            search_val = f"%{search}%"
            norm_val = f"%{norm_search}%"
            params.extend([search_val, norm_val, search_val, search_val, search_val, search_val])
            
        # Dedicated cage search
        if cage_search:
            base_where += " AND (cage_number = %s OR cage_name ILIKE %s)"
            cs = f"%{cage_search}%"
            params.extend([cage_search, cs])
            
        if status_filter:
            base_where += " AND status = %s"
            params.append(status_filter)

        # Get total matching count for pagination
        cur.execute("SELECT COUNT(*) as cnt FROM computers" + base_where, params)
        total_matching = cur.fetchone()['cnt']
        total_pages = (total_matching + per_page - 1) // per_page
            
        # Query results for current page
        query = "SELECT id, barcode, case_number, cage_name, cage_number, location, status, exam_appeal, specs, project, notes, last_technician, scan_time as last_seen FROM computers"
        query += base_where
        
        # Order by logic
        if sort_col == 'scan_time':
            query += f" ORDER BY {sort_col} {sort_dir} NULLS LAST"
        else:
            query += f" ORDER BY {sort_col} {sort_dir}"
            
        query += " LIMIT %s OFFSET %s"
        
        cur.execute(query, params + [per_page, offset])
        computers = cur.fetchall()
        cur.close()
        
        return render_template('computers.html', 
                               computers=computers, 
                               search=search, 
                               status_filter=status_filter,
                               total=total_in_db, 
                               faulty=faulty_count, 
                               not_in_cage=not_in_cage_count,
                               page=page,
                               total_pages=total_pages,
                               total_matching=total_matching,
                               sort=sort,
                               direction=direction,
                               cage_search=cage_search)
    finally:
        release_db_connection(conn)

# נתיבים נוספים שנדרשים בטמפלייט base.html
@app.route('/add-computer', methods=['GET', 'POST'])
@login_required
def add_computer():
    if request.method == 'POST':
        data = request.form
        conn = get_db_connection()
        if not conn: return "DB connection failed", 500
        try:
            cur = get_safe_cursor(conn)
            barcode = re.sub(r'^0+(?=\d)', '', data['barcode'].strip())
            
            # Check for existing barcode
            cur.execute("SELECT id FROM computers WHERE barcode = %s", (barcode,))
            if cur.fetchone():
                flash(f"שגיאה: מחשב {barcode} כבר קיים במערכת!", "danger")
                return redirect(url_for('add_computer'))

            project = data.get('project', '').strip()

            cur.execute("""
                INSERT INTO computers (barcode, case_number, cage_number, status, location, exam_appeal, specs, project, notes, scan_time, last_technician)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, NOW(), %s)
            """, (barcode, data['case_number'], data['cage_number'], data['status'], data['location'], data['exam_appeal'], data['specs'], project, data['notes'], session.get('username')))
            conn.commit()
            cur.close()
            flash("מחשב נוסף בהצלחה!", "success")
            return redirect(url_for('computers'))
        except Exception as e:
            conn.rollback()
            flash(f"שגיאה בהוספת מחשב: {e}", "danger")
        finally:
            release_db_connection(conn)
            
    return render_template('computer_form.html', action='add', computer=None)

@app.route('/edit-computer/<int:cid>', methods=['GET', 'POST'])
@app.route('/edit-computer/<int:cid>', methods=['GET', 'POST'])
@login_required
def edit_computer(cid):
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    try:
        cur = get_safe_cursor(conn)
        if request.method == 'POST':
            data = request.form
            cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
            old_val = cur.fetchone()
            
            project = data.get('project', '').strip()

            cur.execute("""
                UPDATE computers 
                SET case_number=%s, cage_number=%s, status=%s, location=%s, exam_appeal=%s, specs=%s, project=%s, notes=%s, last_technician=%s
                WHERE id=%s
            """, (data['case_number'], data['cage_number'], data['status'], data['location'], data['exam_appeal'], data['specs'], project, data['notes'], session.get('username'), cid))
            
            cur.execute("""
                INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                VALUES (%s, %s, 'Manual Edit', %s, %s)
            """, (
                cid, 
                session.get('username'), 
                json.dumps(dict(old_val), default=str) if old_val else None, 
                json.dumps(dict(data), default=str)
            ))
            
            conn.commit()
            cur.close()
            flash("פרטי המחשב עודכנו!", "success")
            return redirect(url_for('computers'))
            
        cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
        computer = cur.fetchone()
        cur.close()
        return render_template('computer_form.html', action='edit', computer=computer)
    finally:
        release_db_connection(conn)

@app.route('/delete-computer/<int:cid>', methods=['POST'])
@login_required
def delete_computer(cid):
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    try:
        cur = get_safe_cursor(conn)
        # Only admin_uri can hard delete
        if session.get('user') == 'admin_uri':
            cur.execute("DELETE FROM computers WHERE id = %s", (cid,))
            flash("המחשב נמחק מהמערכת סופית (admin_uri)", "warning")
        else:
            cur.execute("UPDATE computers SET status = 'ממתין למחיקה' WHERE id = %s", (cid,))
            flash("הבקשה למחיקת המחשב הועברה לאישור מנהל העל (admin_uri)", "info")
            
        conn.commit()
        cur.close()
    except Exception as e:
        flash(f"שגיאה במחיקה: {e}", "danger")
    finally:
        release_db_connection(conn)
    return redirect(url_for('computers'))

@app.route('/scanner')
@login_required
def scanner():
    return render_template('scanner.html')

@app.route('/exam')
@login_required
def exam_page():
    conn = get_db_connection()
    if not conn: return redirect(url_for('dashboard'))
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT * FROM computers WHERE exam_appeal IS NOT NULL AND TRIM(exam_appeal) != ''")
        computers = cur.fetchall()
        cur.close()
        return render_template('exam.html', computers=computers)
    finally:
        release_db_connection(conn)

import re

@app.route('/history')
@login_required
def history_page():
    conn = get_db_connection()
    if not conn: return redirect(url_for('dashboard'))
    try:
        cur = get_safe_cursor(conn)
        cur.execute("""
            SELECT h.*, c.barcode 
            FROM inventory_history h
            LEFT JOIN computers c ON h.computer_id = c.id
            ORDER BY h.timestamp DESC
            LIMIT 100
        """)
        history = cur.fetchall()
        cur.close()
        
        # Safe processing for templates
        processed_history = []
        for h in history:
            h_dict = dict(h)
            ts = h_dict['timestamp']
            if ts and isinstance(ts, str):
                try:
                    # SQLite default format
                    h_dict['timestamp'] = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
                except:
                    pass
            processed_history.append(h_dict)

        return render_template('history.html', history=processed_history)
    finally:
        release_db_connection(conn)

@app.route('/api/process-scan', methods=['POST'])
@login_required
def process_scan():
    barcode = request.json.get('barcode', '').strip()
    if not barcode:
        return {"error": "No barcode provided"}, 400
    barcode = re.sub(r'^0+(?=\d)', '', barcode)

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        # Check if computer exists
        cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
        computer = cur.fetchone()
        
        if computer:
            # Update the existing record instead of inserting a duplicate
            cur.execute("""
                UPDATE computers 
                SET scan_time = NOW(), last_technician = %s
                WHERE id = %s
            """, (session.get('username'), computer['id'],))
            
            # Fetch the updated record (SQLite doesn't support RETURNING)
            cur.execute("SELECT * FROM computers WHERE id = %s", (computer['id'],))
            new_computer = dict(cur.fetchone())
            
            # Fetch last technician and time
            cur.execute("SELECT technician, timestamp FROM inventory_history WHERE computer_id = %s ORDER BY timestamp DESC LIMIT 1", (new_computer['id'],))
            hist = cur.fetchone()
            new_computer['last_technician'] = hist['technician'] if hist and hist['technician'] else "לא ידוע"
            
            ts = hist['timestamp'] if hist else None
            if ts and not isinstance(ts, str):
                new_computer['last_scan_time'] = ts.strftime("%d/%m/%Y %H:%M")
            else:
                new_computer['last_scan_time'] = str(ts or '')
            
            conn.commit()
            cur.close()
            # Trigger Google Sheets sync in background
            threading.Thread(target=sync_inventory_to_sheets).start()
            
            # מחזירים את המידע הקיים כדי שהטופס יתמלא נכון
            return {"exists": True, "computer": new_computer}
        else:
            # Create completely new record
            cur.execute("""
                INSERT INTO computers (barcode, status, scan_time, notes, last_technician) 
                VALUES (%s, 'תקין', NOW(), %s, %s) 
            """, (barcode, None, session.get('username')))
            
            last_id = cur.lastrowid if hasattr(cur, 'lastrowid') else None
            if last_id:
                cur.execute("SELECT * FROM computers WHERE id = %s", (last_id,))
                new_computer = dict(cur.fetchone())
            else:
                # Fallback if lastrowid fails
                cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
                new_computer = dict(cur.fetchone())

            conn.commit()
            cur.close()
            # Trigger Google Sheets sync in background
            threading.Thread(target=sync_inventory_to_sheets).start()
            
            return {"exists": False, "computer": new_computer}
            
    except Exception as e:
        print(f"Error in process_scan: {e}")
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/update-computer', methods=['POST'])
@login_required
def api_update_computer():
    data = request.json
    cid = data.get('id')
    if not cid: return {"error": "No ID provided"}, 400

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        # Get old values for history
        cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
        old_val = cur.fetchone()
        
        # Update
        updates = []
        params = []
        for key in ['case_number', 'cage_number', 'status', 'location', 'exam_appeal', 'specs', 'project', 'notes']:
            if key in data:
                val = data[key]
                updates.append(f"{key} = %s")
                params.append(val)
        
        # Always update last_technician on scan update
        updates.append("last_technician = %s")
        params.append(session.get('username'))

        if updates:
            params.append(cid)
            cur.execute(f"UPDATE computers SET {', '.join(updates)} WHERE id = %s", params)
            
            # Record history
            cur.execute("""
                INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                VALUES (%s, %s, 'Update via Scan', %s, %s)
            """, (
                cid, 
                session.get('username'), 
                json.dumps(dict(old_val), default=str) if old_val else None, 
                json.dumps(data, default=str)
            ))
            
            conn.commit()
            cur.close()
            # Trigger Google Sheets sync in background
            threading.Thread(target=sync_inventory_to_sheets).start()
            
            return {"success": True}
        return {"success": False, "message": "No fields to update"}
        
    except Exception as e:
        print(f"Error in api_update_computer: {e}")
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

# ── AI ASSISTANT (GEMINI) ───────────────────────────────────────────
@app.route('/api/ai-chat', methods=['POST'])
@login_required
def api_ai_chat():
    data = request.json
    user_msg = data.get('message', '').strip()
    if not user_msg:
        return {"error": "No message provided"}, 400

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        # Fetch stats for context
        cur.execute("SELECT COUNT(*) as total FROM computers")
        total = cur.fetchone()['total']
        
        cur.execute("SELECT status, COUNT(*) as count FROM computers GROUP BY status")
        status_stats = cur.fetchall()
        status_desc = ", ".join([f"{row['status']}: {row['count']}" for row in status_stats])
        
        # Build system context
        system_prompt = f"""
        אתה עוזר ה-AI של מערכת URI לניהול מלאי מחשבים. 
        הנתונים הנוכחיים במערכת הם:
        - סה"כ מחשבים: {total}
        - סטטוסים: {status_desc}
        
        ענה למשתמש בעברית בצורה עוזרת, מקצועית וקצרה. 
        אם המשתמש שואל על המצב, השתמש בנתונים שלעיל.
        משתמש נוכחי: {session.get('username')}
        """
        
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content([system_prompt, user_msg])
        
        return {"response": response.text}
        
    except Exception as e:
        print(f"AI Error: {e}")
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

# ── API: Batch Operations ──────────────────────────────────────────────
@app.route('/api/batch-update', methods=['POST'])
@login_required
def api_batch_update():
    data = request.json
    ids = data.get('ids', [])
    updates = data.get('updates', {})
    
    if not ids or not updates:
        return {"success": False, "error": "Missing ids or updates"}, 400
        
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        
        set_clauses = []
        params = []
        for key in ['location', 'cage_number', 'cage_name', 'status', 'exam_appeal', 'specs', 'project', 'ministry', 'notes']:
            if key in updates:
                set_clauses.append(f"{key} = %s")
                params.append(updates[key])
                
        if not set_clauses:
            return {"success": False, "error": "No valid fields provided"}, 400
            
        params.extend(ids)
        placeholders = ','.join(['%s'] * len(ids))
        
        query = f"UPDATE computers SET {', '.join(set_clauses)} WHERE id IN ({placeholders})"
        cur.execute(query, params)
        
        # Log to history for each (simplified to avoid mass select first, assuming identical change)
        for cid in ids:
            cur.execute("""
                INSERT INTO inventory_history (computer_id, technician, change_type, new_value)
                VALUES (%s, %s, 'Batch Update', %s)
            """, (cid, session.get('username'), json.dumps(updates, default=str)))
            
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error in batch-update: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/batch-delete', methods=['POST'])
@login_required
def api_batch_delete():
    # Only admin should be able to trigger this in the UI, verify on server too
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        return {"success": False, "error": "Unauthorized"}, 403
        
    data = request.json
    ids = data.get('ids', [])
    if not ids: return {"success": False, "error": "No ids provided"}, 400
    
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        placeholders = ','.join(['%s'] * len(ids))
        cur.execute(f"DELETE FROM computers WHERE id IN ({placeholders})", ids)
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error in batch-delete: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)


# ── API: Google Sheets Sync ──────────────────────────────────────────
@app.route('/api/sync-to-sheets', methods=['POST'])
@login_required
def api_sync_to_sheets():
    """
    Manual trigger for Google Sheets sync.
    """
    # Only admin or experienced technicians should sync
    # For now, allow all logged in users as requested "connect table"
    success, message = sync_inventory_to_sheets()
    if success:
        return {"success": True, "message": message}
    else:
        return {"success": False, "error": message}, 500


# ── API: שליפת מידע כלוב ──────────────────────────────────────────────
@app.route('/api/cage/<cage_id>', methods=['GET'])
@login_required
def api_get_cage(cage_id):
    """מחזיר מידע על כלוב + רשימת המחשבים בו"""
    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)

        # שליפת פרטי הכלוב (אם קיים בטבלת cages)
        cur.execute("SELECT * FROM cages WHERE cage_id = %s", (cage_id,))
        cage = cur.fetchone()

        # שליפת מחשבים בכלוב זה
        cur.execute("""
            SELECT id, barcode, status, location, specs, scan_time, notes
            FROM computers
            WHERE cage_number = %s OR cage_name = %s
            ORDER BY scan_time DESC NULLS LAST
        """, (cage_id, cage_id))
        computers_in_cage = cur.fetchall()

        # סטטיסטיקות
        total = len(computers_in_cage)
        status_counts = {}
        for c in computers_in_cage:
            s_val = c.get('status')
            s = s_val if s_val is not None else 'לא ידוע'
            status_counts[s] = status_counts.get(s, 0) + 1  # type: ignore

        cur.close()
        return {
            "cage": dict(cage) if cage else {"cage_id": cage_id, "name": f"כלוב {cage_id}"},
            "computers": [dict(c) for c in computers_in_cage],
            "total": total,
            "status_counts": status_counts
        }
    except Exception as e:
        print(f"Error in api_get_cage: {e}")
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

# ── API: סריקה מהירה (ללא חלון) ───────────────────────────────────────
@app.route('/api/fast-scan', methods=['POST'])
@login_required
def api_fast_scan():
    """
    סריקה מהירה - מעדכן מחשב ישירות ללא תצוגת UI מפורטת.
    user שולח: barcode, location, cage_number, cage_name, status
    """
    data = request.json
    barcode = data.get('barcode', '').strip()
    if not barcode:
        return {"success": False, "error": "ברקוד חסר"}, 400
    barcode = re.sub(r'^0+(?=\d)', '', barcode)

    location   = data.get('location', '')
    cage_number = data.get('cage_number', '')
    cage_name  = data.get('cage_name', cage_number)
    status     = data.get('status', 'תקין')
    specs      = data.get('specs', '')
    project    = data.get('project', '')
    ministry   = data.get('ministry', '') or get_ministry_for_project(project)
    technician = session.get('username', 'לא ידוע')

    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
        computer = cur.fetchone()

        old_val = dict(computer) if computer else None
        last_technician = "לא ידוע"
        notes_val = ""

        if old_val:
            notes_val = old_val.get('notes') or ""
            fields_to_update = []
            params = []
            
            for key in ['location', 'cage_number', 'cage_name', 'status', 'specs', 'project', 'ministry', 'exam_appeal']:
                if key in data:
                    fields_to_update.append(f"{key} = %s")
                    params.append(data[key])
            
            # Special auto-ministry logic
            if 'project' in data and 'ministry' not in data:
                project_name = data['project']
                ministry = ''
                if 'רופאי' in project_name or 'שיניים' in project_name: ministry = 'משרד הבריאות'
                elif 'משפטים' in project_name: ministry = 'משרד המשפטים'
                elif 'עבודה' in project_name: ministry = 'משרד העבודה'
                elif 'חינוך' in project_name: ministry = 'משרד החינוך'
                fields_to_update.append("ministry = %s")
                params.append(ministry)
                
            fields_to_update.extend(["scan_time = NOW()", "last_technician = %s"])
            params.extend([technician, old_val['id']])
            
            query = f"UPDATE computers SET {', '.join(fields_to_update)} WHERE id = %s RETURNING *"
            cur.execute(query, params)
            new_computer = cur.fetchone()
            
            cur.execute("SELECT technician, timestamp FROM inventory_history WHERE computer_id = %s ORDER BY timestamp DESC LIMIT 1", (new_computer['id'],))
            hist = cur.fetchone()
            if hist and hist['technician']:
                last_technician = hist['technician']
                last_scan_time = hist['timestamp'].strftime("%d/%m/%Y %H:%M") if hist['timestamp'] else ""
        else:
            fields = ['barcode', 'scan_time', 'last_technician']
            params = [barcode, technician]
            
            for key in ['location', 'cage_number', 'cage_name', 'status', 'specs', 'project', 'ministry', 'exam_appeal']:
                if key in data:
                    fields.append(key)
                    params.append(data[key])
                    
            if 'project' in data and 'ministry' not in data:
                fields.append('ministry')
                project_name = data['project']
                ministry = ''
                if 'רופאי' in project_name or 'שיניים' in project_name: ministry = 'משרד הבריאות'
                elif 'משפטים' in project_name: ministry = 'משרד המשפטים'
                elif 'עבודה' in project_name: ministry = 'משרד העבודה'
                elif 'חינוך' in project_name: ministry = 'משרד החינוך'
                params.append(ministry)
                
            if 'status' not in data:
                fields.append('status')
                params.append('תקין')
                
            placeholders = ', '.join(['%s'] * len(fields))
            query = f"INSERT INTO computers ({', '.join(fields)}) VALUES ({placeholders}) RETURNING *"
            cur.execute(query, params)
            new_computer = cur.fetchone()

        cur.execute("""
            INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
            VALUES (%s, %s, 'Fast Scan', %s, %s)
        """, (
            new_computer['id'],
            technician,
            json.dumps(old_val, default=str) if old_val else None,
            json.dumps(data, default=str)
        ))

        conn.commit()
        cur.close()
        return {
            "success": True, 
            "barcode": barcode, 
            "is_new": old_val is None,
            "previous_cage": old_val.get('cage_number', '') if old_val else "",
            "notes": notes_val
        }

    except Exception as e:
        print(f"Error in api_fast_scan: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)


@app.route('/api/admin_approve_delete', methods=['POST'])
@login_required
def admin_approve_delete():
    # Only admin_uri can approve deletes
    if session.get('user') != 'admin_uri':
        return {"success": False, "error": "Unauthorized: Only admin_uri can perform this action"}, 403
        
    data = request.json
    barcode = data.get('barcode')
    action = data.get('action')
    
    if not barcode or not action:
        return {"success": False, "error": "Missing parameters"}, 400
        
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB Error"}, 500
    try:
        cur = get_safe_cursor(conn)
        if action == 'hard_delete':
            cur.execute("DELETE FROM computers WHERE barcode = %s", (barcode,))
        elif action == 'restore':
            cur.execute("UPDATE computers SET status = 'תקין' WHERE barcode = %s", (barcode,))
            
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error in admin_approve_delete: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/manage-users')
@login_required
def manage_users():
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        flash("אין לך הרשאה לגשת לעמוד זה", "danger")
        return redirect(url_for('portal'))
        
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    try:
        cur = get_safe_cursor(conn)
        # 1. Fetch users
        cur.execute("SELECT username, role, timestamp FROM users ORDER BY timestamp DESC")
        users = cur.fetchall()
        
        # 2. Fetch pending deletions
        cur.execute("SELECT barcode as computer_number, cage_number FROM computers WHERE status = 'ממתין למחיקה'")
        pending_raw = cur.fetchall()
        
        # Format pending for template
        pending = []
        for p in pending_raw:
            pending.append({
                'computer_number': p['computer_number'],
                'cage_number': p['cage_number'] or 'לא ידוע',
                'scanned_by': 'טכנאי'
            })
            
        cur.close()
        return render_template('manage_users.html', users=users, pending=pending)
    finally:
        release_db_connection(conn)

@app.route('/api/add_user', methods=['POST'])
@login_required
def api_add_user():
    # Only admin should be able to add users
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        return {"success": False, "error": "Unauthorized"}, 403
        
    data = request.json
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    role = data.get('role', 'technician')
    
    if not username or not password:
        return {"success": False, "error": "Missing username or password"}, 400
        
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("INSERT INTO users (username, password, role) VALUES (%s, %s, %s)", (username, password, role))
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error adding user: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/delete_user/<username>', methods=['DELETE'])
@login_required
def api_delete_user(username):
    # Only admin should be able to delete users
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        return {"success": False, "error": "Unauthorized"}, 403
        
    if username == 'admin_uri':
        return {"success": False, "error": "Cannot delete super-admin"}, 400
        
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("DELETE FROM users WHERE username = %s", (username,))
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error deleting user: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/cages')
@login_required
def cages_page():
    conn = get_db_connection()
    if not conn: return "DB Error", 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("""
            SELECT
                c.cage_id, c.name, c.location, c.notes,
                COUNT(comp.id) AS computer_count,
                SUM(CASE WHEN comp.status = 'תקין' THEN 1 ELSE 0 END) AS ok_count,
                SUM(CASE WHEN comp.status = 'תקול' THEN 1 ELSE 0 END) AS broken_count,
                SUM(CASE WHEN comp.status NOT IN ('תקין','תקול') AND comp.status IS NOT NULL THEN 1 ELSE 0 END) AS other_count
            FROM cages c
            LEFT JOIN computers comp ON comp.cage_number = c.cage_id OR comp.cage_name = c.cage_id
            GROUP BY c.id, c.cage_id, c.name, c.location, c.notes
            ORDER BY computer_count DESC
        """)
        cages = cur.fetchall()
        cur.close()
        return render_template('cages.html', cages=cages)
    except Exception as e:
        return f"<h1>Error: {e}</h1>", 500
    finally:
        release_db_connection(conn)

@app.route('/api/cage/save', methods=['POST'])
@login_required
def api_save_cage():
    data = request.json
    cage_id = data.get('cage_id', '').strip()
    existing_id = data.get('existing_id', '').strip()
    if not cage_id: return {'success': False, 'error': 'cage_id is required'}, 400
    conn = get_db_connection()
    if not conn: return {'success': False, 'error': 'DB Error'}, 500
    try:
        cur = get_safe_cursor(conn)
        if existing_id:
            cur.execute("UPDATE cages SET name=%s, location=%s, notes=%s, updated_at=NOW() WHERE cage_id=%s", 
                        (data.get('name',''), data.get('location',''), data.get('notes',''), existing_id))
        else:
            cur.execute("INSERT INTO cages (cage_id, name, location, notes) VALUES (%s, %s, %s, %s) ON CONFLICT (cage_id) DO UPDATE SET name=EXCLUDED.name, location=EXCLUDED.location, notes=EXCLUDED.notes, updated_at=NOW()", 
                        (cage_id, data.get('name',''), data.get('location',''), data.get('notes','')))
        conn.commit()
        cur.close()
        return {'success': True}
    finally: release_db_connection(conn)

@app.route('/api/pack-cage-photo', methods=['POST'])
@login_required
def api_pack_cage_photo():
    data = request.json
    cage_id = data.get('cage_id', '').strip()
    image_b64 = data.get('image', '')
    
    if not cage_id or not image_b64:
        return {"success": False, "error": "Missing cage_id or image"}, 400
        
    # Extract base64 part
    if ',' in image_b64:
        image_b64 = image_b64.split(',', 1)[1]
        
    try:
        image_data = base64.b64decode(image_b64)
    except Exception as e:
        return {"success": False, "error": "Invalid image data"}, 400

    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        # Call Gemini Vision
        model = genai.GenerativeModel('gemini-1.5-flash')
        prompt = "Look at the handwritten numbers written in white on the edges of the laptops in the cage. Extract all of them. Return ONLY a JSON array of strings (e.g. [\"1064\", \"366\", \"1480\"]). Do not add any markdown, comments, or other text."
        response = model.generate_content([
            {"mime_type": "image/jpeg", "data": image_data},
            prompt
        ])
        
        try:
            # Clean response text just in case Gemini adds markdown
            text = response.text.strip()
            if text.startswith('```json'): text = text[7:]
            if text.startswith('```'): text = text[3:]
            if text.endswith('```'): text = text[:-3]
            text = text.strip()
            extracted_numbers = json.loads(text)
            if not isinstance(extracted_numbers, list):
                extracted_numbers = []
        except Exception as e:
            print(f"Gemini parse error: {e}. Raw response: {response.text}")
            return {"success": False, "error": "Could not parse AI response as JSON list."}, 500
            
        cur = get_safe_cursor(conn)
        
        results = {
            "success_count": 0,
            "transferred_count": 0,
            "new_count": 0,
            "details": []
        }
        
        technician = session.get('username', 'לא ידוע')
        
        # Process each extracted number
        for barcode in extracted_numbers:
            barcode = str(barcode).strip()
            barcode = re.sub(r'^0+(?=\d)', '', barcode)
            if not barcode: continue
            
            cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
            computer = cur.fetchone()
            
            if computer:
                old_val = dict(computer)
                prev_cage = old_val.get('cage_number', '')
                
                # Check if it's already in this cage
                is_transfer = (prev_cage and prev_cage != cage_id)
                
                # Update
                cur.execute("""
                    UPDATE computers 
                    SET cage_number = %s, cage_name = %s, scan_time = NOW(), last_technician = %s
                    WHERE id = %s
                """, (cage_id, cage_id, technician, old_val['id']))
                
                # Log history
                cur.execute("""
                    INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                    VALUES (%s, %s, 'Photo Pack', %s, %s)
                """, (
                    old_val['id'],
                    technician,
                    json.dumps(old_val, default=str),
                    json.dumps({"cage_number": cage_id, "cage_name": cage_id}, default=str)
                ))
                
                if is_transfer:
                    results["transferred_count"] += 1
                    results["details"].append({"barcode": barcode, "status": "transferred", "prev_cage": prev_cage})
                else:
                    results["success_count"] += 1
                    results["details"].append({"barcode": barcode, "status": "updated"})
                    
            else:
                # New computer
                cur.execute("""
                    INSERT INTO computers (barcode, cage_number, cage_name, status, scan_time, last_technician)
                    VALUES (%s, %s, %s, 'תקין', NOW(), %s) RETURNING id
                """, (barcode, cage_id, cage_id, technician))
                new_id = cur.fetchone()['id']
                
                cur.execute("""
                    INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                    VALUES (%s, %s, 'Photo Pack (New)', %s, %s)
                """, (
                    new_id,
                    technician,
                    None,
                    json.dumps({"cage_number": cage_id, "cage_name": cage_id, "status": "תקין"}, default=str)
                ))
                
                results["new_count"] += 1
                results["details"].append({"barcode": barcode, "status": "new"})
                
        conn.commit()
        cur.close()
        
        # Trigger async sync
        threading.Thread(target=sync_inventory_to_sheets).start()
        
        results["total_extracted"] = len(extracted_numbers)
        results["success"] = True
        return results

    except Exception as e:
        print(f"Error in api_pack_cage_photo: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/scan-dashboard')
@login_required
def scan_dashboard():
    return render_template('scan_dashboard.html')

@app.route('/api/scan-stats')
@login_required
def api_scan_stats():
    conn = get_db_connection()
    if not conn: return {'error': 'DB Error'}, 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT COUNT(*) as cnt FROM inventory_history WHERE timestamp::date = CURRENT_DATE AND change_type IN ('Fast Scan', 'Update via Scan')")
        today_total = cur.fetchone()['cnt']
        cur.execute("SELECT COUNT(*) as cnt FROM computers")
        total_computers = cur.fetchone()['cnt']
        cur.execute("SELECT COUNT(*) as cnt FROM computers WHERE status = 'תקול'")
        broken = cur.fetchone()['cnt']
        cur.execute("SELECT technician, COUNT(*) as count, MAX(timestamp) as last_scan FROM inventory_history WHERE timestamp::date = CURRENT_DATE AND change_type IN ('Fast Scan', 'Update via Scan') GROUP BY technician ORDER BY count DESC")
        workers = [dict(r) for r in cur.fetchall()]
        cur.close()
        return {'today_total': today_total, 'total_computers': total_computers, 'broken': broken, 'workers': workers}
    finally: release_db_connection(conn)

@app.route('/export/computers')
@login_required
def export_computers():
    conn = get_db_connection()
    if not conn: return 'DB Error', 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT barcode, cage_number, cage_name, location, status, case_number, exam_appeal, notes, scan_time FROM computers ORDER BY cage_number, barcode")
        rows = cur.fetchall()
        cur.close()
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'מחשבים'
        ws.sheet_view.rightToLeft = True
        headers = ['ברקוד', 'כלוב', 'שם כלוב', 'מיקום', 'סטטוס', 'מספר תיק', 'מבחן/ערעור', 'הערות', 'נסרק לאחרונה']
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
        for row_num, row in enumerate(rows, 2):
            values = [row.get('barcode',''), row.get('cage_number',''), row.get('cage_name',''), row.get('location',''), row.get('status',''), row.get('case_number',''), row.get('exam_appeal',''), row.get('notes',''), str(row.get('scan_time',''))[0:16] if row.get('scan_time') else '']
            for col_num, val in enumerate(values, 1): ws.cell(row=row_num, column=col_num, value=val)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True, download_name=f"inventory_{datetime.now().strftime('%Y%m%d')}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    finally: release_db_connection(conn)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# --- End of Routes ---

if __name__ == '__main__':
    print("\n[START] URI SYSTEM IS LIVE! (HTTPS ENABLED FOR MOBILE SCANNER)")
    print("URL Link: https://127.0.0.1:5000")
    print("Mobile Link: https://10.0.0.31:5000\n")
    print("[WARNING] When opening on iPhone, you will see a 'Not Private' warning. Click 'Show Details' -> 'Visit this website' to bypass and test the scanner.")
    app.run(host='0.0.0.0', debug=True, port=5000, ssl_context='adhoc')
