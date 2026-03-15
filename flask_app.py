from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import psycopg2
from psycopg2 import pool
from psycopg2.extras import RealDictCursor
import os
from dotenv import load_dotenv
from functools import wraps
from datetime import datetime
import json
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# טעינת הגדרות
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'uri_system_2026')

# Initialize Connection Pool
db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
try:
    db_pool = psycopg2.pool.SimpleConnectionPool(2, 20, db_url)
    print("[OK] Database connection pool created successfully")
except Exception as e:
    print(f"[ERROR] Error creating connection pool: {e}")
    db_pool = None

# פונקציה לחיבור לענן עם "הגנת תקיעה" ושימוש ב-Pool
def get_db_connection():
    if not db_pool:
        return None
    try:
        # Get connection from pool
        return db_pool.getconn()
    except Exception as e:
        print(f"[ERROR] שגיאה: לא ניתן לקבל חיבור מהמאגר. {e}")
        return None

def release_db_connection(conn):
    if db_pool and conn:
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
    return dict(get_cage_color=get_cage_color)

import json
import ast

@app.template_filter('format_history')
def format_history_filter(val_str):
    if not val_str:
        return ""
    try:
        val = json.loads(val_str)
    except:
        try:
            val = ast.literal_eval(str(val_str))
        except:
            return val_str
            
    if not isinstance(val, dict):
        return str(val)
        
    tmap = {
        'status': 'סטטוס',
        'location': 'מיקום',
        'cage_number': 'כלוב',
        'cage_name': 'שם כלוב',
        'case_number': 'תיק',
        'notes': 'הערות',
        'exam_appeal': 'מבחן/ערעור',
        'barcode': 'ברקוד'
    }
    
    parts = []
    for k, v in val.items():
        if k in ['id', 'computer_id', 'scan_time', 'created_at'] or v is None:
            continue
        parts.append(f"{tmap.get(k, k)}: {v}")
        
    return " | ".join(str(p) for p in parts) if parts else "פעולת מערכת"

@app.template_filter('summarize_history')
def summarize_history_filter(entry):
    if not entry:
        return ""
    
    def parse_val(v_str) -> dict:
        if not v_str: return {}
        if isinstance(v_str, dict): return v_str
        try:
            res = json.loads(v_str)
            if isinstance(res, dict):
                return res
        except:
            pass
        try:
            res = ast.literal_eval(v_str) if v_str else {}
            if isinstance(res, dict):
                return res
        except:
            pass
        return {}

    old = parse_val(entry.get('old_value'))
    new = parse_val(entry.get('new_value'))
    
    if not isinstance(old, dict):
        old = {}
    if not isinstance(new, dict):
        new = {}
    
    # Check for cage movements
    old_cage = old.get('cage_number') or old.get('cage_name')
    new_cage = new.get('cage_number') or new.get('cage_name')
    
    if old_cage and new_cage and str(old_cage).strip() != str(new_cage).strip():
        return f"העביר מכלוב {old_cage} לכלוב {new_cage}"
    
    if old_cage and not new_cage:
        # Check if it was moved to home or test
        loc = new.get('location', '')
        if 'בית' in str(loc) or 'בדיקה' in str(loc):
            return f"לקח מכלוב {old_cage} (עבודה מהבית/בדיקה)"
        return f"לקח מכלוב {old_cage}"
        
    # Default behavior: list what changed if not a simple cage move
    tmap = {
        'status': 'סטטוס',
        'location': 'מיקום',
        'cage_number': 'כלוב',
        'case_number': 'תיק',
        'notes': 'הערות',
        'exam_appeal': 'מבחן/ערעור'
    }
    
    changes = []
    # If it's a "Fast Scan" or "Update", we can see what's in 'new' that's different from 'old'
    for k, v in new.items():
        if k in ['id', 'computer_id', 'scan_time', 'barcode'] or v is None:
            continue
        old_v = old.get(k)
        if str(old_v) != str(v):
            changes.append(f"{tmap.get(k, k)}: {v}")
            
    if changes:
        return "שינוי: " + " | ".join(str(c) for c in changes)
    
    # If no changes detected in common fields, describe by change type
    ctype = entry.get('change_type', 'פעולה')
    if ctype == 'Fast Scan' and not old:
        return f"נוסף מחשב חדש"
        
    return ctype

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
            print(f"[OK] משתמש {username} התחבר החיבור מהיר (hardcoded)")
            return redirect(url_for('portal'))
            
        # Check database
        conn = get_db_connection()
        if conn:
            try:
                cur = conn.cursor(cursor_factory=RealDictCursor)
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
                    print(f"[OK] משתמש {username} התחבר דרך DB")
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
                cur = conn.cursor(cursor_factory=RealDictCursor)
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
        'exam_appeal': 'exam_appeal'
    }
    sort_col = allowed_sorts.get(sort, 'scan_time')
    sort_dir = 'ASC' if direction == 'asc' else 'DESC'
    
    conn = get_db_connection()
    if not conn: return "<h1>⚠️ המערכת לא מצליחה להתחבר לענן. בדוק חיבור אינטרנט.</h1>"
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)
        
        # Dashboard Stats
        cur.execute("SELECT COUNT(*) as total FROM computers")
        total_in_db = cur.fetchone()['total']
        
        cur.execute("SELECT status, COUNT(*) as count FROM computers GROUP BY status")
        stats = cur.fetchall()
        stats_dict = {row['status']: row['count'] for row in stats}
        faulty_count = stats_dict.get('תקול', 0)
        not_in_cage_count = 0 
        
        # Base query for computers and count of total matching records
        base_where = " WHERE 1=1"
        params = []
        
        # Free search across multiple fields
        if search:
            base_where += """ AND (
                barcode = %s OR 
                case_number = %s OR 
                cage_number = %s OR 
                cage_name ILIKE %s OR
                location ILIKE %s OR 
                notes ILIKE %s OR 
                exam_appeal ILIKE %s
            )"""
            s = f"%{search}%"
            params.extend([search, search, search, s, s, s, s])
            
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
        query = "SELECT id, barcode, case_number, cage_name, cage_number, location, status, exam_appeal, notes, last_technician, scan_time as last_seen FROM computers"
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
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO computers (barcode, case_number, cage_number, status, location, exam_appeal, notes, scan_time)
                VALUES (%s, %s, %s, %s, %s, %s, %s, NOW())
            """, (data['barcode'], data['case_number'], data['cage_number'], data['status'], data['location'], data['exam_appeal'], data['notes']))
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
        cur = conn.cursor(cursor_factory=RealDictCursor)
        if request.method == 'POST':
            data = request.form
            cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
            old_val = cur.fetchone()
            
            cur.execute("""
                UPDATE computers 
                SET case_number=%s, cage_number=%s, status=%s, location=%s, exam_appeal=%s, notes=%s
                WHERE id=%s
            """, (data['case_number'], data['cage_number'], data['status'], data['location'], data['exam_appeal'], data['notes'], cid))
            
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
        cur = conn.cursor()
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
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("SELECT * FROM computers WHERE exam_appeal ILIKE '%מבחן%' OR exam_appeal ILIKE '%ערעור%'")
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
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT h.*, c.barcode 
            FROM inventory_history h
            LEFT JOIN computers c ON h.computer_id = c.id
            ORDER BY h.timestamp DESC
            LIMIT 100
        """)
        history = cur.fetchall()
        cur.close()
        return render_template('history.html', history=history)
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
        cur = conn.cursor(cursor_factory=RealDictCursor)
        # Check if computer exists
        cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
        computer = cur.fetchone()
        
        if computer:
            # Update the existing record instead of inserting a duplicate
            cur.execute("""
                UPDATE computers 
                SET scan_time = NOW()
                WHERE id = %s
                RETURNING *
            """, (computer['id'],))
            new_computer = cur.fetchone()
            
            # Fetch last technician
            cur.execute("SELECT technician FROM inventory_history WHERE computer_id = %s ORDER BY timestamp DESC LIMIT 1", (new_computer['id'],))
            hist = cur.fetchone()
            new_computer['last_technician'] = hist['technician'] if hist and hist['technician'] else "לא ידוע"
            
            conn.commit()
            cur.close()
            # מחזירים את המידע הקיים כדי שהטופס יתמלא נכון
            return {"exists": True, "computer": new_computer}
        else:
            # Create completely new record
            cur.execute("""
                INSERT INTO computers (barcode, status, scan_time, notes) 
                VALUES (%s, 'תקין', NOW(), %s) 
                RETURNING *
            """, (barcode, None))
            new_computer = cur.fetchone()
            conn.commit()
            cur.close()
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
        cur = conn.cursor(cursor_factory=RealDictCursor)
        # Get old values for history
        cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
        old_val = cur.fetchone()
        
        # Update
        updates = []
        params = []
        for key in ['case_number', 'cage_number', 'status', 'location', 'exam_appeal', 'notes']:
            if key in data:
                updates.append(f"{key} = %s")
                params.append(data[key])
        
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
            return {"success": True}
        return {"success": False, "message": "No fields to update"}
        
    except Exception as e:
        print(f"Error in api_update_computer: {e}")
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
        cur = conn.cursor()
        
        set_clauses = []
        params = []
        for key in ['location', 'cage_number', 'cage_name', 'status', 'exam_appeal', 'notes']:
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
        cur = conn.cursor()
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


# ── API: שליפת מידע כלוב ──────────────────────────────────────────────
@app.route('/api/cage/<cage_id>', methods=['GET'])
@login_required
def api_get_cage(cage_id):
    """מחזיר מידע על כלוב + רשימת המחשבים בו"""
    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)

        # שליפת פרטי הכלוב (אם קיים בטבלת cages)
        cur.execute("SELECT * FROM cages WHERE cage_id = %s", (cage_id,))
        cage = cur.fetchone()

        # שליפת מחשבים בכלוב זה
        cur.execute("""
            SELECT id, barcode, status, location, scan_time, notes
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
    technician = session.get('username', 'לא ידוע')

    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)

        # בדיקה אם מחשב קיים
        cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
        computer = cur.fetchone()

        old_val = dict(computer) if computer else None

        last_technician = "לא ידוע"
        notes_val = ""

        if old_val:
            notes_val = old_val.get('notes') or ""
            cur.execute("""
                UPDATE computers 
                SET location = %s, cage_number = %s, cage_name = %s, status = %s, scan_time = NOW()
                WHERE id = %s
                RETURNING *
            """, (location, cage_number, cage_name, status, old_val['id']))
            new_computer = cur.fetchone()
            
            # Fetch last technician
            cur.execute("SELECT technician FROM inventory_history WHERE computer_id = %s ORDER BY timestamp DESC LIMIT 1", (new_computer['id'],))
            hist = cur.fetchone()
            if hist and hist['technician']:
                last_technician = hist['technician']
        else:
            cur.execute("""
                INSERT INTO computers (barcode, location, cage_number, cage_name, status, scan_time)
                VALUES (%s, %s, %s, %s, %s, NOW())
                RETURNING *
            """, (barcode, location, cage_number, cage_name, status))
            new_computer = cur.fetchone()

        # תיעוד היסטוריה
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
            "last_technician": last_technician,
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
        cur = conn.cursor()
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

@app.route('/fix-db')
def fix_db():
    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
            conn.commit()
            cur.close()
            return "DB fixed successfully! You can now visit /manage-users"
        except Exception as e:
            return f"Error: {e}"
        finally:
            release_db_connection(conn)
    return "Failed to connect"

@app.route('/manage-users')
@login_required
def manage_users():
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        flash("אין לך הרשאה לגשת לעמוד זה", "danger")
        return redirect(url_for('portal'))
        
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)
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
                'scanned_by': 'טכנאי' # Could join with history for exact user if needed
            })
            
        cur.close()
        return render_template('manage_users.html', users=users, pending=pending)
    finally:
        release_db_connection(conn)


# ════════════════════════════════════════════════════════════════
# CAGE MANAGEMENT
# ════════════════════════════════════════════════════════════════

@app.route('/cages')
@login_required
def cages_page():
    conn = get_db_connection()
    if not conn: return "DB Error", 500
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)
        # Get all cages with computer counts and status breakdowns
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

        # Also include cages that exist in computers but not in cages table
        cur.execute("""
            SELECT cage_number as cage_id, COUNT(*) as computer_count
            FROM computers
            WHERE cage_number IS NOT NULL AND cage_number != ''
            AND cage_number NOT IN (SELECT cage_id FROM cages)
            GROUP BY cage_number
            ORDER BY computer_count DESC
        """)
        implicit_cages = cur.fetchall()
        for ic in implicit_cages:
            cages.append({**dict(ic), 'name': None, 'location': None, 'notes': None, 'ok_count': 0, 'broken_count': 0, 'other_count': 0})

        cur.close()
        return render_template('cages.html', cages=cages)
    except Exception as e:
        print(f"Error in cages_page: {e}")
        return f"<h1>Error: {e}</h1>", 500
    finally:
        release_db_connection(conn)


@app.route('/api/cage/save', methods=['POST'])
@login_required
def api_save_cage():
    """Create or update a cage record"""
    data = request.json
    cage_id = data.get('cage_id', '').strip()
    existing_id = data.get('existing_id', '').strip()
    if not cage_id:
        return {'success': False, 'error': 'cage_id is required'}, 400

    conn = get_db_connection()
    if not conn: return {'success': False, 'error': 'DB Error'}, 500
    try:
        cur = conn.cursor()
        if existing_id:
            cur.execute("""
                UPDATE cages SET name=%s, location=%s, notes=%s, updated_at=NOW()
                WHERE cage_id=%s
            """, (data.get('name',''), data.get('location',''), data.get('notes',''), existing_id))
        else:
            cur.execute("""
                INSERT INTO cages (cage_id, name, location, notes)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (cage_id) DO UPDATE
                SET name=EXCLUDED.name, location=EXCLUDED.location, notes=EXCLUDED.notes, updated_at=NOW()
            """, (cage_id, data.get('name',''), data.get('location',''), data.get('notes','')))
        conn.commit()
        cur.close()
        return {'success': True}
    except Exception as e:
        print(f"Error saving cage: {e}")
        return {'success': False, 'error': str(e)}, 500
    finally:
        release_db_connection(conn)


# ════════════════════════════════════════════════════════════════
# SCAN DASHBOARD — LIVE STATS
# ════════════════════════════════════════════════════════════════

@app.route('/scan-dashboard')
@login_required
def scan_dashboard():
    return render_template('scan_dashboard.html')


@app.route('/api/scan-stats')
@login_required
def api_scan_stats():
    """Returns today's scan stats for the live dashboard"""
    conn = get_db_connection()
    if not conn: return {'error': 'DB Error'}, 500
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)

        # Scans today (via history)
        cur.execute("""
            SELECT COUNT(*) as cnt FROM inventory_history
            WHERE timestamp::date = CURRENT_DATE
            AND change_type IN ('Fast Scan', 'Update via Scan')
        """)
        today_total = cur.fetchone()['cnt']

        # Total computers
        cur.execute("SELECT COUNT(*) as cnt FROM computers")
        total_computers = cur.fetchone()['cnt']

        # Broken count
        cur.execute("SELECT COUNT(*) as cnt FROM computers WHERE status = 'תקול'")
        broken = cur.fetchone()['cnt']

        # Per worker today
        cur.execute("""
            SELECT technician, COUNT(*) as count, MAX(timestamp) as last_scan
            FROM inventory_history
            WHERE timestamp::date = CURRENT_DATE
            AND change_type IN ('Fast Scan', 'Update via Scan')
            AND technician IS NOT NULL
            GROUP BY technician
            ORDER BY count DESC
        """)
        workers = [dict(r) for r in cur.fetchall()]

        # Hourly (today)
        cur.execute("""
            SELECT EXTRACT(HOUR FROM timestamp) as hour, COUNT(*) as cnt
            FROM inventory_history
            WHERE timestamp::date = CURRENT_DATE
            AND change_type IN ('Fast Scan', 'Update via Scan')
            GROUP BY hour
            ORDER BY hour
        """)
        hourly = {int(r['hour']): r['cnt'] for r in cur.fetchall()}

        cur.close()
        return {
            'today_total': today_total,
            'total_computers': total_computers,
            'broken': broken,
            'workers': workers,
            'hourly': hourly
        }
    except Exception as e:
        print(f"Error in api_scan_stats: {e}")
        return {'error': str(e)}, 500
    finally:
        release_db_connection(conn)


# ════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ════════════════════════════════════════════════════════════════

@app.route('/export/computers')
@login_required
def export_computers():
    """Export computers to Excel with optional filters"""
    location_filter = request.args.get('location', '')
    cage_filter = request.args.get('cage', '')
    status_filter = request.args.get('status', '')

    conn = get_db_connection()
    if not conn: return 'DB Error', 500
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)
        query = "SELECT barcode, cage_number, cage_name, location, status, case_number, exam_appeal, notes, scan_time FROM computers WHERE 1=1"
        params = []
        if location_filter:
            query += " AND location = %s"; params.append(location_filter)
        if cage_filter:
            query += " AND (cage_number = %s OR cage_name = %s)"; params.extend([cage_filter, cage_filter])
        if status_filter:
            query += " AND status = %s"; params.append(status_filter)
        query += " ORDER BY cage_number, barcode"
        cur.execute(query, params)
        rows = cur.fetchall()
        cur.close()
    finally:
        release_db_connection(conn)

    # Build Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'מחשבים'
    ws.sheet_view.rightToLeft = True

    headers = ['ברקוד', 'כלוב', 'שם כלוב', 'מיקום', 'סטטוס', 'מספר תיק', 'מבחן/ערעור', 'הערות', 'נסרק לאחרונה']
    header_fill = PatternFill(start_color='1E3A5F', end_color='1E3A5F', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True, size=11)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')

    status_colors = {'תקול': 'FFCCCC', 'בתיקון': 'FFF3CC', 'תקין': 'CCFFCC'}
    for row_num, row in enumerate(rows, 2):
        values = [
            row.get('barcode',''), row.get('cage_number',''), row.get('cage_name',''),
            row.get('location',''), row.get('status',''), row.get('case_number',''),
            row.get('exam_appeal',''), row.get('notes',''),
            str(row.get('scan_time',''))[0:16] if row.get('scan_time') else ''  # type: ignore
        ]
        for col_num, val in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col_num, value=val)
            status = row.get('status','')
            if status in status_colors:
                cell.fill = PatternFill(start_color=status_colors[status], end_color=status_colors[status], fill_type='solid')  # type: ignore

    # Auto column widths
    for col in ws.columns:
        max_len = max((len(str(c.value or '')) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    fname = f"computers_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


if __name__ == '__main__':
    print("\n[START] URI SYSTEM IS LIVE! (HTTPS ENABLED FOR MOBILE SCANNER)")
    print("URL Link: https://127.0.0.1:5000")
    print("Mobile Link: https://10.0.0.31:5000\n")
    print("[WARNING] When opening on iPhone, you will see a 'Not Private' warning. Click 'Show Details' -> 'Visit this website' to bypass and test the scanner.")
    app.run(host='0.0.0.0', debug=True, port=5000, ssl_context='adhoc')
