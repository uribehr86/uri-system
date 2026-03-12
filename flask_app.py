from flask import Flask, render_template, request, redirect, url_for, session, flash
import psycopg2
from psycopg2 import pool
from psycopg2.extras import RealDictCursor
import os
from dotenv import load_dotenv
from functools import wraps
from datetime import datetime
import json

# טעינת הגדרות
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'uri_system_2026')

# Initialize Connection Pool
db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
try:
    db_pool = psycopg2.pool.SimpleConnectionPool(1, 10, db_url)
    print("✅ Database connection pool created successfully")
except Exception as e:
    print(f"❌ Error creating connection pool: {e}")
    db_pool = None

# פונקציה לחיבור לענן עם "הגנת תקיעה" ושימוש ב-Pool
def get_db_connection():
    if not db_pool:
        return None
    try:
        # Get connection from pool
        return db_pool.getconn()
    except Exception as e:
        print(f"❌ שגיאה: לא ניתן לקבל חיבור מהמאגר. {e}")
        return None

def release_db_connection(conn):
    if db_pool and conn:
        db_pool.putconn(conn)

@app.teardown_appcontext
def close_db(error):
    # This ensures connections are released if forgotten, 
    # though it's better to do it manually in routes.
    pass

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
            print(f"✅ משתמש {username} התחבר החיבור מהיר (hardcoded)")
            return redirect(url_for('portal'))
            
        # Check database
        conn = get_db_connection()
        if conn:
            try:
                cur = conn.cursor(cursor_factory=RealDictCursor)
                cur.execute("SELECT * FROM users WHERE username = %s AND password = %s", (username, password))
                user = cur.fetchone()
                cur.close()
                
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
                    print(f"✅ משתמש {username} התחבר דרך DB")
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

@app.route('/portal')
@login_required
def portal():
    return render_template('portal.html')

@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db_connection()
    if not conn:
        return "<h1>⚠️ המערכת לא מצליחה להתחבר לענן. בדוק חיבור אינטרנט.</h1>"
    
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("SELECT COUNT(*) as total FROM computers")
        total = cur.fetchone()['total']
        
        cur.execute("SELECT status, COUNT(*) as count FROM computers GROUP BY status")
        stats = cur.fetchall()
        
        # Stats dict based on template requirements
        stats_dict = {row['status']: row['count'] for row in stats}
        faulty_count = stats_dict.get('תקול', 0)
        not_in_cage_count = 0 
        
        cur.execute("""
            SELECT id, barcode, cage_name, cage_number, location, status, scan_time as last_seen 
            FROM computers 
            ORDER BY scan_time DESC NULLS LAST 
            LIMIT 10
        """)
        recent = cur.fetchall()
        cur.close()
        return render_template('dashboard.html', total=total, faulty=faulty_count, not_in_cage=not_in_cage_count, recent=recent)
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"❌ Error in dashboard:\n{error_details}")
        return f"<h1>❌ שגיאה בשליפת הנתונים מהענן.</h1><pre>{e}</pre>"
    finally:
        release_db_connection(conn)

@app.route('/manage-computers')
@app.route('/computers') # תמיכה בשני השמות
@login_required
def computers():
    conn = get_db_connection()
    if not conn: return redirect(url_for('dashboard'))
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT id, barcode, cage_name as name, location, status, notes, scan_time as last_seen 
            FROM computers 
            ORDER BY scan_time DESC NULLS LAST 
            LIMIT 100
        """)
        computers = cur.fetchall()
        cur.close()
        return render_template('computers.html', computers=computers)
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
        cur.execute("SELECT * FROM computers WHERE exam_appeal IS NOT NULL AND exam_appeal != ''")
        computers = cur.fetchall()
        cur.close()
        return render_template('exam.html', computers=computers)
    finally:
        release_db_connection(conn)

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
    barcode = request.json.get('barcode')
    if not barcode:
        return {"error": "No barcode provided"}, 400

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = conn.cursor(cursor_factory=RealDictCursor)
        # Check if computer exists
        cur.execute("SELECT * FROM computers WHERE barcode = %s", (barcode,))
        computer = cur.fetchone()
        
        if computer:
            # Update scan_time (last seen)
            cur.execute("UPDATE computers SET scan_time = NOW() WHERE id = %s", (computer['id'],))
            conn.commit()
            cur.close()
            return {"exists": True, "computer": computer}
        else:
            # Create new record
            cur.execute("""
                INSERT INTO computers (barcode, status, scan_time) 
                VALUES (%s, 'פעיל', NOW()) 
                RETURNING *
            """, (barcode,))
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
            cur.execute("UPDATE computers SET status = 'פעיל' WHERE barcode = %s", (barcode,))
            
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

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    print("\n🚀 URI SYSTEM IS LIVE!")
    print("🌐 Link: http://127.0.0.1:5000")
    print("📱 Mobile Link: http://10.0.0.31:5000\n")
    app.run(host='0.0.0.0', debug=True, port=5000)