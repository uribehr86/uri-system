from flask import Flask, render_template, request, redirect, url_for, session, flash
import psycopg2
from psycopg2.extras import RealDictCursor
import os
from dotenv import load_dotenv
from functools import wraps
from datetime import datetime

# טעינת הגדרות
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'uri_system_2026')

# פונקציה לחיבור לענן עם "הגנת תקיעה"
def get_db_connection():
    conn_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    try:
        # הוספנו connect_timeout=5 כדי שלא יתקע לך את המחשב
        conn = psycopg2.connect(conn_url, connect_timeout=5)
        return conn
    except Exception as e:
        print(f"❌ שגיאה: המערכת לא מצליחה להגיע לענן. {e}")
        return None

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session: return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/')
def index():
    return redirect(url_for('dashboard')) if 'user' in session else redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        # בדיקה מהירה: אם זה uri ו-1234, תכניס אותו
        if username.lower() == "uri" and password == "1234":
            session.update({
                'user': username,
                'user_id': 1,
                'username': "אורי בר",
                'role': 'admin'
            })
            print(f"✅ משתמש {username} התחבר בהצלחה")
            return redirect(url_for('dashboard'))
        else:
            flash("שם משתמש או סיסמה שגויים", "danger")
            
    return render_template('login.html')

@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db_connection()
    if not conn:
        return "<h1>⚠️ המערכת לא מצליחה להתחבר לענן. בדוק חיבור אינטרנט.</h1>"
    
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT COUNT(*) as total FROM computers")
        total = cur.fetchone()['total']
        
        cur.execute("SELECT status, COUNT(*) as count FROM computers GROUP BY status")
        stats = cur.fetchall()
        
        # התאמה לטמפלייט הקיים כדי שיראו את המספרים בעיצוב היפה
        stats_dict = {row['status']: row['count'] for row in stats}
        active_count = stats_dict.get('פעיל', 0)
        broken_count = stats_dict.get('תקול', 0)
        
        # שליפת סריקות אחרונות (כדי שיהיה תוכן בטבלה)
        cur.execute("""
            SELECT id, barcode, cage_name as name, location, status, scan_time as last_seen 
            FROM computers 
            ORDER BY scan_time DESC NULLS LAST 
            LIMIT 10
        """)
        recent = cur.fetchall()
        
        cur.close()
        conn.close()
        return render_template('dashboard.html', 
                               total=total, 
                               stats=stats, 
                               active=active_count, 
                               broken=broken_count, 
                               recent=recent)
    except Exception as e:
        print(f"Error in dashboard: {e}")
        return "<h1>❌ שגיאה בשליפת הנתונים מהענן.</h1>"

@app.route('/manage-computers')
@app.route('/computers') # תמיכה בשני השמות
@login_required
def manage_computers():
    conn = get_db_connection()
    if not conn: return redirect(url_for('dashboard'))
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    # שאילתה מותאמת לטמפלייט
    cur.execute("""
        SELECT id, barcode, cage_name as name, location, status, notes, scan_time as last_seen 
        FROM computers 
        ORDER BY scan_time DESC NULLS LAST 
        LIMIT 100
    """)
    computers = cur.fetchall()
    cur.close()
    conn.close()
    
    # בודק איזה טמפלייט קיים ומעדיף את computers.html המלא
    return render_template('computers.html', computers=computers)

# נתיבים נוספים שנדרשים בטמפלייט base.html
@app.route('/add-computer', methods=['GET', 'POST'])
@login_required
def add_computer():
    if request.method == 'POST':
        data = request.form
        conn = get_db_connection()
        if not conn: return "DB connection failed", 500
        cur = conn.cursor()
        try:
            cur.execute("""
                INSERT INTO computers (barcode, case_number, cage_number, status, location, exam_appeal, notes, scan_time)
                VALUES (%s, %s, %s, %s, %s, %s, %s, NOW())
            """, (data['barcode'], data['case_number'], data['cage_number'], data['status'], data['location'], data['exam_appeal'], data['notes']))
            conn.commit()
            flash("מחשב נוסף בהצלחה!", "success")
            return redirect(url_for('computers'))
        except Exception as e:
            conn.rollback()
            flash(f"שגיאה בהוספת מחשב: {e}", "danger")
        finally:
            cur.close()
            conn.close()
            
    return render_template('computer_form.html', action='add', computer=None)

@app.route('/edit-computer/<int:cid>', methods=['GET', 'POST'])
@login_required
def edit_computer(cid):
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    try:
        if request.method == 'POST':
            data = request.form
            # Record history before update
            cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
            old_val = cur.fetchone()
            
            cur.execute("""
                UPDATE computers 
                SET case_number=%s, cage_number=%s, status=%s, location=%s, exam_appeal=%s, notes=%s
                WHERE id=%s
            """, (data['case_number'], data['cage_number'], data['status'], data['location'], data['exam_appeal'], data['notes'], cid))
            
            # History log
            cur.execute("""
                INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                VALUES (%s, %s, 'Manual Edit', %s, %s)
            """, (cid, session.get('username'), old_val, dict(data)))
            
            conn.commit()
            flash("פרטי המחשב עודכנו!", "success")
            return redirect(url_for('computers'))
            
        cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
        computer = cur.fetchone()
        return render_template('computer_form.html', action='edit', computer=computer)
    finally:
        cur.close()
        conn.close()

@app.route('/delete-computer/<int:cid>', methods=['POST'])
@login_required
def delete_computer(cid):
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM computers WHERE id = %s", (cid,))
        conn.commit()
        flash("המחשב נמחק מהמערכת", "warning")
    except Exception as e:
        flash(f"שגיאה במחיקה: {e}", "danger")
    finally:
        cur.close()
        conn.close()
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
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT * FROM computers WHERE exam_appeal IS NOT NULL AND exam_appeal != ''")
    computers = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('exam.html', computers=computers)

@app.route('/history')
@login_required
def history_page():
    conn = get_db_connection()
    if not conn: return redirect(url_for('dashboard'))
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
    conn.close()
    return render_template('history.html', history=history)

@app.route('/api/process-scan', methods=['POST'])
@login_required
def process_scan():
    barcode = request.json.get('barcode')
    if not barcode:
        return {"error": "No barcode provided"}, 400

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    try:
        # Check if computer exists
        cur.execute("SELECT * FROM computers WHERE barcode = %s", (barcode,))
        computer = cur.fetchone()
        
        if computer:
            # Update scan_time (last seen)
            cur.execute("UPDATE computers SET scan_time = NOW() WHERE id = %s", (computer['id'],))
            conn.commit()
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
            return {"exists": False, "computer": new_computer}
            
    except Exception as e:
        print(f"Error in process_scan: {e}")
        return {"error": str(e)}, 500
    finally:
        cur.close()
        conn.close()

@app.route('/api/update-computer', methods=['POST'])
@login_required
def api_update_computer():
    data = request.json
    cid = data.get('id')
    if not cid: return {"error": "No ID provided"}, 400

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    try:
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
            """, (cid, session.get('username'), old_val, data))
            
            conn.commit()
            return {"success": True}
        return {"success": False, "message": "No fields to update"}
        
    except Exception as e:
        print(f"Error in api_update_computer: {e}")
        return {"error": str(e)}, 500
    finally:
        cur.close()
        conn.close()

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    print("\n🚀 URI SYSTEM IS LIVE!")
    print("🌐 Link: http://127.0.0.1:5000\n")
    app.run(debug=True, port=5000)