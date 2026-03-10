import os
from flask import Flask, render_template, request, jsonify
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime
import pytz

app = Flask(__name__)

# פונקציה לחיבור לבסיס הנתונים החדש ב-Render
def get_db_connection():
    # הוא מושך את הכתובת מה-Environment Variable שהגדרנו בשלב 2
    conn = psycopg2.connect(os.environ.get('DATABASE_URL'))
    return conn

# יצירת הטבלאות אם הן לא קיימות (חשוב למעבר ל-Render)
def init_db():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS computers (
            id SERIAL PRIMARY KEY,
            barcode TEXT UNIQUE NOT NULL,
            cage_name TEXT,
            status TEXT DEFAULT 'scanned',
            location TEXT,
            notes TEXT,
            scan_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    cur.close()
    conn.close()

@app.route('/')
def index():
    return render_template('checker.html')

@app.route('/scan', methods=['POST'])
def scan():
    data = request.json
    barcode = data.get('barcode')
    cage = data.get('cage_name')
    
    if not barcode:
        return jsonify({"status": "error", "message": "Missing barcode"}), 400

    try:
        conn = get_db_connection()
        cur = conn.cursor()
        
        # שמירת הסריקה
        cur.execute(
            "INSERT INTO computers (barcode, cage_name) VALUES (%s, %s) ON CONFLICT (barcode) DO UPDATE SET cage_name = EXCLUDED.cage_name",
            (barcode, cage)
        )
        
        conn.commit()
        cur.close()
        conn.close()
        return jsonify({"status": "success", "message": f"Barcode {barcode} saved!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/save_scan_full', methods=['POST'])
def save_scan_full():
    data = request.json
    barcode = data.get('barcode')
    cage = data.get('cage')
    location = data.get('location')
    status = data.get('status')
    notes = data.get('notes')
    
    if not barcode or not cage:
        return jsonify({"status": "error", "message": "Missing barcode or cage"}), 400

    try:
        conn = get_db_connection()
        cur = conn.cursor()
        
        # Ensure the table is created with latest columns
        
        cur.execute(
            """INSERT INTO computers (barcode, cage_name, status, location, notes) 
               VALUES (%s, %s, %s, %s, %s) 
               ON CONFLICT (barcode) 
               DO UPDATE SET 
                  cage_name = EXCLUDED.cage_name, 
                  status = EXCLUDED.status,
                  location = EXCLUDED.location,
                  notes = EXCLUDED.notes""",
            (barcode, cage, status, location, notes)
        )
        
        # Get current count for this cage
        cur.execute("SELECT COUNT(*) FROM computers WHERE cage_name = %s", (cage,))
        cage_count = cur.fetchone()[0]
        
        conn.commit()
        cur.close()
        conn.close()
        
        return jsonify({
            "status": "success", 
            "message": f"Barcode {barcode} saved!",
            "cage_count": cage_count
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=10000)