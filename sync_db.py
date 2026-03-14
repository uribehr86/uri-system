import sqlite3
import psycopg2
import os
from dotenv import load_dotenv
from psycopg2.extras import RealDictCursor

load_dotenv()

def sync():
    RENDER_URL = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not RENDER_URL:
        print("❌ DATABASE_URL not found in .env")
        return

    try:
        print("🔗 מתחבר ל-Render (מושך נתונים)...")
        pg_conn = psycopg2.connect(RENDER_URL)
        pg_cur = pg_conn.cursor(cursor_factory=RealDictCursor)

        # משיכת נתונים מכל הטבלאות
        pg_cur.execute("SELECT * FROM computers")
        computers = pg_cur.fetchall()
        pg_cur.execute("SELECT * FROM users")
        users = pg_cur.fetchall()
        pg_cur.execute("SELECT * FROM cages")
        cages = pg_cur.fetchall()

        # חיבור ל-DB המקומי
        sl_conn = sqlite3.connect('system_data.db')
        sl_cur = sl_conn.cursor()

        # יצירת טבלאות עם סכמה מלאה יותר
        sl_cur.execute("DROP TABLE IF EXISTS computers")
        sl_cur.execute("""
            CREATE TABLE computers (
                id INTEGER PRIMARY KEY,
                barcode TEXT,
                case_number TEXT,
                cage_number TEXT,
                cage_name TEXT,
                status TEXT,
                location TEXT,
                exam_appeal TEXT,
                notes TEXT,
                scan_time TEXT
            )
        """)
        
        sl_cur.execute("DROP TABLE IF EXISTS users")
        sl_cur.execute("CREATE TABLE users (id INTEGER PRIMARY KEY, username TEXT UNIQUE, password TEXT, role TEXT, timestamp TEXT)")

        # הכנסת נתונים
        for c in computers:
            sl_cur.execute("""
                INSERT INTO computers (id, barcode, case_number, cage_number, cage_name, status, location, exam_appeal, notes, scan_time) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (c['id'], c['barcode'], c.get('case_number'), c.get('cage_number'), c.get('cage_name'), c['status'], c.get('location'), c.get('exam_appeal'), c.get('notes'), str(c.get('scan_time'))))
        
        for u in users:
            sl_cur.execute("INSERT INTO users (id, username, password, role, timestamp) VALUES (?, ?, ?, ?, ?)", (u['id'], u['username'], u['password'], u['role'], str(u.get('timestamp'))))

        sl_conn.commit()
        print(f"✅ הצלחנו! {len(computers)} מחשבים ו-{len(users)} משתמשים עברו ל-system_data.db")
    except Exception as e:
        print(f"❌ שגיאה: {e}")
    finally:
        if 'pg_conn' in locals(): pg_conn.close()
        if 'sl_conn' in locals(): sl_conn.close()

if __name__ == "__main__":
    sync()

if __name__ == "__main__":
    sync()