import sqlite3
import psycopg2
from psycopg2.extras import RealDictCursor

# הכתובת המעודכנת מה-.env שלך
RENDER_URL = "postgresql://uri_system_db_user:VfsC66ho76RaIYZFYgIFZytreG3JaUtc@dpg-d6nhhuv5gffc73bkekmg-a.oregon-postgres.render.com/uri_system_db?sslmode=require"

def sync():
    try:
        print("🔗 מתחבר ל-Render (מושך נתונים ל-uri_system_db_user)...")
        pg_conn = psycopg2.connect(RENDER_URL)
        pg_cur = pg_conn.cursor(cursor_factory=RealDictCursor)

        # משיכת נתונים
        pg_cur.execute("SELECT * FROM computers")
        computers = pg_cur.fetchall()
        pg_cur.execute("SELECT * FROM users")
        users = pg_cur.fetchall()

        # חיבור ל-DB המקומי
        sl_conn = sqlite3.connect('system_data.db')
        sl_cur = sl_conn.cursor()

        # יצירת טבלאות
        sl_cur.execute("DROP TABLE IF EXISTS computers")
        sl_cur.execute("CREATE TABLE computers (barcode TEXT PRIMARY KEY, cage TEXT, status TEXT, exam TEXT, location TEXT, notes TEXT)")
        sl_cur.execute("DROP TABLE IF EXISTS users")
        sl_cur.execute("CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT, role TEXT)")

        # הכנסת נתונים
        for c in computers:
            sl_cur.execute("INSERT INTO computers (barcode, status) VALUES (?, ?)", (c.get('barcode') or c.get('id'), c['status']))
        for u in users:
            sl_cur.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", (u['username'], u['password'], u['role']))

        sl_conn.commit()
        print(f"✅ הצלחנו! {len(computers)} מחשבים עברו ל-system_data.db")
    except Exception as e:
        print(f"❌ שגיאה: {e}")
    finally:
        if 'pg_conn' in locals(): pg_conn.close()
        if 'sl_conn' in locals(): sl_conn.close()

if __name__ == "__main__":
    sync()