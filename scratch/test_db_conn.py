import os
import psycopg2
from dotenv import load_dotenv

load_dotenv()

db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
print("DB URL:", db_url)

try:
    conn = psycopg2.connect(db_url)
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
except Exception as e:
    print("Postgres master list error (expected as it's not SQLite):", e)

try:
    conn = psycopg2.connect(db_url)
    cur = conn.cursor()
    cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public'")
    tables = cur.fetchall()
    print("Tables:", tables)
    
    # Check projects table
    cur.execute("SELECT * FROM projects")
    rows = cur.fetchall()
    print("Projects:")
    for r in rows:
        print(r)
        
    cur.close()
    conn.close()
except Exception as e:
    print("Error:", e)
