import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def check():
    db_url = os.getenv('DATABASE_URL')
    try:
        conn = psycopg2.connect(db_url)
        cur = conn.cursor()
        cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'computers'")
        cols = [c[0] for c in cur.fetchall()]
        print(f"COLUMNS: {cols}")
        cur.close()
        conn.close()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check()
