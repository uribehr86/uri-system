import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def init_db():
    conn_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not conn_url:
        print("RENDER_DB_URL or DATABASE_URL not found")
        return

    try:
        conn = psycopg2.connect(conn_url)
        cur = conn.cursor()

        print("Creating/Updating 'computers' table...")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS computers (
                id SERIAL PRIMARY KEY,
                barcode TEXT UNIQUE NOT NULL,
                case_number TEXT,
                cage_name TEXT,
                cage_number TEXT,
                status TEXT DEFAULT 'תקין',
                location TEXT DEFAULT 'מחסן',
                exam_appeal TEXT,
                notes TEXT,
                scan_time TIMESTAMP DEFAULT NOW()
            );
        """)

        print("Creating 'inventory_history' table...")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS inventory_history (
                id SERIAL PRIMARY KEY,
                computer_id INT REFERENCES computers(id),
                technician TEXT,
                change_type TEXT,
                old_value JSONB,
                new_value JSONB,
                timestamp TIMESTAMP DEFAULT NOW()
            );
        """)

        conn.commit()
        print("Database initialized successfully!")
        cur.close()
        conn.close()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    init_db()
