import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def init_users():
    conn_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not conn_url:
        print("DATABASE_URL not found")
        return

    try:
        conn = psycopg2.connect(conn_url)
        cur = conn.cursor()

        print("Ensuring 'users' table exists...")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT DEFAULT 'technician',
                timestamp TIMESTAMP DEFAULT NOW()
            );
        """)

        # Add default admin if not exists
        cur.execute("SELECT * FROM users WHERE username = 'admin_uri'")
        if not cur.fetchone():
            print("Adding default admin_uri...")
            cur.execute("INSERT INTO users (username, password, role) VALUES ('admin_uri', 'uri*', 'admin')")
        
        conn.commit()
        print("Users initialization complete (no data was overwritten).")
        cur.close()
        conn.close()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    init_users()
