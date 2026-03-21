import psycopg2
import os
import uuid
from dotenv import load_dotenv

load_dotenv()

def init_examinees_db():
    conn_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not conn_url:
        print("DATABASE_URL not found in .env")
        return

    try:
        conn = psycopg2.connect(conn_url)
        cur = conn.cursor()

        print("Updating 'examinees' table...")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS examinees (
                id SERIAL PRIMARY KEY,
                full_name TEXT NOT NULL,
                id_number TEXT,
                extra_time TEXT,
                username TEXT,
                password TEXT,
                exam_name TEXT,
                exam_date TEXT,
                classroom TEXT,
                seat_number TEXT,
                row_number TEXT,
                laptop_number TEXT,
                is_present INT DEFAULT 0,
                laptop_status TEXT DEFAULT 'תקין',
                notes TEXT,
                scan_time TIMESTAMP,
                scanner_technician TEXT,
                token TEXT UNIQUE NOT NULL,
                created_at TIMESTAMP DEFAULT NOW()
            );
        """)

        # Check if we have sample data, if not add some
        cur.execute("SELECT COUNT(*) FROM examinees")
        if cur.fetchone()[0] == 0:
            print("Adding sample examinee data...")
            sample_data = [
                ('חנה בונימוביץ', '308595057', '—', '74765', '29613', 'משרד הבריאות - מומחיות', '08.09.2020', 'כיתה 1', '1', '2', '222', str(uuid.uuid4())[:8]),
                ('ישראל ישראלי', '123456789', '25%', '88990', '12345', 'משרד החינוך - בגרות', '16.03.2026', 'כיתה 3', '15', '4', '101', str(uuid.uuid4())[:8]),
            ]
            for item in sample_data:
                cur.execute("""
                    INSERT INTO examinees (full_name, id_number, extra_time, username, password, exam_name, exam_date, classroom, seat_number, row_number, laptop_number, token)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, item)

        conn.commit()
        print("Examinees table initialized successfully!")
        
        cur.close()
        conn.close()

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    init_examinees_db()
