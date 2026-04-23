import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()
conn = psycopg2.connect(os.getenv('RENDER_DB_URL'))
cur = conn.cursor()
try:
    cur.execute("ALTER TABLE computers ADD COLUMN project VARCHAR(255)")
    conn.commit()
    print("Successfully added 'project' column to PostgreSQL database.")
except Exception as e:
    conn.rollback()
    print(f"Error adding column: {e}")
conn.close()
