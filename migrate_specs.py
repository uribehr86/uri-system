import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def migrate():
    db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not db_url:
        print("No DATABASE_URL found")
        return

    try:
        conn = psycopg2.connect(db_url)
        cur = conn.cursor()
        
        # Add specs column to computers table
        print("Adding 'specs' column to computers table...")
        cur.execute("ALTER TABLE computers ADD COLUMN IF NOT EXISTS specs TEXT;")
        
        conn.commit()
        print("Database migration completed successfully.")
        
    except Exception as e:
        print(f"Error during migration: {e}")
        conn.rollback()
    finally:
        if 'conn' in locals():
            conn.close()

if __name__ == "__main__":
    migrate()
