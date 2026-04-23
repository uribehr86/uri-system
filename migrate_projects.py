import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def migrate():
    # Use RENDER_DB_URL first as it's often the direct internal/external one
    db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not db_url:
        print("DATABASE_URL not found")
        return
    
    db_url = db_url.strip().strip('"').strip("'")
    
    print(f"Connecting to DB...")
    try:
        # Some environments prefer passing the connection string as 'dsn'
        conn = psycopg2.connect(dsn=db_url)
        cur = conn.cursor()
        print("Connected.")

        cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'computers'")
        cols = [c[0] for c in cur.fetchall()]
        
        if 'project' not in cols:
            cur.execute("ALTER TABLE computers ADD COLUMN project TEXT")
            print("Added project.")
        if 'ministry' not in cols:
            cur.execute("ALTER TABLE computers ADD COLUMN ministry TEXT")
            print("Added ministry.")
        
        conn.commit()
        print("Migration Success.")
        cur.close()
        conn.close()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    migrate()
