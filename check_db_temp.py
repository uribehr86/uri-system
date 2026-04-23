import psycopg2
import os
from dotenv import load_dotenv
from psycopg2.extras import RealDictCursor

load_dotenv()

def check_dell_2018():
    db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not db_url:
        print("No DATABASE_URL found")
        return

    try:
        conn = psycopg2.connect(db_url)
        cur = conn.cursor(cursor_factory=RealDictCursor)
        
        # Search for Dell or 2018 in notes or barcode case-insensitive
        cur.execute("SELECT id, barcode, notes FROM computers WHERE notes ILIKE '%Dell%' OR notes ILIKE '%2018%' OR barcode ILIKE '%Dell%' LIMIT 10")
        rows = cur.fetchall()
        print(f"Found {len(rows)} matching computers:")
        for r in rows:
            print(f"- {r['barcode']}: {r['notes']}")
            
        cur.execute("SELECT COUNT(*) FROM computers WHERE notes ILIKE '%Dell%' OR notes ILIKE '%2018%'")
        count = cur.fetchone()['count']
        print(f"\nTotal Dell/2018 computers: {count}")
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

if __name__ == "__main__":
    check_dell_2018()
