import os
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv

load_dotenv()

db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
if not db_url:
    print("Database URL not found in environment")
    exit(1)

try:
    conn = psycopg2.connect(db_url)
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    cur.execute("SELECT COUNT(*) as total FROM computers")
    total = cur.fetchone()['total']
    print(f"Total computers in DB: {total}")
    
    cur.execute("SELECT status, COUNT(*) as count FROM computers GROUP BY status")
    stats = cur.fetchall()
    print("Stats by status:")
    for row in stats:
        print(f"  {row['status']}: {row['count']}")
        
    cur.close()
    conn.close()
except Exception as e:
    print(f"Error: {e}")
