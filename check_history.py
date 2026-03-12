import os
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv
import json

load_dotenv()

db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
if not db_url:
    print("Database URL not found in environment")
    exit(1)

try:
    conn = psycopg2.connect(db_url)
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    print("Checking recently deleted records in history...")
    cur.execute("""
        SELECT h.*
        FROM inventory_history h
        WHERE h.change_type = 'Manual Edit' OR h.change_type = 'Batch Update'
        ORDER BY h.timestamp DESC
        LIMIT 10
    """)
    rows = cur.fetchall()
    for row in rows:
        print(f"Timestamp: {row['timestamp']}, Change Type: {row['change_type']}")
        print(f"Old Value: {row['old_value']}")
        print(f"New Value: {row['new_value']}")
        print("-" * 20)
            
    cur.close()
    conn.close()
except Exception as e:
    print(f"Error: {e}")
