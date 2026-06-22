import os
import psycopg2
from dotenv import load_dotenv
import sys

sys.stdout.reconfigure(encoding='utf-8')

load_dotenv()
db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')

def run():
    print("Connecting to database...")
    conn = psycopg2.connect(db_url)
    cur = conn.cursor()
    
    # Let's delete "משרד הבריאות" from projects table
    cur.execute("DELETE FROM projects WHERE name = 'משרד הבריאות'")
    print("Deleted 'משרד הבריאות' from projects table.")
    
    conn.commit()
    
    # Print remaining projects
    cur.execute("SELECT id, name, keywords, sheets_id, drive_url FROM projects")
    rows = cur.fetchall()
    print("\nCurrent projects in database:")
    for r in rows:
        print(f"ID: {r[0]}, Name: {r[1]}, Keywords: {r[2]}, Sheet ID: {r[3]}, Drive URL: {r[4]}")
        
    cur.close()
    conn.close()

if __name__ == "__main__":
    run()
