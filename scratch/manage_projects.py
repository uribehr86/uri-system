import os
import psycopg2
from dotenv import load_dotenv
import sys

# Reconfigure stdout for Hebrew encoding support on Windows terminal
sys.stdout.reconfigure(encoding='utf-8')

load_dotenv()
db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')

def add_project(cur, name, keywords, sheets_id, drive_url):
    cur.execute("SELECT id FROM projects WHERE name = %s", (name,))
    existing = cur.fetchone()
    if existing:
        print(f"Project '{name}' exists (ID: {existing[0]}). Updating sheets_id and drive_url...")
        cur.execute(
            "UPDATE projects SET keywords = %s, sheets_id = %s, drive_url = %s WHERE id = %s",
            (keywords, sheets_id, drive_url, existing[0])
        )
    else:
        print(f"Project '{name}' does not exist. Inserting new record...")
        cur.execute(
            "INSERT INTO projects (name, keywords, sheets_id, drive_url) VALUES (%s, %s, %s, %s)",
            (name, keywords, sheets_id, drive_url)
        )

def run():
    print("Connecting to database...")
    conn = psycopg2.connect(db_url)
    cur = conn.cursor()
    
    sheets_id = "1gs3X9KctkQsju_gNpSjyTeeGesksgWKwYVs8S5aX2pQ"
    drive_url = "https://docs.google.com/spreadsheets/d/1gs3X9KctkQsju_gNpSjyTeeGesksgWKwYVs8S5aX2pQ"
    
    # Add "משרד הבריאות"
    add_project(cur, "משרד הבריאות", "משרד הבריאות", sheets_id, drive_url)
    
    # Add "חשמל"
    add_project(cur, "חשמל", "חשמל", sheets_id, drive_url)
    
    conn.commit()
    print("Changes committed successfully!")
    
    # Print projects again
    cur.execute("SELECT id, name, keywords, sheets_id, drive_url FROM projects")
    rows = cur.fetchall()
    print("\nUpdated projects in database:")
    for r in rows:
        print(f"ID: {r[0]}, Name: {r[1]}, Keywords: {r[2]}, Sheet ID: {r[3]}, Drive URL: {r[4]}")
        
    cur.close()
    conn.close()

if __name__ == "__main__":
    run()
