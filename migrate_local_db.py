import sqlite3
import shutil
import os
from datetime import datetime

DB_NAME = 'system_data.db'
BACKUP_NAME = f"system_data_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"

def migrate():
    print(f"Starting migration of {DB_NAME}...")
    
    # 1. Backup
    if os.path.exists(DB_NAME):
        shutil.copy2(DB_NAME, BACKUP_NAME)
        print(f"Backup created: {BACKUP_NAME}")
    else:
        print("Error: DB file not found.")
        return

    try:
        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()

        # 2. Get existing columns
        cur.execute("PRAGMA table_info(computers)")
        existing_cols = [row[1] for row in cur.fetchall()]
        print(f"Existing columns: {existing_cols}")

        # 3. Create new table structure
        cur.execute("ALTER TABLE computers RENAME TO computers_old")
        
        cur.execute("""
            CREATE TABLE computers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                barcode TEXT,
                case_number TEXT,
                cage_number TEXT,
                cage_name TEXT,
                status TEXT,
                location TEXT,
                exam_appeal TEXT,
                specs TEXT,
                project TEXT,
                ministry TEXT,
                notes TEXT,
                scan_time TEXT,
                last_technician TEXT
            )
        """)
        
        # 4. Map columns
        # Map: barcode->barcode, cage->cage_number, status->status, exam->exam_appeal, location->location, notes->notes
        mapping = {
            'barcode': 'barcode',
            'cage': 'cage_number',
            'status': 'status',
            'exam': 'exam_appeal',
            'location': 'location',
            'notes': 'notes'
        }
        
        src_cols = []
        dst_cols = []
        
        for src, dst in mapping.items():
            if src in existing_cols:
                src_cols.append(src)
                dst_cols.append(dst)
        
        if src_cols:
            query = f"INSERT INTO computers ({', '.join(dst_cols)}) SELECT {', '.join(src_cols)} FROM computers_old"
            print(f"Executing: {query}")
            cur.execute(query)
            
        # 5. Handle Users table if needed (check if it exists)
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='users'")
        if not cur.fetchone():
            print("Creating users table...")
            cur.execute("CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT, role TEXT, timestamp TEXT)")
            # Add a default user if none exists
            cur.execute("INSERT OR IGNORE INTO users (username, password, role) VALUES ('uri', '1234', 'admin')")

        # 6. Cleanup
        cur.execute("DROP TABLE computers_old")
        conn.commit()
        print("Migration completed successfully!")
        
    except Exception as e:
        print(f"Migration failed: {e}")
        conn.rollback()
    finally:
        conn.close()

if __name__ == "__main__":
    migrate()
