import psycopg2, os
from dotenv import load_dotenv
load_dotenv()
conn = psycopg2.connect(os.getenv('DATABASE_URL'), connect_timeout=10)
cur = conn.cursor()
cur.execute("SELECT barcode, scanned_by, last_technician FROM computers WHERE barcode='3104'")
r = cur.fetchone()
if r:
    print('barcode:', r[0])
    print('scanned_by repr:', repr(r[1]))
    print('last_technician repr:', repr(r[2]))
else:
    print("לא נמצא מחשב 3104")

# גם בדוק היסטוריה
cur.execute("SELECT technician, change_type FROM inventory_history WHERE computer_id IN (SELECT id FROM computers WHERE barcode='3104') ORDER BY id DESC LIMIT 3")
for row in cur.fetchall():
    print(f"history technician: {repr(row[0])}, change: {repr(row[1])}")

conn.close()
