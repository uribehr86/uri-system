import psycopg2, os
from dotenv import load_dotenv
load_dotenv()
conn = psycopg2.connect(os.getenv('DATABASE_URL'))
cur = conn.cursor()

cur.execute("SELECT id, technician FROM inventory_history WHERE technician IS NOT NULL LIMIT 1")
row = cur.fetchone()
row_id, tech = row
print(f"repr: {repr(tech)}")
print(f"len: {len(tech)}")

# ניסיון תיקון שונה
for enc in ['latin-1', 'cp1252', 'iso-8859-8']:
    try:
        fixed = tech.encode(enc).decode('utf-8')
        print(f"{enc} → utf8: {fixed}")
    except Exception as e:
        print(f"{enc} נכשל: {e}")

# ניסיון ישיר - קח את ה-bytes הגולמיים
cur.execute("SELECT encode(technician::bytea, 'hex') FROM inventory_history WHERE id=%s", (row_id,))
hex_val = cur.fetchone()[0]
print(f"\nhex bytes: {hex_val[:60]}")
try:
    raw = bytes.fromhex(hex_val)
    print(f"decoded as utf8: {raw.decode('utf-8')}")
except Exception as e:
    print(f"hex decode error: {e}")

conn.close()
