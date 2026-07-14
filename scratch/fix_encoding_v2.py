import psycopg2, os
from dotenv import load_dotenv
load_dotenv()

def fix_encoding(text):
    """Convert garbled latin1/cp1252 misread string back to proper Hebrew."""
    if not text:
        return text
    try:
        result = bytearray()
        for char in text:
            cp = ord(char)
            if cp <= 0xFF:
                result.append(cp)
            else:
                try:
                    result.extend(char.encode('cp1252'))
                except:
                    result.append(ord('?'))
        decoded = result.decode('utf-8')
        # וודא שיש עברית בתוצאה
        if any('\u0590' <= c <= '\u05ff' for c in decoded):
            return decoded
        return text
    except:
        return text

conn = psycopg2.connect(os.getenv('DATABASE_URL'), connect_timeout=10)
cur = conn.cursor()
fixed_total = 0

# תיקון computers.last_technician
cur.execute("SELECT id, last_technician FROM computers WHERE last_technician IS NOT NULL AND last_technician != ''")
for row_id, val in cur.fetchall():
    fixed = fix_encoding(val)
    if fixed != val:
        cur.execute("UPDATE computers SET last_technician=%s WHERE id=%s", (fixed, row_id))
        fixed_total += 1

# תיקון inventory_history.technician
cur.execute("SELECT id, technician FROM inventory_history WHERE technician IS NOT NULL")
for row_id, val in cur.fetchall():
    fixed = fix_encoding(val)
    if fixed != val:
        cur.execute("UPDATE inventory_history SET technician=%s WHERE id=%s", (fixed, row_id))
        fixed_total += 1

conn.commit()
print(f"תוקנו {fixed_total} שורות")

# בדיקה
cur.execute("SELECT last_technician FROM computers WHERE barcode='3104'")
r = cur.fetchone()
print(f"3104 last_technician: {r[0] if r else 'not found'}")

cur.execute("SELECT technician FROM inventory_history ORDER BY id DESC LIMIT 2")
for r in cur.fetchall():
    print(f"history: {r[0]}")

conn.close()
