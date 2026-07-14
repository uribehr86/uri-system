import psycopg2, os
from dotenv import load_dotenv
load_dotenv()
conn = psycopg2.connect(os.getenv('DATABASE_URL'))
cur = conn.cursor()

def try_fix(text):
    """Try multiple encodings to fix garbled Hebrew."""
    if not text:
        return text, False
    for enc in ['utf-8', 'latin-1', 'cp1252', 'cp1255']:
        try:
            fixed = text.encode(enc).decode('utf-8')
            if fixed != text and any('\u0590' <= c <= '\u05ff' for c in fixed):
                return fixed, True
        except (UnicodeDecodeError, UnicodeEncodeError):
            pass
    return text, False

total_fixed = 0

# תיקון inventory_history.technician
print("מתקן inventory_history.technician...")
cur.execute("SELECT id, technician FROM inventory_history WHERE technician IS NOT NULL")
for row_id, tech in cur.fetchall():
    fixed, changed = try_fix(tech)
    if changed:
        cur.execute("UPDATE inventory_history SET technician=%s WHERE id=%s", (fixed, row_id))
        total_fixed += 1

# תיקון inventory_history.change_type
print("מתקן inventory_history.change_type...")
cur.execute("SELECT id, change_type FROM inventory_history WHERE change_type IS NOT NULL")
for row_id, val in cur.fetchall():
    fixed, changed = try_fix(val)
    if changed:
        cur.execute("UPDATE inventory_history SET change_type=%s WHERE id=%s", (fixed, row_id))
        total_fixed += 1

# תיקון computers.scanned_by
print("מתקן computers.scanned_by...")
cur.execute("SELECT id, scanned_by FROM computers WHERE scanned_by IS NOT NULL AND scanned_by != ''")
for row_id, val in cur.fetchall():
    fixed, changed = try_fix(val)
    if changed:
        cur.execute("UPDATE computers SET scanned_by=%s WHERE id=%s", (fixed, row_id))
        total_fixed += 1

# תיקון computers.last_technician
print("מתקן computers.last_technician...")
cur.execute("SELECT id, last_technician FROM computers WHERE last_technician IS NOT NULL AND last_technician != ''")
for row_id, val in cur.fetchall():
    fixed, changed = try_fix(val)
    if changed:
        cur.execute("UPDATE computers SET last_technician=%s WHERE id=%s", (fixed, row_id))
        total_fixed += 1

conn.commit()
print(f"\n✅ תוקנו {total_fixed} שורות בסה\"כ")

# בדיקה
cur.execute("SELECT technician FROM inventory_history ORDER BY id DESC LIMIT 2")
print("\nהיסטוריה אחרי תיקון:")
for r in cur.fetchall():
    print(f"  {r[0]}")

cur.execute("SELECT scanned_by FROM computers WHERE scanned_by IS NOT NULL AND scanned_by != '' LIMIT 3")
print("מחשבים scanned_by:")
for r in cur.fetchall():
    print(f"  {r[0]}")

conn.close()
