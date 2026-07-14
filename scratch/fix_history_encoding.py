import psycopg2, os
from dotenv import load_dotenv
load_dotenv()
conn = psycopg2.connect(os.getenv('DATABASE_URL'))
cur = conn.cursor()

cur.execute("SELECT id, technician FROM inventory_history WHERE technician IS NOT NULL")
rows = cur.fetchall()
fixed = 0
for row_id, tech in rows:
    try:
        fixed_text = tech.encode('cp1252').decode('utf-8')
        if fixed_text != tech:
            cur.execute("UPDATE inventory_history SET technician=%s WHERE id=%s", (fixed_text, row_id))
            fixed += 1
    except (UnicodeDecodeError, UnicodeEncodeError):
        pass

conn.commit()
print(f"תוקנו {fixed} שורות")

cur.execute("SELECT technician FROM inventory_history ORDER BY id DESC LIMIT 3")
for r in cur.fetchall():
    print(f"  {r[0]}")

conn.close()
