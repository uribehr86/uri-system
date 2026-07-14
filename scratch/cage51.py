import psycopg2, os
from dotenv import load_dotenv
load_dotenv()
conn = psycopg2.connect(os.getenv('DATABASE_URL'))
cur = conn.cursor()
cur.execute("SELECT barcode, status, location, scanned_by FROM computers WHERE cage_number='51' ORDER BY barcode")
rows = cur.fetchall()
print(f"כלוב 51 — סה\"כ {len(rows)} מחשבים:\n")
for r in rows:
    print(f"  {str(r[0]):<8} | {str(r[1]):<8} | {str(r[2] or '-'):<15} | {str(r[3] or '-')}")
conn.close()
