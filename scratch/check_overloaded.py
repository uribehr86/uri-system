import psycopg2, os
from dotenv import load_dotenv
load_dotenv()
conn = psycopg2.connect(os.getenv('DATABASE_URL'))
cur = conn.cursor()

for cage in ['17', '46', '47', '49']:
    cur.execute("SELECT barcode, location FROM computers WHERE cage_number=%s ORDER BY barcode", (cage,))
    rows = cur.fetchall()
    print(f"--- כלוב {cage}: {len(rows)} מחשבים ---")
    for r in rows:
        print(f"  {str(r[0]):<8} | {str(r[1] or '-')}")
    print()

conn.close()
