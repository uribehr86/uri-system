import psycopg2
from psycopg2.extras import RealDictCursor
import os
from dotenv import load_dotenv
load_dotenv()

db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
conn = psycopg2.connect(db_url)
cur = conn.cursor(cursor_factory=RealDictCursor)

cur.execute("""
    SELECT barcode, COUNT(*) as cnt, array_agg(id) as ids, array_agg(cage_number) as cages
    FROM computers
    GROUP BY barcode
    HAVING COUNT(*) > 1
    ORDER BY cnt DESC
""")
rows = cur.fetchall()

if not rows:
    print("RESULT: אין כפילויות! כל ברקוד מופיע פעם אחת בלבד.")
else:
    print(f"RESULT: נמצאו {len(rows)} ברקודים כפולים:")
    for r in rows:
        print(f"  ברקוד={r['barcode']} | כמות={r['cnt']} | IDs={r['ids']} | כלובים={r['cages']}")

cur.close()
conn.close()
