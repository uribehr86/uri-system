import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")

conn = psycopg2.connect(DATABASE_URL)
cur = conn.cursor()

# קבל מחשבים עם כלוב + מספר מחשב
cur.execute("""
    SELECT cage_number, barcode
    FROM computers
    WHERE cage_number IS NOT NULL AND cage_number != ''
    ORDER BY cage_number, barcode
""")
rows = cur.fetchall()

from collections import defaultdict
cages = defaultdict(list)
for cage, barcode in rows:
    cages[cage].append(barcode)


print(f"{'כלוב':<8} {'סה\"כ':>6} {'מקסימום':>9} {'עודף':>6}")
print("-" * 35)
overloaded = []
for cage, computers in sorted(cages.items(), key=lambda x: x[0]):
    total = len(computers)
    # בדוק אם יש מחשבים 7000-7280 בכלוב
    has_7k = any(c.isdigit() and 7000 <= int(c) <= 7280 for c in computers)
    limit = 70 if has_7k else 50
    exceed = total - limit
    if exceed > 0:
        overloaded.append((cage, total, limit, exceed, computers))
        print(f"❌ {cage:<6} {total:>6} / {limit:<6}  עודף: +{exceed}")
    else:
        print(f"✅ {cage:<6} {total:>6} / {limit:<6}")

print()
print(f"סה\"כ כלובים עם עודף מחשבים: {len(overloaded)}")
for cage, total, limit, exceed, comps in overloaded:
    extras = comps[limit:]
    print(f"\n⚠️  כלוב {cage}: {total} מחשבים (מקסימום {limit})")
    print(f"   מחשבים עודפים: {', '.join(str(c) for c in extras)}")

cur.close()
conn.close()
