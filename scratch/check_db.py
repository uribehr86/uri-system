import psycopg2
from collections import Counter

url = "postgresql://uri_system_db_user:VfsC66ho76RaIYZFYgIFZytreG3JaUtc@dpg-d6nhhuv5gffc73bkekmg-a.oregon-postgres.render.com/uri_system_db?sslmode=require"
try:
    conn = psycopg2.connect(url)
    cur = conn.cursor()
    cur.execute("SELECT id, exam_appeal FROM computers WHERE exam_appeal IS NOT NULL AND TRIM(exam_appeal) != ''")
    rows = cur.fetchall()
    counts = Counter([r[1] for r in rows])
    
    with open("scratch/results.txt", "w", encoding="utf-8") as f:
        f.write(f"TOTAL MATCHING: {len(rows)}\n")
        f.write("MOST COMMON VALUES:\n")
        for val, count in counts.most_common():
            f.write(f"{val}: {count}\n")
except Exception as e:
    print("Error:", e)
