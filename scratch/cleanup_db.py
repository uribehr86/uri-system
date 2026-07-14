import psycopg2

url = "postgresql://uri_system_db_user:VfsC66ho76RaIYZFYgIFZytreG3JaUtc@dpg-d6nhhuv5gffc73bkekmg-a.oregon-postgres.render.com/uri_system_db?sslmode=require"
try:
    conn = psycopg2.connect(url)
    cur = conn.cursor()
    cur.execute("UPDATE computers SET exam_appeal = NULL WHERE LOWER(TRIM(exam_appeal)) = 'none'")
    print("Deleted/Updated rows:", cur.rowcount)
    conn.commit()
except Exception as e:
    print("Error:", e)
