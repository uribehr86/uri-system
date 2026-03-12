import os
import psycopg2
from dotenv import load_dotenv

load_dotenv()

db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
conn = psycopg2.connect(db_url)
cur = conn.cursor()

cur.execute("SELECT DISTINCT cage_number FROM computers WHERE cage_number IS NOT NULL AND cage_number != ''")
cages = [r[0] for r in cur.fetchall()]

print("Cages in DB:", cages)

cur.close()
conn.close()
