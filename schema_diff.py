import sqlite3
import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def get_sqlite_schema():
    conn = sqlite3.connect('system_data.db')
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [r[0] for r in cur.fetchall()]
    schema = {}
    for table in tables:
        cur.execute(f"PRAGMA table_info({table})")
        schema[table] = [row[1] for row in cur.fetchall()]
    conn.close()
    return schema

def get_postgres_schema():
    conn = psycopg2.connect(os.getenv('RENDER_DB_URL'))
    cur = conn.cursor()
    cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public'")
    tables = [r[0] for r in cur.fetchall()]
    schema = {}
    for table in tables:
        cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name=%s", (table,))
        schema[table] = [row[0] for row in cur.fetchall()]
    conn.close()
    return schema

sqlite_schema = get_sqlite_schema()
postgres_schema = get_postgres_schema()

print("--- SQLITE SCHEMA ---")
for t, cols in sqlite_schema.items():
    print(f"{t}: {cols}")

print("\n--- POSTGRES SCHEMA ---")
for t, cols in postgres_schema.items():
    print(f"{t}: {cols}")

print("\n--- DISCREPANCIES ---")
for t in sqlite_schema:
    if t not in postgres_schema:
        print(f"Table '{t}' missing in Postgres!")
    else:
        missing_in_pg = set(sqlite_schema[t]) - set(postgres_schema[t])
        missing_in_sqlite = set(postgres_schema[t]) - set(sqlite_schema[t])
        if missing_in_pg:
            print(f"Col(s) missing in Postgres '{t}': {missing_in_pg}")
        if missing_in_sqlite:
            print(f"Col(s) missing in SQLite '{t}': {missing_in_sqlite}")
