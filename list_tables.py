import sys, io, sqlite3
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
conn = sqlite3.connect(r'c:\uri system scan\uri-system\uri_system.db')
tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
print("Tables:", [t[0] for t in tables])
for t in tables:
    try:
        rows = conn.execute(f"SELECT * FROM {t[0]} LIMIT 3").fetchall()
        print(f"\n{t[0]}:", rows[:2])
    except: pass
conn.close()
