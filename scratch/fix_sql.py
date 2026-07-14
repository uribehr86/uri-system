import psycopg2, os
from dotenv import load_dotenv
load_dotenv()
try:
    conn = psycopg2.connect(os.getenv("DATABASE_URL"), connect_timeout=10)
    cur = conn.cursor()

    # תיקון history
    cur.execute("""
        UPDATE inventory_history
        SET technician = convert_from(convert_to(technician, 'WIN1252'), 'UTF8')
        WHERE technician IS NOT NULL AND technician ~ '[^\\x00-\\x7F]'
    """)
    h = cur.rowcount

    # תיקון computers scanned_by
    cur.execute("""
        UPDATE computers
        SET scanned_by = convert_from(convert_to(scanned_by, 'WIN1252'), 'UTF8')
        WHERE scanned_by IS NOT NULL AND scanned_by != '' AND scanned_by ~ '[^\\x00-\\x7F]'
    """)
    c = cur.rowcount

    # תיקון computers last_technician
    cur.execute("""
        UPDATE computers
        SET last_technician = convert_from(convert_to(last_technician, 'WIN1252'), 'UTF8')
        WHERE last_technician IS NOT NULL AND last_technician != '' AND last_technician ~ '[^\\x00-\\x7F]'
    """)
    lt = cur.rowcount

    conn.commit()
    print(f"history: {h} rows, scanned_by: {c} rows, last_technician: {lt} rows")

    cur.execute("SELECT technician FROM inventory_history ORDER BY id DESC LIMIT 2")
    for r in cur.fetchall():
        print(f"  {r[0]}")

    conn.close()
except Exception as e:
    print(f"ERROR: {e}")
