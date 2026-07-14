import psycopg2, os
from dotenv import load_dotenv
load_dotenv()
conn = psycopg2.connect(os.getenv('DATABASE_URL'))
cur = conn.cursor()

# בדוק מה עומד להימחק
cur.execute("""
    SELECT barcode FROM computers 
    WHERE cage_number='51' 
    AND (barcode !~ '^6[0-9]+$')
    ORDER BY barcode
""")
to_clear = [r[0] for r in cur.fetchall()]
print(f"מחשבים שיאבדו את כלוב 51 ({len(to_clear)}):")
print(', '.join(str(b) for b in to_clear))

confirm = input("\nלהמשיך? (כן/לא): ")
if confirm.strip() == 'כן':
    cur.execute("""
        UPDATE computers 
        SET cage_number='', cage_name=''
        WHERE cage_number='51' 
        AND (barcode !~ '^6[0-9]+$')
    """)
    updated = cur.rowcount
    conn.commit()
    print(f"✅ עודכנו {updated} מחשבים — כלוב 51 נוקה מהם")
else:
    print("בוטל.")

cur.close()
conn.close()
