import os
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv

load_dotenv()

db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
if not db_url:
    print("Database URL not found in environment")
    exit(1)

try:
    conn = psycopg2.connect(db_url)
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    test_barcodes = ['DEBUG-123-URI', 'TEST-BARCODE-123']
    print(f"Checking for test barcodes: {test_barcodes}")
    
    for barcode in test_barcodes:
        cur.execute("SELECT * FROM computers WHERE barcode = %s", (barcode,))
        row = cur.fetchone()
        if row:
            print(f"Found test record: {row}")
        else:
            print(f"No record found for {barcode}")
            
    cur.close()
    conn.close()
except Exception as e:
    print(f"Error: {e}")
