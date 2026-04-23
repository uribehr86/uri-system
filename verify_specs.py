import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def verify_specs():
    db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not db_url:
        print("No DATABASE_URL found")
        return

    try:
        conn = psycopg2.connect(db_url)
        cur = conn.cursor()
        
        # Insert a test computer with specs
        test_barcode = "TEST-SPEC-123"
        print(f"Inserting test computer {test_barcode} with specs...")
        
        # Delete if exists
        cur.execute("DELETE FROM computers WHERE barcode = %s", (test_barcode,))
        
        cur.execute("""
            INSERT INTO computers (barcode, status, specs, scan_time, last_technician)
            VALUES (%s, 'תקין', '16GB RAM, TEST', NOW(), 'System Test')
            RETURNING id;
        """, (test_barcode,))
        
        cid = cur.fetchone()[0]
        print(f"Inserted with ID: {cid}")
        
        # Verify
        cur.execute("SELECT specs FROM computers WHERE id = %s", (cid,))
        specs = cur.fetchone()[0]
        print(f"Verified specs: {specs}")
        
        # Cleanup
        cur.execute("DELETE FROM computers WHERE id = %s", (cid,))
        conn.commit()
        print("Test completed successfully.")
        
    except Exception as e:
        print(f"Error during verification: {e}")
        if 'conn' in locals():
            conn.rollback()
    finally:
        if 'conn' in locals():
            conn.close()

if __name__ == "__main__":
    verify_specs()
