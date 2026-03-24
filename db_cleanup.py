import psycopg2
import os
import re
from dotenv import load_dotenv

def cleanup_duplicates():
    load_dotenv()
    db_url = os.getenv('DATABASE_URL')
    if not db_url:
        print("DATABASE_URL not found")
        return

    try:
        conn = psycopg2.connect(db_url)
        cur = conn.cursor()

        print("1. Normalizing all barcodes (removing leading zeros)...")
        cur.execute("SELECT id, barcode FROM computers")
        rows = cur.fetchall()
        for cid, barcode in rows:
            if not barcode: continue
            norm = re.sub(r'^0+(?=\d)', '', str(barcode))
            if norm != barcode:
                cur.execute("UPDATE computers SET barcode = %s WHERE id = %s", (norm, cid))
        
        conn.commit()
        print("Normalization complete.")

        print("2. Identifying and merging duplicates...")
        cur.execute("""
            SELECT barcode, COUNT(*) 
            FROM computers 
            GROUP BY barcode 
            HAVING COUNT(*) > 1
        """)
        dups = cur.fetchall()
        
        deleted_count = 0
        for barcode, count in dups:
            print(f"Merging {count} copies of barcode {barcode}...")
            # Fetch all records for this barcode, newest first
            cur.execute("""
                SELECT id, case_number, cage_number, cage_name, status, location, exam_appeal, notes, last_technician, scan_time 
                FROM computers 
                WHERE barcode = %s 
                ORDER BY scan_time DESC NULLS LAST, id DESC
            """, (barcode,))
            records = cur.fetchall()
            
            # The first record is our "target" (the one we keep)
            target = list(records[0])
            target_id = target[0]
            
            # Iterate through the others and pull missing data into the target
            for i in range(1, len(records)):
                source = records[i]
                source_id = source[0]
                
                # Merge logic: if target field is empty but source has it, copy it
                updates = []
                params = []
                # indices: 1-case, 2-cage_num, 3-cage_name, 4-status, 5-loc, 6-exam, 7-notes, 8-tech, 9-time
                for idx, col in enumerate(['case_number', 'cage_number', 'cage_name', 'status', 'location', 'exam_appeal', 'notes', 'last_technician']):
                    source_val = source[idx+1]
                    target_val = target[idx+1]
                    if source_val and not target_val:
                        updates.append(f"{col} = %s")
                        params.append(source_val)
                        target[idx+1] = source_val # update local target copy too
                
                if updates:
                    params.append(target_id)
                    cur.execute(f"UPDATE computers SET {', '.join(updates)} WHERE id = %s", params)
                
                # Delete the source record
                cur.execute("DELETE FROM computers WHERE id = %s", (source_id,))
                deleted_count += 1
                
        conn.commit()
        print(f"Merge complete. Deleted {deleted_count} redundant records.")
        
        # 3. Prevent future duplicates
        try:
            print("3. Adding unique constraint to barcode column...")
            cur.execute("ALTER TABLE computers ADD CONSTRAINT unique_barcode UNIQUE (barcode)")
            conn.commit()
            print("Unique constraint added successfully.")
        except Exception as e:
            print(f"Could not add unique constraint: {e}")
            conn.rollback()

        cur.close()
        conn.close()
        print("Database cleanup finished.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    cleanup_duplicates()
