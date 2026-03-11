import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def update_schema():
    conn_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not conn_url:
        print("❌ RENDER_DB_URL or DATABASE_URL not found in .env")
        return

    try:
        conn = psycopg2.connect(conn_url)
        cur = conn.cursor()

        print("🔄 Updating 'computers' table columns...")
        # Add columns if they don't exist
        columns = [
            ("case_number", "TEXT"),
            ("exam_appeal", "TEXT"),
            ("cage_number", "TEXT"), # Based on user request "מספר כלוב"
            ("notes", "TEXT")
        ]
        
        for col_name, col_type in columns:
            cur.execute(f"""
                DO $$ 
                BEGIN 
                    IF NOT EXISTS (SELECT 1 FROM information_schema.columns 
                                   WHERE table_name='computers' AND column_name='{col_name}') THEN 
                        ALTER TABLE computers ADD COLUMN {col_name} {col_type}; 
                    END IF; 
                END $$;
            """)

        print("➕ Creating 'inventory_history' table...")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS inventory_history (
                id SERIAL PRIMARY KEY,
                computer_id INT,
                technician TEXT,
                change_type TEXT,
                old_value JSONB,
                new_value JSONB,
                timestamp TIMESTAMP DEFAULT NOW()
            );
        """)

        conn.commit()
        print("✅ Database schema updated successfully!")
        
        cur.close()
        conn.close()

    except Exception as e:
        print(f"❌ Error updating database: {e}")

if __name__ == "__main__":
    update_schema()
