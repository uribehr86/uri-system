from flask_app import app, get_db_connection, get_safe_cursor
with app.app_context():
    conn = get_db_connection()
    if conn:
        cur = get_safe_cursor(conn)
        # Get all unique cage numbers
        cur.execute("SELECT DISTINCT cage_number FROM computers WHERE cage_number IS NOT NULL AND cage_number != ''")
        cages = [row['cage_number'] for row in cur.fetchall()]
        print("Unique cage numbers in computers table:", cages)
        
        # Check specifically for 4 and 04
        for c in ['4', '04', 'c-4', 'c-04']:
            cur.execute("SELECT COUNT(*) as count FROM computers WHERE cage_number = %s", (c,))
            count = cur.fetchone()['count']
            print(f"Computers in cage '{c}': {count}")
            
        cur.close()
        conn.close()
    else:
        print("No DB conn")
