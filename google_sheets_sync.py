import gspread
from google.oauth2.service_account import Credentials
import psycopg2
from psycopg2.extras import RealDictCursor
import os
from dotenv import load_dotenv
from datetime import datetime

# טעינת משתני סביבה
load_dotenv()

def get_db_connection():
    """חיבור למסד הנתונים PostgreSQL"""
    db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    try:
        return psycopg2.connect(db_url)
    except Exception as e:
        print(f"שגיאה בחיבור ל-DB: {e}")
        return None

def sync_inventory_to_sheets():
    """
    מסנכרן את טבלת ה-computers ממסד הנתונים לגיליון Google Sheets.
    """
    spreadsheet_id = os.getenv('GOOGLE_SHEETS_ID')
    service_account_file = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', 'service_account.json')

    if not spreadsheet_id:
        return False, "GOOGLE_SHEETS_ID לא מוגדר ב-.env"

    if not os.path.exists(service_account_file):
        return False, f"קובץ ההרשאות ({service_account_file}) לא נמצא בתיקיית הפרויקט."

    try:
        # 1. הגדרת גישה ל-API של גוגל
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(service_account_file, scopes=scopes)
        client = gspread.authorize(creds)
        
        # פתיחת הגיליון
        sh = client.open_by_key(spreadsheet_id)
        worksheet = sh.get_worksheet(0) # הגיליון הראשון

        # 2. שליפת נתונים מה-DB
        conn = get_db_connection()
        if not conn:
            return False, "לא ניתן להתחבר למסד הנתונים."
            
        cur = conn.cursor(cursor_factory=RealDictCursor)
        cur.execute("""
            SELECT barcode, case_number, cage_number, status, location, exam_appeal, notes, scan_time
            FROM computers
            ORDER BY scan_time DESC
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()

        # 3. עיצוב הנתונים עבור ה-Sheets (כותרות וערכים)
        header = ["ברקוד", "מספר תיק", "מספר כלוב", "סטטוס", "מיקום", "מבחן/ערעור", "הערות", "נצפה לאחרונה"]
        data_to_write = [header]
        
        for row in rows:
            data_to_write.append([
                row['barcode'],
                row['case_number'] or '',
                row['cage_number'] or '',
                row['status'] or '',
                row['location'] or '',
                row['exam_appeal'] or '',
                row['notes'] or '',
                str(row['scan_time']) if row['scan_time'] else ''
            ])

        # 4. עדכון הגיליון בפעולה אחת
        worksheet.clear()
        worksheet.update('A1', data_to_write)
        
        # עיצוב כותרת (בולד וצבע רקע)
        worksheet.format("A1:H1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1.0}
        })

        return True, f"הסנכרון הושלם! {len(rows)} מחשבים עודכנו בגיליון."

    except Exception as e:
        print(f"שגיאת סנכרון: {e}")
        return False, f"שגיאה בתהליך הסנכרון: {str(e)}"

if __name__ == "__main__":
    success, msg = sync_inventory_to_sheets()
    print(msg)
