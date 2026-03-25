import gspread
from google.oauth2.service_account import Credentials
import psycopg2
from psycopg2.extras import RealDictCursor
import os
from dotenv import load_dotenv
from datetime import datetime
import json
from utils import format_history, summarize_history

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

def update_worksheet(sh, name, header, rows):
    """עדכון גיליון ספציפי בתוך הקובץ"""
    try:
        # פתיחת הגיליון או יצירתו אם לא קיים
        try:
            worksheet = sh.worksheet(name)
        except gspread.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=name, rows="100", cols="10")
        
        data_to_write = [header]
        data_to_write.extend(rows)

        # עדכון הגיליון
        worksheet.clear()
        # gspread 6.x uses values first, then range
        worksheet.update(data_to_write, 'A1')
        
        # עיצוב כותרת
        worksheet.format("A1:" + chr(ord('A') + len(header) - 1) + "1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1.0}
        })
        return True
    except Exception as e:
        print(f"שגיאה בעדכון גיליון {name}: {e}")
        return False

def sync_inventory_to_sheets():
    """
    מסנכרן את נתוני המערכת לגיליון Google Sheets עם מספר טאבים.
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
        sh = client.open_by_key(spreadsheet_id)

        # 2. חיבור ל-DB
        conn = get_db_connection()
        if not conn:
            return False, "לא ניתן להתחבר למסד הנתונים."
        cur = conn.cursor(cursor_factory=RealDictCursor)

        # --- טאב 1: מלאי כללי ---
        cur.execute("""
            SELECT c.barcode, c.case_number, c.cage_number, c.status, c.location, c.exam_appeal, c.notes, c.scan_time,
                   h.technician, h.change_type, h.old_value, h.new_value
            FROM computers c
            LEFT JOIN (
                SELECT DISTINCT ON (computer_id) computer_id, technician, change_type, old_value, new_value
                FROM inventory_history
                ORDER BY computer_id, timestamp DESC
            ) h ON c.id = h.computer_id
            ORDER BY c.scan_time DESC NULLS LAST
        """)
        inv_rows = cur.fetchall()
        inv_header = ["מחשב", "מספר תיק", "מספר כלוב", "סטטוס", "מיקום", "מבחן/ערעור", "הערות", "טכנאי", "פעולה אחרונה", "נצפה לאחרונה"]
        inv_data = []
        for r in inv_rows:
            # Prepare entry for summarize_history helper
            entry = {
                'old_value': r['old_value'],
                'new_value': r['new_value'],
                'change_type': r['change_type']
            }
            last_action = summarize_history(entry) if r['change_type'] else ''
            
            inv_data.append([
                r['barcode'], 
                r['case_number'] or '', 
                r['cage_number'] or '', 
                r['status'] or '', 
                r['location'] or '', 
                r['exam_appeal'] or '', 
                r['notes'] or '', 
                r['technician'] or '',
                last_action,
                r['scan_time'].strftime("%d/%m/%Y %H:%M") if r['scan_time'] else ''
            ])
        update_worksheet(sh, "מלאי מחשבים", inv_header, inv_data)

        # --- טאב 2: מבחן וערעור ---
        cur.execute("""
            SELECT barcode, case_number, exam_appeal, status, location, scan_time
            FROM computers 
            WHERE exam_appeal IS NOT NULL 
              AND TRIM(exam_appeal) != '' 
              AND LOWER(TRIM(exam_appeal)) != 'none'
            ORDER BY scan_time DESC NULLS LAST
        """)
        exam_rows = cur.fetchall()
        exam_header = ["מחשב", "מספר תיק", "מבחן/ערעור", "סטטוס", "מיקום", "נצפה לאחרונה"]
        exam_data = []
        for r in exam_rows:
            exam_data.append([r['barcode'], r['case_number'] or '', r['exam_appeal'] or '', r['status'] or '', 
                              r['location'] or '', r['scan_time'].strftime("%d/%m/%Y %H:%M") if r['scan_time'] else ''])
        update_worksheet(sh, "מבחן-ערעור", exam_header, exam_data)

        # --- טאב 3: היסטוריית שינויים ---
        cur.execute("""
            SELECT h.timestamp, c.barcode, h.technician, h.change_type, h.old_value, h.new_value
            FROM inventory_history h
            LEFT JOIN computers c ON h.computer_id = c.id
            ORDER BY h.timestamp DESC LIMIT 200
        """)
        hist_rows = cur.fetchall()
        hist_header = ["זמן", "טכנאי", "מחשב", "סוג שינוי", "תיאור פעולה", "לפני", "אחרי"]
        hist_data = []
        for r in hist_rows:
            # Prepare entry for summarize_history helper
            entry = {
                'old_value': r['old_value'],
                'new_value': r['new_value'],
                'change_type': r['change_type']
            }
            
            hist_data.append([
                r['timestamp'].strftime("%d/%m/%Y %H:%M") if r['timestamp'] else '',
                r['technician'] or '',
                r['barcode'] or 'מחשב',
                r['change_type'] or '',
                summarize_history(entry),
                format_history(r['old_value']),
                format_history(r['new_value'])
            ])
        update_worksheet(sh, "היסטוריה", hist_header, hist_data)

        # --- טאב 4: סיכום סטטיסטי ---
        cur.execute("""
            SELECT 
                COUNT(*) as total,
                COUNT(*) FILTER (WHERE status = 'תקין') as ok,
                COUNT(*) FILTER (WHERE status = 'תקול') as faulty,
                COUNT(*) FILTER (WHERE status = 'בתיקון') as repairing,
                COUNT(*) FILTER (WHERE status = 'מאוחסן') as stored
            FROM computers
        """)
        stats = cur.fetchone()
        
        summary_header = ["קטגוריה", "כמות"]
        summary_data = [
            ["סה\"כ מחשבים במערכת", stats['total']],
            ["✅ תקין", stats['ok']],
            ["❌ תקול", stats['faulty']],
            ["🔧 בתיקון", stats['repairing']],
            ["📦 מאוחסן", stats['stored']]
        ]
        
        update_worksheet(sh, "סיכום", summary_header, summary_data)

        cur.close()
        conn.close()

        return True, "הסנכרון הושלם בהצלחה עבור כל הגיליונות!"

    except Exception as e:
        print(f"שגיאת סנכרון כללית: {e}")
        return False, f"שגיאה בתהליך הסנכרון: {str(e)}"

if __name__ == "__main__":
    success, msg = sync_inventory_to_sheets()
    print(msg)
