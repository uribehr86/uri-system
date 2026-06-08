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

import sqlite3

def get_db_connection():
    """חיבור למסד הנתונים PostgreSQL עם Fallback ל-SQLite מקומי"""
    db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    try:
        # ניסיון חיבור לענן
        conn = psycopg2.connect(db_url)
        return conn, False # False = NOT SQLite
    except Exception as e:
        print(f"שגיאה בחיבור לענן (מנסה מקומי): {e}")
        try:
            # ניסיון חיבור מקומי
            conn = sqlite3.connect('system_data.db')
            conn.row_factory = sqlite3.Row
            return conn, True # True = IS SQLite
        except Exception as e2:
            print(f"שגיאה קריטית בחיבור למסד נתונים: {e2}")
            return None, False

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
        res = get_db_connection()
        if not res or not res[0]:
            return False, "לא ניתן להתחבר למסד הנתונים."
        conn, is_sqlite = res
        
        cur_factory = None if is_sqlite else RealDictCursor
        cur = conn.cursor() if is_sqlite else conn.cursor(cursor_factory=RealDictCursor)

        def safe_execute(query, params=None):
            if is_sqlite:
                query = query.replace('%s', '?')
                # FILTER (WHERE ...) is not supported in basic SQLite
                if "FILTER (WHERE" in query:
                    # Specific rewrite for the stats query
                    query = """
                        SELECT 
                            COUNT(*) as total,
                            SUM(CASE WHEN status = 'תקין' THEN 1 ELSE 0 END) as ok,
                            SUM(CASE WHEN status = 'תקול' THEN 1 ELSE 0 END) as faulty,
                            SUM(CASE WHEN status = 'בתיקון' THEN 1 ELSE 0 END) as repairing,
                            SUM(CASE WHEN status = 'מאוחסן' THEN 1 ELSE 0 END) as stored
                        FROM computers
                    """
            if params:
                cur.execute(query, params)
            else:
                cur.execute(query)

        # --- טאב 1: מלאי מחשבים ---
        # שליפת נתוני המלאי
        q_inv = """
            SELECT barcode, case_number, cage_number, status, location, specs, project, exam_appeal, notes, scan_time,
                   COALESCE(sheets_delete_request, FALSE) as sheets_delete_request
            FROM computers 
            ORDER BY scan_time DESC
        """
        safe_execute(q_inv)
        inv_rows = cur.fetchall()
        inv_header = ["מחשב", "מספר תיק", "מספר כלוב", "סטטוס", "מיקום", "מפרט", "פרויקט", "מבחן/ערעור", "הערות", "נצפה לאחרונה"]
        
        inv_data = [inv_header]
        # שמור את אינדקסי השורות שמסומנות למחיקה (1-based, כולל כותרת)
        delete_flagged_rows = []
        for idx, r in enumerate(inv_rows):
            is_flagged = r['sheets_delete_request'] if not is_sqlite else bool(r[10])
            if is_flagged:
                delete_flagged_rows.append(idx + 2)  # +2: כותרת בשורה 1, נתונים מ-2
            inv_data.append([
                r['barcode'], r['case_number'] or '', r['cage_number'] or '', r['status'] or '', 
                r['location'] or '', r['specs'] or '', r['project'] or '', 
                r['exam_appeal'] or '', r['notes'] or '', 
                r['scan_time'].strftime("%d/%m/%Y %H:%M") if r['scan_time'] and not isinstance(r['scan_time'], str) else str(r['scan_time'] or '')
            ])
            
        # חישוב סטטיסטיקות
        q_stats = """
            SELECT 
                COUNT(*) as total,
                COUNT(*) FILTER (WHERE status = 'תקין') as ok,
                COUNT(*) FILTER (WHERE status = 'תקול') as faulty,
                COUNT(*) FILTER (WHERE status = 'בתיקון') as repairing,
                COUNT(*) FILTER (WHERE status = 'מאוחסן') as stored
            FROM computers
        """
        safe_execute(q_stats)
        stats = cur.fetchone()
        
        summary_data = [
            ["--- סיכום מלאי ---", ""],
            ["סה\"כ מחשבים במערכת", stats['total']],
            ["✅ תקין", stats['ok']],
            ["❌ תקול", stats['faulty']],
            ["🔧 בתיקון", stats['repairing']],
            ["📦 מאוחסן", stats['stored']]
        ]

        try:
            worksheet = sh.worksheet("מלאי מחשבים")
        except gspread.WorksheetNotFound:
            worksheet = sh.add_worksheet(title="מלאי מחשבים", rows="1000", cols="15")
        
        worksheet.clear()
        worksheet.update(inv_data, 'A1')
        
        # עיצוב כותרות
        worksheet.format("A1:J1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1.0}
        })

        # צביעת שורות שמסומנות למחיקה בירוק זוהר
        for row_num in delete_flagged_rows:
            try:
                worksheet.format(f"A{row_num}:J{row_num}", {
                    "backgroundColor": {"red": 0.0, "green": 1.0, "blue": 0.39},
                    "textFormat": {"bold": True, "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}
                })
            except Exception:
                pass  # אם העיצוב נכשל — לא נעצור את כל הסנכרון

        # --- טאב תקולים ---
        q_faulty = """
            SELECT barcode, case_number, cage_number, location, specs, project, notes, scan_time, last_technician
            FROM computers
            WHERE status = 'תקול'
            ORDER BY scan_time DESC
        """
        safe_execute(q_faulty)
        faulty_rows = cur.fetchall()
        faulty_header = ["מחשב", "מספר תיק", "כלוב", "מיקום", "מפרט", "פרויקט", "הערות", "נצפה לאחרונה", "טכנאי אחרון"]
        faulty_data = []
        for r in faulty_rows:
            ts = r['scan_time']
            ts_str = ts.strftime("%d/%m/%Y %H:%M") if ts and not isinstance(ts, str) else str(ts or '')
            faulty_data.append([
                r['barcode'],
                r['case_number'] or '',
                r['cage_number'] or '',
                r['location'] or '',
                r['specs'] or '',
                r['project'] or '',
                r['notes'] or '',
                ts_str,
                r['last_technician'] or ''
            ])
        update_worksheet(sh, "תקולים", faulty_header, faulty_data)

        # --- טאבים לפי פרויקט ---
        safe_execute("SELECT DISTINCT project FROM computers WHERE project IS NOT NULL AND TRIM(project) != ''")
        projects = [row['project'] for row in cur.fetchall()]
        
        for project_name in projects:
            q_proj = """
                SELECT barcode, case_number, cage_number, status, location, specs, notes, scan_time
                FROM computers 
                WHERE project = %s
                ORDER BY scan_time DESC
            """
            safe_execute(q_proj, (project_name,))
            p_rows = cur.fetchall()
            p_header = ["מחשב", "מספר תיק", "מספר כלוב", "סטטוס", "מיקום", "מפרט", "הערות", "נצפה לאחרונה"]
            p_data = []
            for r in p_rows:
                p_data.append([r['barcode'], r['case_number'] or '', r['cage_number'] or '', r['status'] or '', 
                               r['location'] or '', r['specs'] or '', r['notes'] or '',
                               r['scan_time'].strftime("%d/%m/%Y %H:%M") if r['scan_time'] and not isinstance(r['scan_time'], str) else str(r['scan_time'] or '')])
            
            # Create/Update worksheet for this project
            update_worksheet(sh, f"פרויקט-{project_name}", p_header, p_data)

        # --- טאב 2: מבחן וערעור ---
        q_exam = """
            SELECT barcode, case_number, exam_appeal, status, location, scan_time
            FROM computers 
            WHERE exam_appeal IS NOT NULL 
              AND TRIM(exam_appeal) != '' 
              AND LOWER(TRIM(exam_appeal)) != 'none'
            ORDER BY scan_time DESC
        """
        safe_execute(q_exam)
        exam_rows = cur.fetchall()
        exam_header = ["מחשב", "מספר תיק", "מבחן/ערעור", "סטטוס", "מיקום", "נצפה לאחרונה"]
        exam_data = []
        for r in exam_rows:
            exam_data.append([r['barcode'], r['case_number'] or '', r['exam_appeal'] or '', r['status'] or '', 
                              r['location'] or '', r['scan_time'].strftime("%d/%m/%Y %H:%M") if r['scan_time'] and not isinstance(r['scan_time'], str) else str(r['scan_time'] or '')])
        update_worksheet(sh, "מבחן-ערעור", exam_header, exam_data)

        # --- טאב 3: היסטוריית שינויים ---
        q_hist = """
            SELECT h.timestamp, c.barcode, h.technician, h.change_type, h.old_value, h.new_value
            FROM inventory_history h
            LEFT JOIN computers c ON h.computer_id = c.id
            ORDER BY h.timestamp DESC LIMIT 200
        """
        safe_execute(q_hist)
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
            
            ts = r['timestamp']
            if ts and not isinstance(ts, str):
                ts_str = ts.strftime("%d/%m/%Y %H:%M")
            else:
                ts_str = str(ts or '')

            hist_data.append([
                ts_str,
                r['technician'] or '',
                r['barcode'] or 'מחשב',
                r['change_type'] or '',
                summarize_history(entry),
                format_history(r['old_value']),
                format_history(r['new_value'])
            ])
        update_worksheet(sh, "היסטוריה", hist_header, hist_data)

        # --- טאב 4: נוכחות נבחנים ---
        try:
            safe_execute("SELECT full_name, id_number, username, laptop_number, exam_name, classroom, is_present, scan_time, notes FROM examinees ORDER BY exam_name, full_name")
            exam_rows = cur.fetchall()
            exam_header = ["שם נבחן", "ת.ז.", "קוד משתמש", "מחשב", "שם בחינה", "מיקום", "הגיע?", "שעת הגעה", "התאמות"]
            exam_data = []
            for r in exam_rows:
                attended = r['is_present']
                attended_str = "✅ כן" if attended in (True, 1, 'true', 't') else "❌ לא"
                attend_time = r['scan_time']
                if attend_time and not isinstance(attend_time, str):
                    attend_time_str = attend_time.strftime("%d/%m/%Y %H:%M")
                else:
                    attend_time_str = str(attend_time or '')
                exam_data.append([r['full_name'], r['id_number'], r['username'] or '', r['laptop_number'] or '',
                                   r['exam_name'] or '', r['classroom'] or '', attended_str, attend_time_str, r['notes'] or ''])
            update_worksheet(sh, "נוכחות נבחנים", exam_header, exam_data)
        except Exception as e:
            print(f"לא ניתן לסנכרן נוכחות נבחנים: {e}")

        cur.close()
        conn.close()

        return True, "הסנכרון הושלם בהצלחה עבור כל הגיליונות!"

    except Exception as e:
        print(f"שגיאת סנכרון כללית: {e}")
        return False, f"שגיאה בתהליך הסנכרון: {str(e)}"

def import_from_sheets():
    """
    ייבוא מגוגל שיטס → מסד נתונים.
    - מעדכן שדות שהשתנו (סטטוס, מיקום, כלוב, הערות, מפרט, פרויקט, מבחן)
    - אם שורה נמחקה מהגיליון — מסמן sheets_delete_request=True (לא מוחק!)
    - רק admin יכול למחוק לצמיתות
    מחזיר: (success, message, stats_dict)
    """
    spreadsheet_id = os.getenv('GOOGLE_SHEETS_ID')
    service_account_file = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', 'service_account.json')

    if not spreadsheet_id:
        return False, "GOOGLE_SHEETS_ID לא מוגדר ב-.env", {}
    if not os.path.exists(service_account_file):
        return False, f"קובץ ההרשאות ({service_account_file}) לא נמצא.", {}

    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(service_account_file, scopes=scopes)
        client = gspread.authorize(creds)
        sh = client.open_by_key(spreadsheet_id)

        # קריאת גיליון מלאי מחשבים
        try:
            worksheet = sh.worksheet("מלאי מחשבים")
        except gspread.WorksheetNotFound:
            return False, "גיליון 'מלאי מחשבים' לא נמצא בספרדשיט.", {}

        rows = worksheet.get_all_values()
        if len(rows) < 2:
            return False, "הגיליון ריק.", {}

        # כותרות: מחשב, מספר תיק, מספר כלוב, סטטוס, מיקום, מפרט, פרויקט, מבחן/ערעור, הערות, נצפה לאחרונה
        # אינדקסים: 0      1           2            3       4       5      6         7               8       9
        sheet_barcodes = set()
        sheet_rows_by_barcode = {}
        for row in rows[1:]:  # דלג על כותרת
            if not row or not row[0].strip():
                continue
            bc = row[0].strip()
            sheet_barcodes.add(bc)
            sheet_rows_by_barcode[bc] = row

        # חיבור למסד
        conn, is_sqlite = get_db_connection()
        if not conn:
            return False, "לא ניתן להתחבר למסד הנתונים.", {}

        cur = conn.cursor() if is_sqlite else conn.cursor(cursor_factory=RealDictCursor)
        q = "SELECT id, barcode, case_number, cage_number, status, location, specs, project, exam_appeal, notes FROM computers"
        cur.execute(q)
        db_rows = cur.fetchall()

        stats = {"updated": 0, "delete_flagged": 0, "delete_cleared": 0, "skipped": 0}

        EDITABLE_FIELDS = {
            "case_number": 1,
            "cage_number": 2,
            "status": 3,
            "location": 4,
            "specs": 5,
            "project": 6,
            "exam_appeal": 7,
            "notes": 8,
        }

        for db_row in db_rows:
            bc = db_row['barcode'] if not is_sqlite else db_row[1]
            db_id = db_row['id'] if not is_sqlite else db_row[0]

            if bc not in sheet_barcodes:
                # שורה נמחקה מהגיליון — סמן כבקשת מחיקה
                upd = "UPDATE computers SET sheets_delete_request = TRUE WHERE id = %s"
                if is_sqlite:
                    upd = upd.replace('%s', '?').replace('TRUE', '1')
                cur.execute(upd, (db_id,))
                stats["delete_flagged"] += 1
            else:
                sheet_row = sheet_rows_by_barcode[bc]
                updates = []
                params = []
                for field, col_idx in EDITABLE_FIELDS.items():
                    if col_idx >= len(sheet_row):
                        continue
                    sheet_val = sheet_row[col_idx].strip() if sheet_row[col_idx] else ''
                    db_val = str(db_row[field] or '') if not is_sqlite else str(db_row[list(EDITABLE_FIELDS.keys()).index(field) + 2] or '')
                    if sheet_val != db_val:
                        updates.append(f"{field} = %s")
                        params.append(sheet_val if sheet_val else None)

                # אם המחשב היה מסומן למחיקה אבל חזר לגיליון — נקה את הסימון
                clear_flag_q = "UPDATE computers SET sheets_delete_request = FALSE WHERE id = %s AND sheets_delete_request = TRUE"
                if is_sqlite:
                    clear_flag_q = clear_flag_q.replace('%s', '?').replace('FALSE', '0').replace('TRUE', '1')
                cur.execute(clear_flag_q, (db_id,))
                if cur.rowcount > 0:
                    stats["delete_cleared"] += 1

                if updates:
                    q_upd = f"UPDATE computers SET {', '.join(updates)} WHERE id = %s"
                    if is_sqlite:
                        q_upd = q_upd.replace('%s', '?')
                    params.append(db_id)
                    cur.execute(q_upd, params)
                    stats["updated"] += 1
                else:
                    stats["skipped"] += 1

        conn.commit()
        cur.close()
        conn.close()

        msg = (
            f"ייבוא הושלם! "
            f"עודכנו: {stats['updated']} | "
            f"מסומנים למחיקה: {stats['delete_flagged']} | "
            f"בוטלו סימוני מחיקה: {stats['delete_cleared']}"
        )
        return True, msg, stats

    except Exception as e:
        print(f"שגיאת ייבוא כללית: {e}")
        return False, f"שגיאה בייבוא: {str(e)}", {}


if __name__ == "__main__":
    success, msg = sync_inventory_to_sheets()
    print(msg)
