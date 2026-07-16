"""
sync_attendance_drive.py
========================
סנכרון נוכחות נבחנים לתיקייה ב-Google Drive.

לוגיקה:
1. מחפש/יוצר תיקיית פרויקט תחת PARENT_FOLDER_ID (תיקייה שאתה שיתפת עם ה-SA)
2. מחפש גיליון קיים בתוך תיקיית הפרויקט (שיצרת ידנית)
   → אם נמצא: כותב לתוכו
   → אם לא נמצא: Fallback - כותב ל-Tab בגיליון EXAM_ATTENDANCE_SHEET_ID
3. לא יוצר גיליונות חדשים (SA אין Storage)

"""
import os
import json
import traceback
from dotenv import load_dotenv
load_dotenv()

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import gspread

# קריאה מה-.env
MAIN_FOLDER_ID      = os.getenv('PARENT_FOLDER_ID', '1t0OE1n8Ydav3juiWJys91B70feyaH-Rr')
SERVICE_ACCOUNT_FILE = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', 'service_account.json')
FALLBACK_SHEET_ID   = os.getenv('EXAM_ATTENDANCE_SHEET_ID', '1YWLJA5T8Uq7IGzlzXSA1PPwrdSIPx9eazEcwWXwh3uM')

DEFAULT_HEADERS = ["שם נבחן", "תעודת זהות", "קוד משתמש", "מחשב", "הגיע?", "שעת סריקה", "מיקום", "הערות"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


def _get_clients():
    """מחזיר (gspread_client, drive_service)"""
    sa_json_str = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
    if sa_json_str:
        creds = Credentials.from_service_account_info(json.loads(sa_json_str), scopes=SCOPES)
    elif os.path.exists(SERVICE_ACCOUNT_FILE):
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    else:
        raise FileNotFoundError(f"service account not found: {SERVICE_ACCOUNT_FILE}")
    return gspread.authorize(creds), build('drive', 'v3', credentials=creds)


def _find_or_create_folder(drive, name, parent_id):
    """מחפש תיקייה, יוצר אם לא קיימת"""
    q = (f"mimeType='application/vnd.google-apps.folder' "
         f"and name='{name}' and '{parent_id}' in parents and trashed=false")
    res = drive.files().list(q=q, fields='files(id,name)').execute()
    files = res.get('files', [])
    if files:
        print(f"[Drive] Folder exists: '{name}'")
        return files[0]['id']
    # יצירת תיקייה — לא אוכלת storage
    meta = {'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    folder = drive.files().create(body=meta, fields='id').execute()
    print(f"[Drive] Created folder: '{name}' → {folder['id']}")
    return folder['id']


def _find_sheet_in_folder(drive, folder_id):
    """מחפש גיליון (כלשהו) בתוך תיקייה. מחזיר (sheet_id, sheet_name) או (None, None)"""
    q = (f"mimeType='application/vnd.google-apps.spreadsheet' "
         f"and '{folder_id}' in parents and trashed=false")
    res = drive.files().list(q=q, fields='files(id,name)', pageSize=5).execute()
    files = res.get('files', [])
    if files:
        f = files[0]
        print(f"[Drive] Found sheet in folder: '{f['name']}' → {f['id']}")
        return f['id'], f['name']
    return None, None


def _write_to_sheet(gs, sheet_id, tab_name, examinees, create_tab_if_missing=False):
    """כותב נוכחות לגיליון קיים (לפי sheet_id + tab_name אם קיים)"""
    sh = gs.open_by_key(sheet_id)

    # נסה למצוא tab לפי שם
    try:
        ws = sh.worksheet(tab_name) if tab_name else sh.sheet1
    except gspread.WorksheetNotFound:
        if create_tab_if_missing:
            ws = sh.add_worksheet(title=tab_name, rows=1000, cols=20)
            print(f"[Sheets] Created tab: '{tab_name}'")
        else:
            ws = sh.sheet1  # fallback ל-Sheet1

    ws.clear()
    data = [DEFAULT_HEADERS]
    for r in examinees:
        attended = r.get('is_present', False)
        attended_str = "✅ כן" if attended in (True, 1, 'true', 't') else "❌ לא"
        scan_time = r.get('scan_time')
        scan_time_str = (scan_time.strftime("%d/%m/%Y %H:%M")
                         if scan_time and not isinstance(scan_time, str)
                         else str(scan_time or ''))
        data.append([
            r.get('full_name', ''),
            r.get('id_number', ''),
            r.get('username', ''),
            r.get('laptop_number', ''),
            attended_str,
            scan_time_str,
            r.get('classroom', ''),
            r.get('notes', '')
        ])

    ws.update(data, 'A1')
    last_col = chr(ord('A') + len(DEFAULT_HEADERS) - 1)
    ws.format(f"A1:{last_col}1", {
        "textFormat": {"bold": True},
        "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1.0},
        "horizontalAlignment": "CENTER"
    })
    print(f"[Sheets] Written {len(examinees)} rows to '{ws.title}' in sheet {sheet_id}")
    return sh.id


def sync_exam_to_drive(exam_name, examinees_for_exam):
    """
    מקבל שם בחינה בפורמט 'פרויקט - שם_בחינה' ורשימת נבחנים,
    ומסנכרן לתיקייה/גיליון הנכון ב-Drive.

    מחזיר: (sheet_id, sheet_url) או (None, None) אם נכשל
    """
    try:
        # חילוץ שם פרויקט ושם בחינה
        parts = exam_name.split('-', 1)
        if len(parts) == 2:
            project_name = parts[0].strip()
            exam_subname = parts[1].strip()
        else:
            project_name = "כללי"
            exam_subname = exam_name.strip()

        print(f"[Drive] Syncing: project='{project_name}' exam='{exam_subname}'")

        gs, drive = _get_clients()

        # שלב 1 — תיקיית פרויקט
        folder_id = _find_or_create_folder(drive, project_name, MAIN_FOLDER_ID)

        # שלב 2 — חיפוש גיליון בתיקייה (שיצרת ידנית)
        sheet_id, sheet_name = _find_sheet_in_folder(drive, folder_id)

        if sheet_id:
            # כתוב לגיליון שנמצא — Tab לפי שם הבחינה
            final_id = _write_to_sheet(gs, sheet_id, exam_subname, examinees_for_exam, create_tab_if_missing=True)
            url = f"https://docs.google.com/spreadsheets/d/{final_id}/edit"
            print(f"[Drive] Success! URL: {url}")
            return final_id, url
        else:
            # Fallback: Tab בגיליון הנוכחות הראשי
            print(f"[Drive] No sheet in folder → Fallback to main attendance sheet")
            final_id = _write_to_sheet(gs, FALLBACK_SHEET_ID, exam_subname, examinees_for_exam, create_tab_if_missing=True)
            url = f"https://docs.google.com/spreadsheets/d/{FALLBACK_SHEET_ID}/edit"
            print(f"[Drive] Fallback success! URL: {url}")
            return FALLBACK_SHEET_ID, url

    except Exception as e:
        print(f"[Drive ERROR] {e}")
        traceback.print_exc()
        return None, None


# ── הרצה עצמאית לבדיקה ──────────────────────────────────────────────
if __name__ == '__main__':
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

    test_name = sys.argv[1] if len(sys.argv) > 1 else 'משרד הבריאות - רישוי חשמלאים 16.7.26'
    test_examinees = [
        {'full_name': 'ישראל ישראלי', 'id_number': '123456789', 'username': 'israel1',
         'laptop_number': '5001', 'is_present': True, 'scan_time': None, 'classroom': 'אולם 1', 'notes': ''},
        {'full_name': 'שרה כהן', 'id_number': '987654321', 'username': 'sarah2',
         'laptop_number': '5002', 'is_present': False, 'scan_time': None, 'classroom': 'אולם 1', 'notes': 'לא הגיעה'},
    ]
    sid, url = sync_exam_to_drive(test_name, test_examinees)
    if sid:
        print(f'\nSheet ID: {sid}')
        print(f'Link: {url}')
    else:
        print('\nFailed!')
