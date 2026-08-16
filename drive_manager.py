"""
drive_manager.py
================
יצירת תיקייה + Google Sheet אוטומטית לכל פרויקט/בחינה.

create_folder_and_sheet_if_not_exists(project_name)
  1. מתחבר ל-Google Drive + Sheets דרך service_account.json
  2. מחפש תיקייה בשם project_name בתוך PARENT_FOLDER_ID
     → אם לא קיימת — יוצר
  3. מחפש קובץ Sheets באותו שם בתוך התיקייה
     → אם לא קיים — יוצר + מוסיף כותרות
  4. מחזיר: sheet_id (מחרוזת) או None אם נכשל
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', write_through=True)

import os
from dotenv import load_dotenv
load_dotenv()

PARENT_FOLDER_ID = os.getenv('PARENT_FOLDER_ID', '18-VtXbYxvT8EqVzJgdAZv54aDnRWhNvM')
SA_FILE          = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', 'service_account.json')

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# כותרות ברירת מחדל — לפי מבנה האקסל
DEFAULT_HEADERS = [
    'שם פרטי', 'שם משפחה', 'ת.ז', 'התאמות', 'סיסמה',
    'שם משתמש', 'גרסה', 'אולם/כיתה', 'טור', 'כסא',
    'מ.מחשב', 'נוכחות', 'שעת סריקה', 'טכנאי'
]


def _get_clients():
    """מחזיר (gspread_client, drive_service)"""
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    import gspread

    creds = Credentials.from_service_account_file(SA_FILE, scopes=SCOPES)
    gs    = gspread.authorize(creds)
    drive = build('drive', 'v3', credentials=creds)
    return gs, drive


def _find_folder(drive, name, parent_id):
    """מחפש תיקייה לפי שם בתוך parent_id. מחזיר folder_id או None."""
    q = (
        f"mimeType='application/vnd.google-apps.folder' "
        f"and name='{name}' "
        f"and '{parent_id}' in parents "
        f"and trashed=false"
    )
    res = drive.files().list(q=q, fields='files(id,name)').execute()
    files = res.get('files', [])
    return files[0]['id'] if files else None


def _create_folder(drive, name, parent_id):
    """יוצר תיקייה ומחזיר folder_id."""
    meta = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = drive.files().create(body=meta, fields='id').execute()
    print(f"[Drive] Created folder: '{name}' → {folder['id']}", flush=True)
    return folder['id']


def _find_sheet(drive, name, parent_id):
    """מחפש Google Sheet לפי שם בתוך parent_id. מחזיר sheet_id או None."""
    q = (
        f"mimeType='application/vnd.google-apps.spreadsheet' "
        f"and name='{name}' "
        f"and '{parent_id}' in parents "
        f"and trashed=false"
    )
    res = drive.files().list(q=q, fields='files(id,name)').execute()
    files = res.get('files', [])
    return files[0]['id'] if files else None


def _create_sheet_with_headers(gs, drive, name, parent_id):
    """יוצר Google Sheet ישירות בתיקייה (ללא שימוש ב-storage), מוסיף כותרות. מחזיר sheet_id."""
    # יצירת Sheets ישירות בתיקייה — Google Docs/Sheets לא תופסים storage
    file_meta = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': [parent_id]
    }
    sheet_file = drive.files().create(body=file_meta, fields='id').execute()
    sheet_id = sheet_file['id']
    print(f"[Drive] Created sheet: '{name}' → {sheet_id}", flush=True)

    # הוספת כותרות
    sh = gs.open_by_key(sheet_id)
    ws = sh.sheet1
    # יצירת נתונים: שורה 1 - כותרת ממוזגת, שורה 2 - כותרות עמודות
    title_row = [name] + [''] * (len(DEFAULT_HEADERS) - 1)
    ws.update(values=[title_row, DEFAULT_HEADERS], range_name='A1')

    # עיצוב
    last_col = chr(ord('A') + len(DEFAULT_HEADERS) - 1)
    
    # מיזוג ועיצוב שורה 1 (כותרת ראשית)
    ws.merge_cells(f'A1:{last_col}1')
    ws.format(f'A1:{last_col}1', {
        'textFormat': {'bold': True, 'fontSize': 14},
        'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}, # אפור בהיר
        'horizontalAlignment': 'CENTER',
        'verticalAlignment': 'MIDDLE'
    })

    # עיצוב שורה 2 (עמודות)
    ws.format(f'A2:{last_col}2', {
        'textFormat': {'bold': True},
        'backgroundColor': {'red': 0.8, 'green': 0.9, 'blue': 1.0},
        'horizontalAlignment': 'CENTER'
    })
    print(f"[Drive] Headers set for '{name}'", flush=True)
    return sheet_id



def create_folder_and_sheet_if_not_exists(project_name):
    """
    הפונקציה הראשית.
    1. מוצא/יוצר תיקייה בשם הפרויקט תחת PARENT_FOLDER_ID
    2. מחפש גיליון בתוך התיקייה:
       - נמצא  → מחזיר את ה-sheet_id שלו
       - לא נמצא → מחזיר (None, project_name) כדי שהקורא ישתמש ב-Tab
    מחזירה: (sheet_id, sheet_name/tab_name)
    """
    import os
    from dotenv import load_dotenv
    load_dotenv()

    MAIN_SHEET_ID = os.getenv('EXAM_ATTENDANCE_SHEET_ID', '1YWLJA5T8Uq7IGzlzXSA1PPwrdSIPx9eazEcwWXwh3uM')

    try:
        gs, drive = _get_clients()

        # שלב 1 — תיקייה ב-Drive
        folder_id = _find_folder(drive, project_name, PARENT_FOLDER_ID)
        if folder_id:
            print(f"[Drive] Folder exists: '{project_name}'", flush=True)
        else:
            folder_id = _create_folder(drive, project_name, PARENT_FOLDER_ID)

        # שלב 2 — חיפוש גיליון בתוך התיקייה (שהמשתמש יצר ידנית)
        # מחפש כל גיליון בתיקייה (לאו דווקא בשם המדויק)
        q = (
            f"mimeType='application/vnd.google-apps.spreadsheet' "
            f"and '{folder_id}' in parents "
            f"and trashed=false"
        )
        res = drive.files().list(q=q, fields='files(id,name)', pageSize=10).execute()
        sheets_in_folder = res.get('files', [])

        if sheets_in_folder:
            # השתמש בגיליון הראשון שנמצא בתיקייה
            found = sheets_in_folder[0]
            print(f"[Drive] Found sheet in folder: '{found['name']}' → {found['id']}", flush=True)
            return found['id'], found['name']
        else:
            # לא נמצא גיליון — Fallback: Tab בגיליון הראשי
            print(f"[Drive] No sheet in folder — falling back to tab in main sheet", flush=True)
            sh = gs.open_by_key(MAIN_SHEET_ID)
            existing_titles = [w.title for w in sh.worksheets()]

            if project_name in existing_titles:
                print(f"[Sheets] Tab already exists: '{project_name}'", flush=True)
            else:
                ws_new = sh.add_worksheet(title=project_name, rows=1000, cols=20)
                title_row = [project_name] + [''] * (len(DEFAULT_HEADERS) - 1)
                ws_new.update(values=[title_row, DEFAULT_HEADERS], range_name='A1')
                
                last_col = chr(ord('A') + len(DEFAULT_HEADERS) - 1)
                
                ws_new.merge_cells(f'A1:{last_col}1')
                ws_new.format(f'A1:{last_col}1', {
                    'textFormat': {'bold': True, 'fontSize': 14},
                    'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment': 'MIDDLE'
                })
                ws_new.format(f'A2:{last_col}2', {
                    'textFormat': {'bold': True},
                    'backgroundColor': {'red': 0.8, 'green': 0.9, 'blue': 1.0},
                    'horizontalAlignment': 'CENTER'
                })
                print(f"[Sheets] Created tab: '{project_name}'", flush=True)

            return MAIN_SHEET_ID, project_name

    except Exception as ex:
        import traceback
        print(f"[Drive ERROR] {ex}", flush=True)
        traceback.print_exc()
        return None, None


# ── הרצה עצמאית לבדיקה ──────────────────────────────────────────────
if __name__ == '__main__':
    import sys
    name = sys.argv[1] if len(sys.argv) > 1 else 'רישוי חשמלאים'
    sid, tab = create_folder_and_sheet_if_not_exists(name)
    if sid:
        print(f'\nSheet ID: {sid}  |  Name: {tab}')
        print(f'Link: https://docs.google.com/spreadsheets/d/{sid}/edit')
    else:
        print('\nFailed!')


