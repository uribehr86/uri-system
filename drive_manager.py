"""
drive_manager.py
================
יצירת מבנה Drive דו-שכבתי לכל מבחן:

    PARENT_FOLDER_ID
      └── <שם המשרד>            ← תיקיית משרד (למשל "משרד הבריאות")
            └── <שם המבחן_תאריך>  ← גיליון לכל מבחן (למשל "רופאים_6.10.26")

get_exam_sheet(exam_title)
  1. מתחבר ל-Google Drive + Sheets דרך service_account.json
  2. מפרק את שם המבחן (exam_naming.parse_exam_title) למשרד/מבחן/תאריך
  3. מוצא/יוצר את תיקיית המשרד תחת PARENT_FOLDER_ID
  4. מוצא/יוצר גיליון בשם המבחן בתוך תיקיית המשרד
  5. מחזיר dict עם sheet_id, worksheet ומידע הפירוק

חשוב: הגיליון מאותר לפי שם מדויק — לא "הגיליון הראשון בתיקייה".
כך שני מבחנים של אותו משרד לא דורסים זה את זה.
"""

import sys, io
import os
import json
from dotenv import load_dotenv
load_dotenv()

from exam_naming import parse_exam_title

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

# מספר השורות שאינן נתונים: [1] כותרת ממוזגת, [2] כותרות עמודות
HEADER_ROWS = 2


def _get_clients():
    """מחזיר (gspread_client, drive_service)"""
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    import gspread

    sa_json_str = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
    if sa_json_str:
        creds = Credentials.from_service_account_info(json.loads(sa_json_str), scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(SA_FILE, scopes=SCOPES)
    gs    = gspread.authorize(creds)
    drive = build('drive', 'v3', credentials=creds)
    return gs, drive


def _escape(name):
    """בריחה לתווים מיוחדים בשאילתת Drive"""
    return str(name).replace('\\', '\\\\').replace("'", "\\'")


def _find_folder(drive, name, parent_id):
    """מחפש תיקייה לפי שם בתוך parent_id. מחזיר folder_id או None."""
    q = (
        f"mimeType='application/vnd.google-apps.folder' "
        f"and name='{_escape(name)}' "
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
    print(f"[Drive] Created folder: '{name}' -> {folder['id']}", flush=True)
    return folder['id']


def get_or_create_folder(drive, name, parent_id):
    """מוצא תיקייה או יוצר אותה. מחזיר folder_id."""
    folder_id = _find_folder(drive, name, parent_id)
    if folder_id:
        print(f"[Drive] Folder exists: '{name}'", flush=True)
        return folder_id
    return _create_folder(drive, name, parent_id)


def _find_sheet(drive, name, parent_id):
    """מחפש Google Sheet בשם מדויק בתוך parent_id. מחזיר sheet_id או None."""
    q = (
        f"mimeType='application/vnd.google-apps.spreadsheet' "
        f"and name='{_escape(name)}' "
        f"and '{parent_id}' in parents "
        f"and trashed=false"
    )
    res = drive.files().list(q=q, fields='files(id,name)').execute()
    files = res.get('files', [])
    return files[0]['id'] if files else None


def apply_header_formatting(ws, title_text, headers=None):
    """שורה 1 — כותרת ממוזגת, שורה 2 — כותרות עמודות."""
    headers = headers or DEFAULT_HEADERS
    title_row = [title_text] + [''] * (len(headers) - 1)
    # gspread 6.x: values/range_name בשמות מפורשים
    ws.update(values=[title_row, headers], range_name='A1', value_input_option='USER_ENTERED')

    last_col = chr(ord('A') + len(headers) - 1)
    try:
        ws.merge_cells(f'A1:{last_col}1')
    except Exception:
        pass  # כבר ממוזג
    ws.format(f'A1:{last_col}1', {
        'textFormat': {'bold': True, 'fontSize': 14},
        'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
        'horizontalAlignment': 'CENTER',
        'verticalAlignment': 'MIDDLE'
    })
    ws.format(f'A2:{last_col}2', {
        'textFormat': {'bold': True},
        'backgroundColor': {'red': 0.8, 'green': 0.9, 'blue': 1.0},
        'horizontalAlignment': 'CENTER'
    })


def _create_sheet_with_headers(gs, drive, name, parent_id, title_text=None):
    """יוצר Google Sheet בתוך התיקייה + כותרות. מחזיר sheet_id."""
    # Google Sheets לא תופס storage של ה-Service Account
    file_meta = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': [parent_id]
    }
    sheet_file = drive.files().create(body=file_meta, fields='id').execute()
    sheet_id = sheet_file['id']
    print(f"[Drive] Created sheet: '{name}' -> {sheet_id}", flush=True)

    sh = gs.open_by_key(sheet_id)
    apply_header_formatting(sh.sheet1, title_text or name)
    print(f"[Drive] Headers set for '{name}'", flush=True)
    return sheet_id


def get_exam_sheet(exam_title, filename=None, create=True):
    """
    הפונקציה הראשית — מחזירה את הגיליון של מבחן מסוים.

    exam_title: הכותרת הגולמית (שורה 1 באקסל / parts[0] של ה-QR)
    filename:   שם קובץ, משמש כ-fallback אם אין כותרת
    create:     True — יוצר תיקייה/גיליון אם חסרים. False — רק מאתר.

    מחזיר dict:
      {'sheet_id', 'worksheet', 'spreadsheet', 'folder_id', 'created',
       'office', 'exam_base', 'date', 'hall', 'sheet_name', 'exam_name', 'url'}
    או None אם נכשל / לא נמצא.
    """
    info = parse_exam_title(exam_title, filename)
    try:
        gs, drive = _get_clients()

        # שלב 1 — תיקיית המשרד תחת התיקייה הראשית
        office_folder_id = _find_folder(drive, info['office'], PARENT_FOLDER_ID)
        if not office_folder_id:
            if not create:
                print(f"[Drive] Office folder not found: '{info['office']}'", flush=True)
                return None
            office_folder_id = _create_folder(drive, info['office'], PARENT_FOLDER_ID)

        # שלב 2 — גיליון המבחן בתוך תיקיית המשרד (שם מדויק!)
        sheet_id = _find_sheet(drive, info['sheet_name'], office_folder_id)
        created = False
        if not sheet_id:
            if not create:
                print(f"[Drive] Exam sheet not found: '{info['sheet_name']}'", flush=True)
                return None
            sheet_id = _create_sheet_with_headers(
                gs, drive, info['sheet_name'], office_folder_id,
                title_text=info['exam_name']
            )
            created = True
        else:
            print(f"[Drive] Exam sheet exists: '{info['sheet_name']}' -> {sheet_id}", flush=True)

        sh = gs.open_by_key(sheet_id)
        result = dict(info)
        result.update({
            'sheet_id': sheet_id,
            'spreadsheet': sh,
            'worksheet': sh.sheet1,
            'folder_id': office_folder_id,
            'created': created,
            'url': f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
        })
        return result

    except Exception as ex:
        import traceback
        print(f"[Drive ERROR] {ex}", flush=True)
        traceback.print_exc()
        return None


def find_header_row(all_values):
    """
    מאתר את אינדקס שורת הכותרות (0-based). מחזיר -1 אם הגיליון ריק.
    תומך גם בגיליון עם שורת כותרת ממוזגת לפני הכותרות וגם בלעדיה.
    """
    markers = ['ת.ז', 'תעודת', 'שם', 'נוכחות', 'מחשב', 'סיסמ', 'משתמש']
    for idx, row in enumerate(all_values[:10]):
        row_text = ' '.join(str(c).strip() for c in row)
        if sum(1 for m in markers if m in row_text) >= 2:
            return idx
    return -1


def _col_index(headers, keywords):
    for k in keywords:
        for i, h in enumerate(headers):
            if k in str(h):
                return i
    return None


def build_examinee_row(rec, headers=None):
    """ממפה רשומת נבחן לשורה לפי סדר DEFAULT_HEADERS."""
    headers = headers or DEFAULT_HEADERS
    return [
        rec.get('full_name', ''),    # שם פרטי (שם מלא)
        rec.get('last_name', ''),    # שם משפחה
        rec.get('id_number', ''),    # ת.ז
        rec.get('adaptations', ''),  # התאמות
        rec.get('password', ''),     # סיסמה
        rec.get('username', ''),     # שם משתמש
        rec.get('version', ''),      # גרסה
        rec.get('hall', ''),         # אולם/כיתה
        rec.get('row', ''),          # טור
        rec.get('seat', ''),         # כסא
        rec.get('computer', ''),     # מ.מחשב   ← נסרק
        rec.get('is_present', ''),   # נוכחות   ← נסרק
        rec.get('scan_time', ''),    # שעת סריקה ← נסרק
        rec.get('technician', ''),   # טכנאי    ← נסרק
    ][:len(headers)]


# עמודות שהסריקה ממלאת — ייבוא חוזר לעולם לא דורס אותן
SCAN_COLUMNS = {10, 11, 12, 13}


def merge_examinees_into_sheet(ws, records, title_text=None):
    """
    ממזג נבחנים לגיליון קיים — בלי למחוק כלום.

    מבחן מחולק לאולמות מגיע בכמה קבצי אקסל, ולכן ייבוא חוזר חייב להוסיף
    ולעדכן, לא לאפס. לכל רשומה:
      - ת.ז שכבר בגיליון  → מעדכן רק את עמודות הזיהוי (לא נוכחות/מחשב/שעה/טכנאי)
      - ת.ז חדשה          → נוספת בסוף

    מחזיר (added, updated).
    """
    all_values = ws.get_all_values()
    header_idx = find_header_row(all_values)

    if header_idx == -1:
        # גיליון ריק (או בלי כותרות) — מתקין כותרות ומתחיל מאפס
        apply_header_formatting(ws, title_text or 'נבחנים')
        all_values = ws.get_all_values()
        header_idx = find_header_row(all_values)
        if header_idx == -1:
            header_idx = HEADER_ROWS - 1

    headers = [str(h).strip() for h in all_values[header_idx]]
    id_col = _col_index(headers, ['ת.ז', 'תעודת', 'id'])
    if id_col is None:
        id_col = 2  # מיקום ת.ז בפריסת ברירת המחדל

    # מיפוי ת.ז → אינדקס שורה בגיליון (1-based)
    existing = {}
    for offset, row in enumerate(all_values[header_idx + 1:], start=header_idx + 2):
        if id_col < len(row):
            key = str(row[id_col]).strip()
            if key:
                existing.setdefault(key, offset)

    updates = []
    to_append = []
    updated = 0

    for rec in records:
        values = build_examinee_row(rec, headers)
        key = str(rec.get('id_number', '')).strip()
        sheet_row = existing.get(key) if key else None

        if sheet_row:
            current = all_values[sheet_row - 1]
            for idx, val in enumerate(values):
                if idx in SCAN_COLUMNS or not val:
                    continue  # לא נוגעים בנתוני סריקה, ולא מוחקים בערך ריק
                old = current[idx].strip() if idx < len(current) else ''
                if old != str(val):
                    updates.append({
                        'range': f'{_a1_col(idx)}{sheet_row}',
                        'values': [[val]],
                    })
            updated += 1
        else:
            to_append.append(values)

    if updates:
        ws.batch_update(updates, value_input_option='USER_ENTERED')
    if to_append:
        ws.append_rows(to_append, value_input_option='USER_ENTERED',
                       table_range=f'A{header_idx + 1}')

    print(f"[Sheets] Merged into '{ws.title}': {len(to_append)} added, {updated} updated", flush=True)
    return len(to_append), updated


def _a1_col(idx):
    """0 → A, 25 → Z, 26 → AA"""
    letters = ''
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(ord('A') + rem) + letters
    return letters


def create_folder_and_sheet_if_not_exists(project_name):
    """
    תאימות לאחור — מחזיר (sheet_id, sheet_name) או (None, None).
    קוד חדש צריך להשתמש ב-get_exam_sheet.
    """
    res = get_exam_sheet(project_name)
    if not res:
        return None, None
    return res['sheet_id'], res['sheet_name']


# ── הרצה עצמאית לבדיקה ──────────────────────────────────────────────
if __name__ == '__main__':
    # עטיפת stdout רק בהרצה ישירה — לא בייבוא, כדי לא לשבור את הלוג של השרת
    if hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', write_through=True)

    name = sys.argv[1] if len(sys.argv) > 1 else 'משרד הבריאות - רופאים 6.10.26'
    res = get_exam_sheet(name)
    if res:
        print(f"\nOffice folder : {res['office']}")
        print(f"Exam sheet    : {res['sheet_name']}")
        print(f"Sheet ID      : {res['sheet_id']}")
        print(f"Link          : {res['url']}")
    else:
        print('\nFailed!')
