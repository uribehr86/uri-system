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

# ברירת המחדל היא ה-ID של תיקיית "מערכת בחינות" עצמה (לא התיקייה
# הכללית שמכילה אותה) — כדי שגם בלי PARENT_FOLDER_ID מוגדר ב-Render,
# תיקיות המשרדים ייווצרו בפנים ולא יתפזרו החוצה
PARENT_FOLDER_ID = os.getenv('PARENT_FOLDER_ID', '1bmoF9oe2O6hB4v2MCJEtWRWI7nnO8llv')
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


def get_exam_sheet(exam_title, filename=None, create=True, raise_errors=False):
    """
    הפונקציה הראשית — מחזירה את הגיליון של מבחן מסוים.

    exam_title: הכותרת הגולמית (שורה 1 באקסל / parts[0] של ה-QR)
    filename:   שם קובץ, משמש כ-fallback אם אין כותרת
    create:     True — יוצר תיקייה/גיליון אם חסרים. False — רק מאתר.
    raise_errors: False (ברירת מחדל) — כשל מוחזר כ-None, כמו תמיד.
                  True — כשל אמיתי מול Drive (לא "לא נמצא") נזרק החוצה,
                  כדי שהקורא (למשל מסך הסריקה) יוכל להציג את הסיבה
                  האמיתית למשתמש במקום "נכשל" גנרי.

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
        if raise_errors:
            raise
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


# שדה → מילות מפתח לזיהוי העמודה בגיליון, לפי סדר DEFAULT_HEADERS
FIELD_KEYWORDS = [
    ('full_name',   ['שם פרטי', 'שם נבחן', 'שם', 'name']),
    ('last_name',   ['שם משפחה', 'משפחה', 'family']),
    ('id_number',   ['ת.ז', 'תעודת', 'id']),
    ('adaptations', ['התאמות', 'הערות', 'notes']),
    ('password',    ['סיסמה', 'סיסמא', 'סיסמ', 'password']),
    ('username',    ['שם משתמש', 'משתמש', 'קוד', 'username', 'user']),
    ('version',     ['גרסה', 'בחינה', 'exam']),
    ('hall',        ['אולם', 'כיתה', 'מיקום', 'hall', 'location']),
    ('row',         ['טור', 'עמודה']),
    ('seat',        ['כסא', 'כיסא', 'מושב', 'seat']),
    ('computer',    ['מ.מחשב', 'מחשב', 'computer']),
    ('is_present',  ['נוכחות', 'הגיע', 'attendance']),
    ('scan_time',   ['שעת', 'זמן', 'time', 'scan']),
    ('technician',  ['טכנאי', 'technician']),
    ('pc_status',   ['תקין', 'סטטוס', 'status', 'valid']),
]

# השדות שהסריקה ממלאת — ייבוא רשימת נבחנים לעולם לא דורס אותם
SCAN_FIELDS = {'computer', 'is_present', 'scan_time', 'technician', 'pc_status'}


def map_columns(headers):
    """
    ממפה שדה → אינדקס עמודה לפי הכותרות בפועל.
    עמודה נתפסת פעם אחת בלבד, כדי ש'שם פרטי' ו'שם משפחה' לא ייפלו לאותו מקום.
    """
    headers = [str(h).strip() for h in headers]
    mapping = {}
    taken = set()
    # סבב ראשון: התאמה מדויקת. סבב שני: הכלה חלקית.
    for exact in (True, False):
        for field, keywords in FIELD_KEYWORDS:
            if field in mapping:
                continue
            for kw in keywords:
                for i, h in enumerate(headers):
                    if i in taken:
                        continue
                    hit = (h == kw) if exact else (kw in h)
                    if hit:
                        mapping[field] = i
                        taken.add(i)
                        break
                if field in mapping:
                    break
    return mapping


def build_examinee_row(rec, headers=None):
    """
    בונה שורה לפי הכותרות בפועל של הגיליון — לא לפי מיקום קבוע,
    כדי שגיליון בפריסה שונה לא יקבל נתונים בעמודות הלא נכונות.
    """
    if not headers:
        headers = DEFAULT_HEADERS
    mapping = map_columns(headers)
    row = [''] * max(len(headers), max(mapping.values(), default=0) + 1)
    for field, idx in mapping.items():
        row[idx] = rec.get(field, '')
    return row


def merge_examinees_into_sheet(ws, records, title_text=None, include_scan_columns=False):
    """
    ממזג נבחנים לגיליון קיים — בלי למחוק כלום.

    מבחן מחולק לאולמות מגיע בכמה קבצי אקסל, ולכן ייבוא חוזר חייב להוסיף
    ולעדכן, לא לאפס. לכל רשומה:
      - נבחן שכבר בגיליון → מעדכן שדות זיהוי בלבד
      - נבחן חדש          → נוסף בסוף

    include_scan_columns=True נדרש לסנכרון נוכחות (sync_exam_to_drive),
    שבו עמודות הסריקה הן כל מטרת הכתיבה. בייבוא רשימה הוא נשאר False,
    אחרת רשימה טרייה הייתה מוחקת נוכחות שכבר נסרקה.

    מחזיר (added, updated).
    """
    all_values = ws.get_all_values()
    header_idx = find_header_row(all_values)

    if header_idx == -1:
        has_content = any(any(str(c).strip() for c in row) for row in all_values)
        if has_content:
            # יש תוכן שלא זיהינו ככותרות — עדיף להיכשל מאשר לדרוס נתונים
            raise ValueError(
                "הגיליון מכיל נתונים אך לא נמצאה בו שורת כותרות מזוהה. "
                "בדוק את הגיליון ידנית לפני ייבוא."
            )
        apply_header_formatting(ws, title_text or 'נבחנים')
        all_values = ws.get_all_values()
        header_idx = find_header_row(all_values)
        if header_idx == -1:
            header_idx = HEADER_ROWS - 1

    headers = [str(h).strip() for h in all_values[header_idx]]
    mapping = map_columns(headers)
    id_col = mapping.get('id_number')
    name_col = mapping.get('full_name')

    def row_key(id_number, full_name):
        """ת.ז היא המפתח; בהיעדרה — שם מלא, כדי שרשומה בלי ת.ז לא תשוכפל."""
        id_number = str(id_number or '').strip()
        if id_number:
            return ('id', id_number)
        full_name = str(full_name or '').strip()
        return ('name', full_name) if full_name else None

    # מיפוי מפתח → אינדקס שורה בגיליון (1-based)
    existing = {}
    for offset, row in enumerate(all_values[header_idx + 1:], start=header_idx + 2):
        key = row_key(
            row[id_col] if id_col is not None and id_col < len(row) else '',
            row[name_col] if name_col is not None and name_col < len(row) else '',
        )
        if key:
            existing.setdefault(key, offset)

    scan_cols = {mapping[f] for f in SCAN_FIELDS if f in mapping}
    protected = set() if include_scan_columns else scan_cols
    # בסנכרון נוכחות המסד הוא מקור האמת, ולכן ערך ריק שם חייב לנקות
    # את התא בגיליון (ביטול סריקה). בייבוא ערך ריק לעולם לא מוחק.
    blankable = scan_cols if include_scan_columns else set()

    updates = []
    to_append = []
    updated = 0
    queued = set()   # מפתחות שכבר נוספו בקובץ הנוכחי

    for rec in records:
        values = build_examinee_row(rec, headers)
        key = row_key(rec.get('id_number'), rec.get('full_name'))
        if key and key in queued:
            continue  # אותה רשומה מופיעה פעמיים באותו קובץ
        sheet_row = existing.get(key) if key else None

        if sheet_row:
            current = all_values[sheet_row - 1]
            for idx, val in enumerate(values):
                if idx in protected:
                    continue  # נתוני סריקה מוגנים מפני ייבוא רשימה
                if val == '' and idx not in blankable:
                    continue  # ערך ריק לא מוחק נתון קיים
                old = current[idx].strip() if idx < len(current) else ''
                if old != str(val):
                    updates.append({
                        'range': f'{_a1_col(idx)}{sheet_row}',
                        'values': [[val]],
                    })
            updated += 1
        else:
            to_append.append(values)
            if key:
                queued.add(key)

    if updates:
        ws.batch_update(updates, value_input_option='USER_ENTERED')
    if to_append:
        ws.append_rows(to_append, value_input_option='USER_ENTERED',
                       table_range=f'A{header_idx + 1}')

    print(f"[Sheets] Merged into '{ws.title}': {len(to_append)} added, {updated} updated", flush=True)
    return len(to_append), updated


def load_examinee_records(ws):
    """
    קורא את כל הנבחנים מגיליון — לשימוש כמטמון-תהליך בזיכרון (RAM בלבד,
    לא נשמר לדיסק). המערכת הזו לא מחזיקה שום מסד נתונים לנבחנים —
    Google Drive הוא מקור האמת היחיד, וזו רק קריאה שלו.

    מחזיר dict: מפתח → dict(שדות...). מפתח הוא ת.ז, ובהיעדרה שם מלא
    (כמו במיזוג הייבוא, כדי לא לפספס נבחן בלי ת.ז).
    """
    all_values = ws.get_all_values()
    header_idx = find_header_row(all_values)
    if header_idx == -1:
        return {}

    headers = [str(h).strip() for h in all_values[header_idx]]
    mapping = map_columns(headers)
    id_col = mapping.get('id_number')
    name_col = mapping.get('full_name')

    records = {}
    for row in all_values[header_idx + 1:]:
        def get(field):
            idx = mapping.get(field)
            return row[idx].strip() if idx is not None and idx < len(row) else ''

        id_number = get('id_number')
        full_name = get('full_name')
        key = id_number or full_name
        if not key:
            continue
        records[key] = {f: get(f) for f, _ in FIELD_KEYWORDS}
    return records


def write_examinee_scan(ws, id_number, full_name='', computer='', col='', seat='',
                        pc_status='', scan_time='', technician='', is_present=1,
                        username='', password=''):
    """
    כותב תוצאת סריקה בודדת לגיליון — פונקציה אחת לכל נקודות הסריקה
    (סריקה כפולה, ביקון, סריקה פשוטה), כדי שלא יהיו כמה מימושים
    שסוטים זה מזה עם הזמן.

    מוצא שורה קיימת לפי ת.ז (ובהיעדרה לפי שם) ומעדכן; אם לא נמצאה —
    מוסיף שורה חדשה. לעולם לא נוגע בעמודת 'גרסה' — זו שייכת לרשימת
    הייבוא, לא לנתוני הסריקה.
    """
    all_data = ws.get_all_values()
    header_idx = find_header_row(all_data)
    if header_idx == -1:
        header_idx = 0
    headers = [str(h).strip() for h in (all_data[header_idx] if all_data else [])]
    mapping = map_columns(headers)

    id_col = mapping.get('id_number')
    name_col = mapping.get('full_name')

    target_row = None
    for i, row in enumerate(all_data[header_idx + 1:], start=header_idx + 1):
        if (id_number and id_col is not None and id_col < len(row)
                and str(row[id_col]).strip() == str(id_number).strip()):
            target_row = i
            break
        if (target_row is None and full_name and name_col is not None and name_col < len(row)
                and str(row[name_col]).strip() == str(full_name).strip()):
            target_row = i
            break

    scan_values = {
        'computer': computer, 'is_present': str(is_present), 'scan_time': scan_time,
        'technician': technician, 'pc_status': pc_status,
    }

    if target_row is not None:
        sheet_row = target_row + 1
        existing = all_data[target_row]
        updates = []
        for field, val in scan_values.items():
            idx = mapping.get(field)
            if idx is not None and val:
                updates.append({'range': f'{_a1_col(idx)}{sheet_row}', 'values': [[val]]})
        # username/password: רק אם התא ריק — לא דורסים מה שכבר בגיליון
        for field, val in [('username', username), ('password', password)]:
            idx = mapping.get(field)
            if idx is not None and val:
                current = existing[idx].strip() if idx < len(existing) else ''
                if not current:
                    updates.append({'range': f'{_a1_col(idx)}{sheet_row}', 'values': [[val]]})
        if updates:
            ws.batch_update(updates, value_input_option='USER_ENTERED')
        print(f"[Sheets] Updated row {sheet_row} for {full_name or id_number} "
              f"(present={is_present})", flush=True)
        return 'updated'

    # לא נמצאה שורה — נבחן חדש שנוצר ישירות מה-QR בסריקה
    new_row = [''] * max(len(headers), max(mapping.values(), default=0) + 1)
    row_values = dict(scan_values, full_name=full_name, id_number=id_number,
                      row=col, seat=seat, username=username, password=password)
    for field, val in row_values.items():
        idx = mapping.get(field)
        if idx is not None and val:
            new_row[idx] = val
    ws.append_row(new_row, value_input_option='USER_ENTERED')
    print(f"[Sheets] Appended new row for {full_name or id_number}", flush=True)
    return 'appended'


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
