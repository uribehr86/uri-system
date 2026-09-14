"""
sync_attendance_drive.py
========================
סנכרון נוכחות נבחנים לגיליון המבחן ב-Google Drive.

לוגיקה:
1. מפרק את שם המבחן (exam_naming) למשרד / מבחן / תאריך
2. מוצא/יוצר: PARENT_FOLDER_ID / <משרד> / <מבחן_תאריך>
3. ממזג את הנבחנים לגיליון — מעדכן קיימים, מוסיף חדשים.
   לעולם לא מוחק: מבחן מחולק לאולמות מסונכרן בכמה פעימות.
"""
import os
import traceback
from dotenv import load_dotenv
load_dotenv()

from exam_naming import parse_exam_title
from drive_manager import get_exam_sheet, merge_examinees_into_sheet


def _present_flag(value):
    return "1" if value in (True, 1, '1', 'true', 't', 'כן') else ""


def _scan_time_str(value):
    if not value:
        return ''
    if isinstance(value, str):
        return value
    try:
        return value.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(value)


def _to_sheet_record(r):
    """
    ממיר שורת examinees (שמות העמודות לפי הגדרת הטבלה) לרשומה לגיליון.
    שמות העמודות: hall / computer / technician — לא classroom/laptop_number.
    """
    return {
        'full_name':   r.get('full_name', ''),
        'id_number':   r.get('id_number', ''),
        'username':    r.get('username', ''),
        'password':    r.get('password', ''),
        'adaptations': r.get('adaptations', ''),
        'hall':        r.get('hall', ''),
        'row':         r.get('row', ''),
        'seat':        r.get('seat', ''),
        'computer':    r.get('computer', ''),
        'is_present':  _present_flag(r.get('is_present')),
        'scan_time':   _scan_time_str(r.get('scan_time')),
        'technician':  r.get('technician', ''),
        'version':     r.get('exam_name', ''),
    }


def sync_exam_to_drive(exam_name, examinees_for_exam):
    """
    מקבל שם מבחן (למשל 'משרד הבריאות - רופאים 6.10.26') ורשימת נבחנים,
    ומסנכרן לגיליון הנכון ב-Drive.

    מחזיר: (sheet_id, sheet_url) או (None, None) אם נכשל.
    """
    try:
        info = parse_exam_title(exam_name)
        print(f"[Drive] Syncing: office='{info['office']}' sheet='{info['sheet_name']}'")

        target = get_exam_sheet(exam_name, create=True)
        if not target:
            print("[Drive ERROR] Could not resolve exam sheet")
            return None, None

        records = [_to_sheet_record(r) for r in examinees_for_exam]
        # מיזוג — הנוכחות שכבר בגיליון לא נמחקת, וגם שאר המבחנים של המשרד
        merge_examinees_into_sheet(target['worksheet'], records,
                                   title_text=info['exam_name'])

        print(f"[Drive] Success! URL: {target['url']}")
        return target['sheet_id'], target['url']

    except Exception as e:
        print(f"[Drive ERROR] {e}")
        traceback.print_exc()
        return None, None


# ── הרצה עצמאית לבדיקה ──────────────────────────────────────────────
if __name__ == '__main__':
    import sys, io
    if hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

    test_name = sys.argv[1] if len(sys.argv) > 1 else 'משרד הבריאות - רופאים 6.10.26'
    test_examinees = [
        {'full_name': 'ישראל ישראלי', 'id_number': '123456789', 'username': 'israel1',
         'computer': '5001', 'is_present': True, 'scan_time': None, 'hall': 'אולם 1',
         'adaptations': '', 'exam_name': test_name},
        {'full_name': 'שרה כהן', 'id_number': '987654321', 'username': 'sarah2',
         'computer': '5002', 'is_present': False, 'scan_time': None, 'hall': 'אולם 1',
         'adaptations': 'הארכת זמן', 'exam_name': test_name},
    ]
    sid, url = sync_exam_to_drive(test_name, test_examinees)
    if sid:
        print(f'\nSheet ID: {sid}')
        print(f'Link: {url}')
    else:
        print('\nFailed!')
