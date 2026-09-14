# -*- coding: utf-8 -*-
"""
exam_naming.py
==============
פירוק אחיד של שם מבחן ל: משרד / שם מבחן / תאריך / אולם.

מקור השם הוא כותרת הקובץ (שורה 1 באקסל) או שם הקובץ — לפי החלטת המשתמש.
כל הרכיבים במערכת (ייבוא, סריקה, סנכרון ל-Drive) חייבים להשתמש בפונקציה
הזו כדי שכולם יגיעו לאותה תיקייה ולאותו גיליון.

דוגמה:
    "משרד הבריאות - רופאים 6.10.26 אולם 2"
    → office     = "משרד הבריאות"
      exam_base  = "רופאים"
      date       = "6.10.26"
      hall       = "אולם 2"
      sheet_name = "רופאים_6.10.26"     ← שם הגיליון בתוך תיקיית המשרד
      exam_name  = "משרד הבריאות - רופאים 6.10.26"   ← מפתח קנוני במסד/QR

מבנה ה-Drive שנוצר:
    PARENT_FOLDER_ID / <office> / <sheet_name>
"""

import os
import re

DEFAULT_OFFICE = 'כללי'

# תאריך: 6.10.26 / 06/10/2026 / 6-10-26
_DATE_RE = re.compile(r'\b(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})\b')
# אולם / כיתה + מספר או אות
_HALL_RE = re.compile(r'(?:אולם|כיתה)\s*[\'"]?[\w֐-׿]+[\'"]?')
# מפרידים בין משרד לשם המבחן
_SEP_RE = re.compile(r'\s[-–—]\s')
# סיומות רעש שנגררות משמות קבצים
_NOISE_RE = re.compile(r'[-–—]?\s*(נוכחות|רשימת נבחנים|רשימה|final|סופי)\s*$', re.IGNORECASE)
# תווים שאסורים/בעייתיים בשם קובץ ב-Drive ובשאילתות
_BAD_CHARS_RE = re.compile(r'[\\/\[\]*?:\'"|]')


def _squash(text):
    """מכווץ רווחים בלבד — לא נוגע בתווים (התאריך עדיין מכיל '/')"""
    return re.sub(r'\s+', ' ', str(text or '')).strip(' -–—_')


def _clean(text):
    """ניקוי מלא — רק אחרי ששלפנו תאריך/אולם, כדי לא להרוס '6/10/2026'"""
    text = _BAD_CHARS_RE.sub(' ', str(text or ''))
    return re.sub(r'\s+', ' ', text).strip(' -–—_')


def normalize_date(d, m, y):
    """6/10/2026 → 6.10.26 (תמיד אותה צורה, כדי ששני קבצים של אותו מבחן יתלכדו)"""
    y = int(y)
    if y > 100:
        y = y % 100
    return f"{int(d)}.{int(m)}.{y:02d}"


def parse_exam_title(raw, filename=None):
    """
    מפרק כותרת מבחן. אם raw ריק — נופל לשם הקובץ.
    מחזיר dict עם: office, exam_base, date, hall, sheet_name, exam_name, raw.
    לעולם לא זורק חריגה — במקרה הגרוע מחזיר ערכי ברירת מחדל.
    """
    source = (raw or '').strip()
    if not source and filename:
        source = os.path.splitext(os.path.basename(filename))[0]
    original = source
    source = _squash(_NOISE_RE.sub('', _squash(source)))

    # אולם — נשלף ומוסר, כדי ששני קבצים של אותו מבחן (אולמות שונים)
    # יגיעו לאותו גיליון בדיוק
    hall = ''
    m_hall = _HALL_RE.search(source)
    if m_hall:
        hall = _clean(m_hall.group(0))
        source = _squash(source[:m_hall.start()] + ' ' + source[m_hall.end():])

    date = ''
    m_date = _DATE_RE.search(source)
    if m_date:
        date = normalize_date(*m_date.groups())
        source = _squash(source[:m_date.start()] + ' ' + source[m_date.end():])

    # משרד — החלק שלפני המקף הראשון
    parts = _SEP_RE.split(source, maxsplit=1)
    if len(parts) == 2 and _clean(parts[0]) and _clean(parts[1]):
        office = _clean(parts[0])
        exam_base = _clean(parts[1])
    else:
        office = DEFAULT_OFFICE
        exam_base = _clean(source)

    if not exam_base:
        exam_base = _clean(original) or 'מבחן'

    sheet_name = f"{exam_base}_{date}" if date else exam_base
    exam_name = f"{office} - {exam_base} {date}".strip() if office != DEFAULT_OFFICE \
        else f"{exam_base} {date}".strip()

    return {
        'raw': original,
        'office': office,
        'exam_base': exam_base,
        'date': date,
        'hall': hall,
        'sheet_name': sheet_name,
        'exam_name': _clean(exam_name),
    }


if __name__ == '__main__':
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    samples = [
        'משרד הבריאות - רופאים 6.10.26',
        'משרד הבריאות - רופאים 6.10.26 אולם 2',
        'משרד הבריאות – רופאים 06/10/2026 - נוכחות',
        'רופאים 6.10.26',
        'משרד הבריאות - רופאים',
    ]
    for s in sys.argv[1:] or samples:
        info = parse_exam_title(s)
        print(f"{s!r}\n  office={info['office']} | sheet={info['sheet_name']} | "
              f"exam_name={info['exam_name']} | hall={info['hall']}\n")
