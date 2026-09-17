# -*- coding: utf-8 -*-
"""
get_oauth_refresh_token.py
===========================
מריצים את זה **פעם אחת, מקומית** (לא ב-Render!) כדי להפיק refresh
token לחשבון ה-Gmail האישי שלך. אחרי זה שומרים שלושה ערכים
ב-Environment Variables של Render, ולא צריך להריץ את זה שוב (אלא אם
מבטלים את הגישה מצד גוגל).

למה זה צריך לרוץ מקומית ולא ב-Render: זה פותח דפדפן ומבקש ממך
להתחבר ולאשר גישה — Render הוא שרת בלי דפדפן ובלי אתה יושב מולו.

── הכנה חד-פעמית ──────────────────────────────────────────────────
1. התקנה (רק פעם אחת, במחשב שלך):
     pip install google-auth-oauthlib

2. ב-Google Cloud Console (אותו פרויקט של ה-service account,
   uri-system-sheets):
   - APIs & Services → Credentials → Create Credentials → OAuth client ID
   - Application type: Desktop app
   - הורד את קובץ ה-JSON שנוצר, שמור אותו כאן בתיקייה בשם
     client_secret.json (או שנה את המשתנה CLIENT_SECRETS_FILE למטה)
   - אם עדיין אין OAuth consent screen מוגדר: APIs & Services →
     OAuth consent screen → User Type: External → תוסיף את עצמך
     (uribehr@gmail.com) כ-Test user

3. הרצה:
     python get_oauth_refresh_token.py
   דפדפן ייפתח — תתחבר עם uribehr@gmail.com ותאשר.

4. הסקריפט ידפיס שלושה ערכים — תעתיק אותם ל-Render:
     GOOGLE_OAUTH_CLIENT_ID
     GOOGLE_OAUTH_CLIENT_SECRET
     GOOGLE_OAUTH_REFRESH_TOKEN
"""
import json
import sys

# ווינדוס פותח את הפלט ב-cp1252 כשמפנים אותו לקובץ או ל-pipe, ואז כל
# הדפסה בעברית מפילה את הסקריפט ב-UnicodeEncodeError — אחרי שהמשתמש
# כבר עבר את כל תהליך ההרשאה בדפדפן. errors='replace' מוודא שזה לא
# יקרה: במקרה הגרוע תווים יוצגו כסימני שאלה, אבל שום דבר לא ייפול.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

CLIENT_SECRETS_FILE = 'client_secret.json'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]


def main():
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        print("חסרה החבילה google-auth-oauthlib.")
        print("הרץ קודם: pip install google-auth-oauthlib")
        sys.exit(1)

    import os
    if not os.path.exists(CLIENT_SECRETS_FILE):
        print(f"לא נמצא הקובץ '{CLIENT_SECRETS_FILE}'.")
        print("הורד אותו מ-Google Cloud Console (Credentials → OAuth client ID → Desktop app)")
        print(f"ושמור אותו כאן בתיקייה בשם {CLIENT_SECRETS_FILE}.")
        sys.exit(1)

    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
    # access_type='offline' + prompt='consent' מבטיחים שגוגל באמת יחזיר
    # refresh_token (לא רק access_token זמני) — בלי זה לפעמים גוגל
    # מדלג על זה אם כבר אישרת בעבר לאותו client_id
    creds = flow.run_local_server(port=0, access_type='offline', prompt='consent')

    with open(CLIENT_SECRETS_FILE, encoding='utf-8') as f:
        client_config = json.load(f)
    installed = client_config.get('installed', client_config.get('web', {}))

    print("\n" + "=" * 70)
    if not creds.refresh_token:
        print("לא התקבל refresh_token!")
        print("בטל את הגישה הקיימת ב-https://myaccount.google.com/permissions")
        print("ואז הרץ את הסקריפט שוב.")
        print("=" * 70)
        return

    # הטוקן לא מודפס למסך בכוונה: העתקה ידנית בעכבר מ-Console חותכת
    # אותו באמצע בקלות (הוא ארוך ונשבר לשורות), והתוצאה היא טוקן פגום
    # שנכנס ל-Render, עובר Deploy בהצלחה, ונכשל רק בשימוש אמיתי.
    copied = copy_to_clipboard(creds.refresh_token)

    print("הצליח!")
    print("=" * 70)
    print(f"GOOGLE_OAUTH_CLIENT_ID={installed.get('client_id')}")
    print(f"GOOGLE_OAUTH_CLIENT_SECRET={mask(installed.get('client_secret'))}")
    print(f"GOOGLE_OAUTH_REFRESH_TOKEN={mask(creds.refresh_token)}")
    print("=" * 70)

    if copied:
        print("\nהטוקן המלא הועתק ללוח (clipboard).")
        print("ב-Render, בשדה GOOGLE_OAUTH_REFRESH_TOKEN:")
        print("  Ctrl+A  (בחר הכל)  ואז  Ctrl+V  (הדבק)")
    else:
        print("\nלא הצלחתי להעתיק ללוח. להדפסת הטוקן המלא:")
        print("  python get_oauth_refresh_token.py --show")

    print("\nה-CLIENT_SECRET המלא נמצא בקובץ client_secret.json שבתיקייה.")


def mask(value):
    """מציג קצוות בלבד — מספיק לזהות, לא מספיק לדלוף מצילום מסך."""
    if not value:
        return '(ריק)'
    if '--show' in sys.argv:
        return value
    if len(value) <= 12:
        return value[:2] + '…'
    return f"{value[:6]}…{value[-4:]}  ({len(value)} תווים)"


def copy_to_clipboard(text):
    """מעתיק ללוח. מחזיר True בהצלחה. לא מפיל את הסקריפט בכישלון."""
    import subprocess
    commands = {
        'win32':  ['clip'],
        'darwin': ['pbcopy'],
    }
    cmd = commands.get(sys.platform, ['xclip', '-selection', 'clipboard'])
    try:
        # הטוקן הוא base64url — ASCII בלבד, אז utf-8 בטוח בכל הפלטפורמות
        proc = subprocess.run(cmd, input=text.encode('utf-8'), check=True)
        return proc.returncode == 0
    except Exception:
        return False


if __name__ == '__main__':
    main()
