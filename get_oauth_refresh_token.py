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
    print("הצליח! תעתיק את שלושת הערכים האלה ל-Render → Environment:")
    print("=" * 70)
    print(f"GOOGLE_OAUTH_CLIENT_ID={installed.get('client_id')}")
    print(f"GOOGLE_OAUTH_CLIENT_SECRET={installed.get('client_secret')}")
    print(f"GOOGLE_OAUTH_REFRESH_TOKEN={creds.refresh_token}")
    print("=" * 70)
    if not creds.refresh_token:
        print("\n⚠️  לא התקבל refresh_token! נסה שוב אחרי שתבטל גישה קודמת")
        print("   דרך https://myaccount.google.com/permissions ואז תריץ שוב.")


if __name__ == '__main__':
    main()
