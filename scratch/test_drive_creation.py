import sys
import os
import re
import gspread

sys.stdout.reconfigure(encoding='utf-8')

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

PARENT_FOLDER_ID = '1bmoF9oe2O6hB4v2MCJEtWRWI7nnO8llv'
sa_file = 'service_account.json'

if not os.path.exists(sa_file):
    print("Error: service_account.json not found.")
    sys.exit(1)

try:
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file(sa_file, scopes=scopes)
    client = gspread.authorize(creds)
    drive = build('drive', 'v3', credentials=creds)

    ministry_name = 'משרד הבריאות - בדיקה'
    sheet_file_name = 'רופאים בדיקה 17.6.26'

    print(f"1. Searching or creating folder: {ministry_name}")
    folder_id = None
    q = (f"name='{ministry_name}' and "
         f"mimeType='application/vnd.google-apps.folder' and "
         f"'{PARENT_FOLDER_ID}' in parents and trashed=false")
    folders = drive.files().list(q=q, fields='files(id,name)').execute().get('files', [])
    if folders:
        folder_id = folders[0]['id']
        print(f"   Folder already exists in Parent: {folder_id}")
    else:
        folder_id = drive.files().create(body={
            'name': ministry_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [PARENT_FOLDER_ID]
        }, fields='id').execute()['id']
        print(f"   Created new folder in Parent: {folder_id}")

    print(f"2. Searching or creating spreadsheet: {sheet_file_name}")
    q2 = (f"name='{sheet_file_name}' and "
          f"mimeType='application/vnd.google-apps.spreadsheet' and "
          f"'{folder_id}' in parents and trashed=false")
    sheets = drive.files().list(q=q2, fields='files(id,name)').execute().get('files', [])
    if sheets:
        ss_id = sheets[0]['id']
        print(f"   Spreadsheet already exists: {ss_id}")
    else:
        ss_id = drive.files().create(body={
            'name': sheet_file_name,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [folder_id]
        }, fields='id').execute()['id']
        print(f"   Created new spreadsheet: {ss_id}")

    sh = client.open_by_key(ss_id)
    ws_tab = sh.sheet1
    try:
        ws_tab.update_title('נוכחות')
    except Exception:
        pass

    # Share spreadsheet to anyone with link so user can open it!
    try:
        drive.permissions().create(
            fileId=ss_id,
            body={'role': 'reader', 'type': 'anyone'},
            fields='id'
        ).execute()
        print("   Shared sheet with 'anyone with link' for testing.")
    except Exception as e_share:
        print(f"   Could not share sheet: {e_share}")

    print("3. Writing test header and student data...")
    header = ['שם נבחן', 'תעודת זהות', 'קוד משתמש', 'סיסמה', 'טור', 'כסא', 'מחשב', 'הערות', 'בחינה', 'מיקום', 'נוכחות', 'מצב מחשב', 'שעת סריקה']
    mock_student = ['ישראל ישראלי', '123456789', 'israel123', 'pass123', 'א', '12', '', 'הערת בדיקה', 'רופאים בדיקה 17.6.26', 'כיתה א', '', '', '']
    ws_tab.clear()
    ws_tab.update('A1', [header, mock_student], value_input_option='USER_ENTERED')
    
    sheet_url = f"https://docs.google.com/spreadsheets/d/{ss_id}"
    print("\n================ SUCCESS! ================")
    print(f"התיקייה שנוצרה: {ministry_name}")
    print(f"קובץ הגיליון שנוצר: {sheet_file_name}")
    print(f"קישור ישיר לגיליון בגוגל שיטס:")
    print(sheet_url)
    print("==========================================")

except Exception as e:
    print(f"Error occurred: {e}")
