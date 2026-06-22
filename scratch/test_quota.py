import sys
import os
import gspread

sys.stdout.reconfigure(encoding='utf-8')

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

PARENT_FOLDER_ID = '1bmoF9oe2O6hB4v2MCJEtWRWI7nnO8llv'
sa_file = 'service_account.json'

try:
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file(sa_file, scopes=scopes)
    client = gspread.authorize(creds)
    drive = build('drive', 'v3', credentials=creds)

    sheet_file_name = 'בדיקת מכסה ישירה 17.6.26'

    print(f"Creating spreadsheet directly inside shared Parent Folder: {PARENT_FOLDER_ID}")
    ss_id = drive.files().create(body={
        'name': sheet_file_name,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': [PARENT_FOLDER_ID]
    }, fields='id').execute()['id']
    print(f"Created successfully! ID: {ss_id}")

    sh = client.open_by_key(ss_id)
    ws = sh.sheet1
    ws.update('A1', [['בדיקה', 'הצלחה']])
    print("Writing succeeded!")
    print(f"Link: https://docs.google.com/spreadsheets/d/{ss_id}")

except Exception as e:
    print(f"Error: {e}")
