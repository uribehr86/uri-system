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

    # The template spreadsheet file owned by you (נוכחות בחינות)
    template_id = '1YWLJA5T8Uq7IGzlzXSA1PPwrdSIPx9eazEcwWXwh3uM'
    
    print(f"Attempting to copy template {template_id} to folder {PARENT_FOLDER_ID}...")
    
    # 1. Copy the file
    copied_file = drive.files().copy(
        fileId=template_id,
        body={
            'name': 'רופאים בדיקה מועתק 17.6.26',
            'parents': [PARENT_FOLDER_ID]
        }
    ).execute()
    
    copied_id = copied_file['id']
    print(f"Copy succeeded! Copied ID: {copied_id}")
    
    # 2. Transfer ownership to uribehr
    print("Transferring ownership back to uribehr...")
    drive.permissions().create(
        fileId=copied_id,
        body={
            'role': 'owner',
            'type': 'user',
            'emailAddress': 'uribehr@gmail.com'
        },
        transferOwnership=True
    ).execute()
    print("Ownership transferred successfully!")

except Exception as e:
    print(f"Error: {e}")
