import sys
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

sa_file = 'service_account.json'
scopes = ['https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(sa_file, scopes=scopes)
drive = build('drive', 'v3', credentials=creds)

try:
    print("Emptying Google Drive trash for the service account...")
    drive.files().emptyTrash().execute()
    print("Trash emptied successfully!")
except Exception as e:
    print(f"Error: {e}")
