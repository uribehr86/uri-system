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
    results = drive.files().list(
        q="trashed=false",
        fields='files(id, name, size, mimeType)'
    ).execute()
    files = results.get('files', [])
    print(f"Found {len(files)} files in Drive:")
    for f in files:
        size = f.get('size', 'unknown')
        print(f"- {f['name']} (ID: {f['id']}), Mime: {f['mimeType']}, Size: {size}")
        
    # Let's delete them to free up quota
    for f in files:
        print(f"Deleting {f['name']}...")
        drive.files().delete(fileId=f['id']).execute()
    print("Cleanup completed.")
except Exception as e:
    print(f"Error: {e}")
