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
    folder_id = '1bmoF9oe2O6hB4v2MCJEtWRWI7nnO8llv'
    f = drive.files().get(fileId=folder_id, fields='name, owners, capabilities').execute()
    owners = [o.get('displayName') for o in f.get('owners', [])]
    capabilities = f.get('capabilities', {})
    print(f"Folder Name: {f['name']}")
    print(f"Owners: {owners}")
    print(f"Can Add Children: {capabilities.get('canAddChildren')}")
    print(f"Can Edit: {capabilities.get('canEdit')}")
except Exception as e:
    print(f"Error: {e}")
