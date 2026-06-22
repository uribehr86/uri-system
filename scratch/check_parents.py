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
    file_id = '1YWLJA5T8Uq7IGzlzXSA1PPwrdSIPx9eazEcwWXwh3uM'
    f = drive.files().get(fileId=file_id, fields='parents, name, owners').execute()
    parents = f.get('parents', [])
    owners = [o.get('displayName') for o in f.get('owners', [])]
    print(f"File '{f['name']}' has parents: {parents}, Owners: {owners}")
    
    # Check details of each parent folder
    for p_id in parents:
        try:
            p = drive.files().get(fileId=p_id, fields='name, owners').execute()
            p_owners = [o.get('displayName') for o in p.get('owners', [])]
            print(f"Parent Folder '{p['name']}' (ID: {p_id}), Owners: {p_owners}")
        except Exception as ex_p:
            print(f"Parent Folder ID: {p_id} (cannot retrieve info: {ex_p})")
except Exception as e:
    print(f"Error: {e}")
