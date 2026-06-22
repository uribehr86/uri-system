import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

creds = Credentials.from_service_account_file(
    'service_account.json',
    scopes=['https://www.googleapis.com/auth/drive']
)
drive = build('drive', 'v3', credentials=creds)

res = drive.files().list(
    q="mimeType='application/vnd.google-apps.folder'",
    fields='files(id,name)',
    pageSize=30
).execute()

folders = res.get('files', [])
if folders:
    for f in folders:
        print(f"{f['id']}  -->  {f['name']}")
else:
    print("NO FOLDERS - need to share a folder with the service account!")
