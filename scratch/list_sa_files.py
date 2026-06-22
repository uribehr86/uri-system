import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

creds = Credentials.from_service_account_file(
    'service_account.json',
    scopes=['https://www.googleapis.com/auth/drive']
)
drive = build('drive', 'v3', credentials=creds)

# כל הקבצים שהסרוויס אקאונט הוא הבעלים עליהם
res = drive.files().list(
    q="'me' in owners",
    fields='files(id,name,mimeType,size)',
    pageSize=50
).execute()

files = res.get('files', [])
print(f"נמצאו {len(files)} קבצים:")
for f in files:
    size = int(f.get('size', 0))
    size_mb = f"{size/1024/1024:.1f}MB" if size > 0 else "0KB (Google file)"
    print(f"  [{f['mimeType'].split('.')[-1][:10]}] {f['name']} | {size_mb} | {f['id']}")
