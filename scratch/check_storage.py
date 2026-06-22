import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

creds = Credentials.from_service_account_file(
    'service_account.json',
    scopes=['https://www.googleapis.com/auth/drive']
)
drive = build('drive', 'v3', credentials=creds)

# בדיקת storage
about = drive.about().get(fields='storageQuota,user').execute()
quota = about.get('storageQuota', {})
user = about.get('user', {})

print(f"User: {user.get('emailAddress', 'unknown')}")
limit = int(quota.get('limit', 0))
usage = int(quota.get('usage', 0))
usage_drive = int(quota.get('usageInDrive', 0))
usage_trash = int(quota.get('usageInDriveTrash', 0))

print(f"Limit:  {limit/1024/1024/1024:.2f} GB")
print(f"Usage:  {usage/1024/1024:.1f} MB")
print(f"Drive:  {usage_drive/1024/1024:.1f} MB")
print(f"Trash:  {usage_trash/1024/1024:.1f} MB")

# נסיון ריקון אשפה
print("\nEmptying trash...")
try:
    drive.files().emptyTrash().execute()
    print("Trash emptied!")
except Exception as e:
    print(f"Trash error: {e}")
