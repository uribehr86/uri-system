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
    about = drive.about().get(fields='storageQuota, user').execute()
    print("User Details:")
    print(f"- Name: {about.get('user', {}).get('displayName')}")
    print(f"- Email: {about.get('user', {}).get('emailAddress')}")
    print("\nStorage Quota Details:")
    quota = about.get('storageQuota', {})
    limit = int(quota.get('limit', 0))
    usage = int(quota.get('usage', 0))
    print(f"- Limit: {limit} bytes ({limit / (1024*1024):.2f} MB)")
    print(f"- Usage: {usage} bytes ({usage / (1024*1024):.2f} MB)")
    print(f"- Remaining: {limit - usage} bytes ({(limit - usage) / (1024*1024):.2f} MB)")
except Exception as e:
    print(f"Error: {e}")
