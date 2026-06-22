import gspread
from google.oauth2.service_account import Credentials
import os
import sys

# Set encoding for Windows output
sys.stdout.reconfigure(encoding='utf-8')

service_account_file = "service_account.json"
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

try:
    creds = Credentials.from_service_account_file(service_account_file, scopes=scopes)
    client = gspread.authorize(creds)
    
    # Try to open the spreadsheet
    sheet_id = "1gs3X9KctkQsju_gNpSjyTeeGesksgWKwYVs8S5aX2pQ"
    print(f"Attempting to open spreadsheet {sheet_id}...")
    sh = client.open_by_key(sheet_id)
    print("Successfully opened spreadsheet!")
    print("Spreadsheet Title:", sh.title)
    
    # List worksheets
    worksheets = sh.worksheets()
    print("Worksheets:")
    for ws in worksheets:
        print(f" - {ws.title}")
        
except Exception as e:
    print("Error:", e)
