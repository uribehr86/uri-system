import gspread, os
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file(os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', 'service_account.json'), scopes=scopes)
client = gspread.authorize(creds)
sh = client.open_by_key(os.getenv('GOOGLE_SHEETS_ID'))

deleted_total = 0
for ws in sh.worksheets():
    rows = ws.get_all_values()
    rows_to_delete = []
    for i, row in enumerate(rows[1:], start=2):
        val = str(row[0]) if row else ''
        if '|' in val or len(val) > 50:
            rows_to_delete.append(i)
    for row_idx in reversed(rows_to_delete):
        ws.delete_rows(row_idx)
        deleted_total += 1

print(f"Done. Deleted {deleted_total} rows.")
