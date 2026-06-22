import sys
import os
import gspread
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

sa_file = 'service_account.json'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

if not os.path.exists(sa_file):
    print("Error: service_account.json not found.")
    sys.exit(1)

try:
    creds = Credentials.from_service_account_file(sa_file, scopes=scopes)
    client = gspread.authorize(creds)
    
    # Open the existing sheet owned by you
    sheet_id = '1YWLJA5T8Uq7IGzlzXSA1PPwrdSIPx9eazEcwWXwh3uM'
    sh = client.open_by_key(sheet_id)
    
    # Try creating the test tab
    tab_name = 'רופאים בדיקה 17.6.26'
    
    # Check if exists, delete to overwrite
    try:
        ws = sh.worksheet(tab_name)
        sh.del_worksheet(ws)
        print(f"Removed existing tab '{tab_name}' first.")
    except Exception:
        pass
        
    print(f"Creating new tab '{tab_name}' inside your spreadsheet...")
    ws_tab = sh.add_worksheet(title=tab_name, rows=1000, cols=15)
    
    print("Writing header and mock examinee...")
    header = ['שם נבחן', 'תעודת זהות', 'קוד משתמש', 'סיסמה', 'טור', 'כסא', 'מחשב', 'הערות', 'בחינה', 'מיקום', 'נוכחות', 'מצב מחשב', 'שעת סריקה']
    mock_student = ['ישראל ישראלי', '123456789', 'israel123', 'pass123', 'א', '12', '', 'הערת בדיקה', 'רופאים בדיקה 17.6.26', 'כיתה א', '', '', '']
    ws_tab.update('A1', [header, mock_student], value_input_option='USER_ENTERED')
    
    print("\n================ SUCCESS! ================")
    print(f"הלשונית '{tab_name}' נוצרה בהצלחה!")
    print(f"קישור ישיר למסמך בגוגל שיטס:")
    print(f"https://docs.google.com/spreadsheets/d/{sheet_id}")
    print("==========================================")
    
except Exception as e:
    print(f"Error occurred: {e}")
