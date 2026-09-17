import gspread, os
from drive_manager import get_google_credentials

scopes   = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
sheet_id = '1YWLJA5T8Uq7IGzlzXSA1PPwrdSIPx9eazEcwWXwh3uM'

creds  = get_google_credentials(scopes)
client = gspread.authorize(creds)
ws     = client.open_by_key(sheet_id).sheet1

headers = ['שם פרטי','שם משפחה','ת.ז','התאמות','סיסמה','שם משתמש','גרסה','סה\' אולם','טור','כסא','מ.מחשב','נוכחות','שעת סריקה','טכנאי']

ws.update(values=[headers], range_name='A1:N1')
print('[OK] Headers set! Columns: ' + ', '.join(headers))
