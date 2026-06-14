import gspread, os
from google.oauth2.service_account import Credentials

scopes   = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
sa_file  = 'service_account.json'
sheet_id = '1YWLJA5T8Uq7IGzlzXSA1PPwrdSIPx9eazEcwWXwh3uM'

creds  = Credentials.from_service_account_file(sa_file, scopes=scopes)
client = gspread.authorize(creds)
ws     = client.open_by_key(sheet_id).sheet1

headers = ['שם נבחן','תעודת זהות','מחשב','נוכח','סטטוס מחשב','שם בחינה','שעת סריקה','טכנאי','טור','כסא']

ws.update(values=[headers], range_name='A1:J1')
print('[OK] Headers set! Columns: ' + ', '.join(headers))
