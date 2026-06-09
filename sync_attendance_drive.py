import os
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import gspread

# ה-ID של התיקייה הראשית שיצרת בדרייב
MAIN_FOLDER_ID = '1bmoF9oe2O6hB4v2MCJEtWRWI7nnO8llv'
SERVICE_ACCOUNT_FILE = 'c:/uri system scan/uri-system/service_account.json'

def get_or_create_folder(drive_service, folder_name, parent_id):
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and '{parent_id}' in parents and trashed=false"
    results = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    files = results.get('files', [])
    if files:
        return files[0].get('id')
    
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = drive_service.files().create(body=file_metadata, fields='id').execute()
    return folder.get('id')

def get_or_create_sheet(drive_service, client, sheet_name, parent_id):
    query = f"mimeType='application/vnd.google-apps.spreadsheet' and name='{sheet_name}' and '{parent_id}' in parents and trashed=false"
    results = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    files = results.get('files', [])
    if files:
        return client.open_by_key(files[0].get('id'))
    
    # Create new sheet inside the folder
    sh = client.create(sheet_name, folder_id=parent_id)
    return sh

def update_sheet_with_attendance(sh, examinees):
    worksheet = sh.sheet1
    worksheet.clear()
    
    header = ["שם נבחן", "תעודת זהות", "קוד משתמש", "מחשב", "הגיע?", "שעת סריקה", "מיקום", "הערות"]
    data = [header]
    
    for r in examinees:
        attended = r.get('is_present', False)
        attended_str = "✅ כן" if attended in (True, 1, 'true', 't') else "❌ לא"
        
        scan_time = r.get('scan_time')
        scan_time_str = scan_time.strftime("%d/%m/%Y %H:%M") if scan_time and not isinstance(scan_time, str) else str(scan_time or '')
        
        data.append([
            r.get('full_name', ''),
            r.get('id_number', ''),
            r.get('username', ''),
            r.get('laptop_number', ''),
            attended_str,
            scan_time_str,
            r.get('classroom', ''),
            r.get('notes', '')
        ])
        
    worksheet.update(data, 'A1')
    worksheet.format("A1:H1", {
        "textFormat": {"bold": True},
        "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1.0}
    })

def sync_exam_to_drive(exam_name, examinees_for_exam):
    """
    מקבל את השם המלא של המבחן (למשל 'משרד הבריאות - רופאים 9.2.26')
    ורשימה של הנבחנים למבחן זה, ומסנכרן לתיקייה הנכונה בדרייב.
    """
    try:
        # חילוץ שם פרויקט ושם מבחן
        parts = exam_name.split('-', 1)
        if len(parts) == 2:
            project_name = parts[0].strip()
            exam_subname = parts[1].strip()
        else:
            project_name = "כללי"
            exam_subname = exam_name.strip()

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
        drive_service = build('drive', 'v3', credentials=creds)
        client = gspread.authorize(creds)

        # 1. השגת תיקיית הפרויקט
        project_folder_id = get_or_create_folder(drive_service, project_name, MAIN_FOLDER_ID)
        
        # 2. השגת קובץ האקסל
        sh = get_or_create_sheet(drive_service, client, exam_subname, project_folder_id)
        
        # 3. עדכון הנתונים
        update_sheet_with_attendance(sh, examinees_for_exam)
        print(f"✅ סונכרן בהצלחה: {project_name} -> {exam_subname}")
        
    except Exception as e:
        print(f"שגיאה בסנכרון לדרייב: {e}")
