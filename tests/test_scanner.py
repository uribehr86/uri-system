import requests
import json

# URL של השרת המקומי (וודא ש-flask_app.py רץ)
BASE_URL = "http://127.0.0.1:5000"

def test_scan_logic(barcode):
    print(f"--- בדיקת סורק עבור ברקוד טסט: {barcode} ---")
    
    payload = {"barcode": barcode}
    try:
        # שליחת בקשה ל-API
        response = requests.post(f"{BASE_URL}/api/process-scan", json=payload)
        
        if response.status_code == 200:
            data = response.json()
            print("✅ השרת הגיב!")
            print(f"תוצאה: {'מחשב נמצא' if data.get('exists') else 'מחשב חדש נוצר'}")
            print(f"מידע: {json.dumps(data.get('computer'), indent=2, ensure_ascii=False)}")
            return data.get('computer', {}).get('id')
        else:
            print(f"❌ שגיאה: סטטוס {response.status_code}")
            print("וודא שאתה מחובר ושהשרת רץ.")
    except Exception as e:
        print(f"❌ שגיאה בתקשורת: {e}")
    return None

if __name__ == "__main__":
    # שימוש בברקוד דמיוני שלא יפריע למלאי האמיתי
    dummy_barcode = "DEBUG-123-URI"
    test_scan_logic(dummy_barcode)
