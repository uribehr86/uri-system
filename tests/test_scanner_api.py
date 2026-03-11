import requests
import json

# URL של השרת המקומי (וודא שהשרת רץ לפני ההרצה)
BASE_URL = "http://127.0.0.1:5000"

def test_scan_api(barcode):
    print(f"--- בודק סריקה עבור ברקוד: {barcode} ---")
    
    # 1. שליחת סריקה לעיבוד
    # הערה: בגלל שיש login_required, הטסט הזה עשוי להיכשל אם אין Session.
    # לצורך הבדיקה הראשונית, אנחנו מניחים שהשרת מאפשר גישה או שנבדוק את הלוגיקה הפנימית.
    
    payload = {"barcode": barcode}
    try:
        response = requests.post(f"{BASE_URL}/api/process-scan", json=payload)
        if response.status_code == 200:
            print("✅ API process-scan עובד!")
            data = response.json()
            print(f"תוצאה: {'נמצא קיים' if data.get('exists') else 'נוצר חדש'}")
            return data.get('computer', {}).get('id')
        elif response.status_code == 302 or response.status_code == 401:
            print("❌ שגיאה: נדרשת התחברות (Login Required).")
        else:
            print(f"❌ שגיאה בקריאת API: {response.status_code}")
            print(response.text)
    except Exception as e:
        print(f"❌ שגיאה בחיבור לשרת: {e}")
    return None

if __name__ == "__main__":
    print("🚀 מתחיל טסט סריקה...")
    # נסה לסרוק ברקוד דמיוני
    comp_id = test_scan_api("TEST-BARCODE-123")
    
    if comp_id:
        print(f"\n✅ המערכת זיהתה/יצרה את המחשב עם ID: {comp_id}")
        print("עכשיו אתה יכול לראות אותו בטבלת המחשבים באתר!")
    else:
        print("\n💡 וודא שהשרת רץ (python flask_app.py) לפני הרצת הטסט.")
