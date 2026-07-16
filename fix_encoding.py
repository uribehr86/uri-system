"""
Targeted fix: replace ONLY the specific garbled Hebrew strings that affect functionality.
Works by finding exact byte patterns and replacing them.
"""

# Map of garbled (stored as UTF-8 of Latin-1 of UTF-8 Hebrew) → correct Hebrew
# Each garbled string was identified from the file
REPLACEMENTS = [
    # 'תקין' stored garbled
    ('תקין', 'תקין'),
    # 'ממתין למחיקה' stored garbled  
    ('ממתין למחיק×"', 'ממתין למחיקה'),
    # 'מחשב נוסף בהצלחה!'
    ('מחש×' × וסף ×'×"צלח×"!', 'מחשב נוסף בהצלחה!'),
    # 'פרטי המחשב עודכנו!'
    ('פרטי ×"מחש×' עו×"כ× ו!', 'פרטי המחשב עודכנו!'),
    # 'המחשב נמחק מהמערכת סופית (admin_uri)'
    ('×"מחש×' × מחק מ×"מערכת סופית (admin_uri)', 'המחשב נמחק מהמערכת סופית (admin_uri)'),
    # flash messages
    ('ש×'י××" ×'×"וספת מחש×'', 'שגיאה בהוספת מחשב'),
    # 'לא ידוע'
    ("ל× י×"וע", 'לא ידוע'),
    # 'תקול' in exam keywords etc
    ('תקול', 'תקול'),
]

with open('flask_app.py', 'r', encoding='utf-8') as f:
    content = f.read()

total_fixes = 0
for garbled, correct in REPLACEMENTS:
    count = content.count(garbled)
    if count > 0:
        content = content.replace(garbled, correct)
        print(f"  Fixed {count}x: '{garbled[:20]}...' → '{correct}'")
        total_fixes += count

with open('flask_app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print(f"\nTotal fixes: {total_fixes}")

# Verify key strings now exist
for _, correct in REPLACEMENTS:
    count = content.count(correct)
    if count > 0:
        print(f"  ✓ '{correct}' found {count} times")
