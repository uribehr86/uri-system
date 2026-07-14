"""
Targeted fix: replace ONLY the specific garbled Hebrew strings that affect functionality.
Works by finding exact byte patterns and replacing them.
"""

# Map of garbled (stored as UTF-8 of Latin-1 of UTF-8 Hebrew) → correct Hebrew
# Each garbled string was identified from the file
REPLACEMENTS = [
    # 'תקין' stored garbled
    ('×ª×§×™×Ÿ', 'תקין'),
    # 'ממתין למחיקה' stored garbled  
    ('×ž×ž×ª×™×Ÿ ×œ×ž×—×™×§×"', 'ממתין למחיקה'),
    # 'מחשב נוסף בהצלחה!'
    ('×ž×—×©×' × ×•×¡×£ ×'×"×¦×œ×—×"!', 'מחשב נוסף בהצלחה!'),
    # 'פרטי המחשב עודכנו!'
    ('×¤×¨×˜×™ ×"×ž×—×©×' ×¢×•×"×›× ×•!', 'פרטי המחשב עודכנו!'),
    # 'המחשב נמחק מהמערכת סופית (admin_uri)'
    ('×"×ž×—×©×' × ×ž×—×§ ×ž×"×ž×¢×¨×›×ª ×¡×•×¤×™×ª (admin_uri)', 'המחשב נמחק מהמערכת סופית (admin_uri)'),
    # flash messages
    ('×©×'×™××" ×'×"×•×¡×¤×ª ×ž×—×©×'', 'שגיאה בהוספת מחשב'),
    # 'לא ידוע'
    ("×œ× ×™×"×•×¢", 'לא ידוע'),
    # 'תקול' in exam keywords etc
    ('×ª×§×•×œ', 'תקול'),
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
