import json
import ast

def parse_val(v_str):
    """Parses a string or dict into a dictionary, handling JSON and literal_eval."""
    if not v_str:
        return {}
    if isinstance(v_str, dict):
        return v_str
    try:
        res = json.loads(v_str)
        if isinstance(res, dict):
            return res
    except:
        pass
    try:
        # Handle single quotes or other Python-style dict representations
        res = ast.literal_eval(str(v_str))
        if isinstance(res, dict):
            return res
    except:
        pass
    return {}

def format_history(val_str):
    """Formats a technical history value (JSON/dict) into a readable Hebrew string."""
    val = parse_val(val_str)
    
    if not val:
        return str(val_str) if val_str else ""
        
    tmap = {
        'status': 'סטטוס',
        'location': 'מיקום',
        'cage_number': 'כלוב',
        'cage_name': 'שם כלוב',
        'case_number': 'תיק',
        'notes': 'הערות',
        'exam_appeal': 'מבחן/ערעור',
        'barcode': 'מחשב'
    }
    
    parts = []
    for k, v in val.items():
        if k in ['id', 'computer_id', 'scan_time', 'created_at', 'last_technician'] or v is None:
            continue
        parts.append(f"{tmap.get(k, k)}: {v}")
        
    return " | ".join(str(p) for p in parts) if parts else "פעולת מערכת"

def summarize_history(entry):
    """Creates a human-readable Hebrew summary of a history entry."""
    if not entry:
        return ""
    
    old = parse_val(entry.get('old_value'))
    new = parse_val(entry.get('new_value'))
    
    # Check for cage movements
    old_cage = old.get('cage_number') or old.get('cage_name')
    new_cage = new.get('cage_number') or new.get('cage_name')
    
    if old_cage and new_cage and str(old_cage).strip() != str(new_cage).strip():
        return f"העביר מכלוב {old_cage} לכלוב {new_cage}"
    
    if old_cage and not new_cage:
        # Check if it was moved to home or test
        loc = new.get('location', '')
        if 'בית' in str(loc) or 'בדיקה' in str(loc):
            return f"לקח מכלוב {old_cage} (עבודה מהבית/בדיקה)"
        return f"לקח מכלוב {old_cage}"
        
    # Default behavior: list what changed if not a simple cage move
    tmap = {
        'status': 'סטטוס',
        'location': 'מיקום',
        'cage_number': 'כלוב',
        'case_number': 'תיק',
        'notes': 'הערות',
        'exam_appeal': 'מבחן/ערעור'
    }
    
    changes = []
    for k, v in new.items():
        if k in ['id', 'computer_id', 'scan_time', 'barcode'] or v is None:
            continue
        old_v = old.get(k)
        if str(old_v) != str(v):
            changes.append(f"{tmap.get(k, k)}: {v}")
            
    if changes:
        return "שינוי: " + " | ".join(str(c) for c in changes)
    
    ctype = entry.get('change_type', 'פעולה')
    if ctype == 'Fast Scan' and not old:
        return f"נוסף מחשב חדש"
        
    return ctype
