import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', write_through=True)

from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify, abort
from werkzeug.security import generate_password_hash, check_password_hash
import psycopg2
import sqlite3

from psycopg2 import pool
from psycopg2.extras import RealDictCursor
import os
import secrets
import threading
from dotenv import load_dotenv
from functools import wraps
from datetime import datetime, timedelta, timezone
import json
import base64
import qrcode
import re
import uuid
import traceback
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
try:
    from google import genai
except ImportError:
    try:
        import google.generativeai as genai
    except ImportError:
        genai = None
from google_sheets_sync import sync_inventory_to_sheets
from utils import format_history, summarize_history
from threading import Timer
try:
    from docxtpl import DocxTemplate, InlineImage
    from docx.shared import Mm
    from docxcompose.composer import Composer
    from docx import Document
except ImportError as e:
    print(f"[WARNING] Word/docx packages not available: {e}")
    DocxTemplate = InlineImage = Mm = Composer = Document = None

_sync_timer = None
_last_sheets_import = None   # זמן הסנכרון האחרון שיטסâ†’אתר

def trigger_debounced_sync():
    global _sync_timer
    if _sync_timer is not None:
        _sync_timer.cancel()
    # דיליי של 5 שניות כדי לא להציף את גוגל בבקשות
    _sync_timer = Timer(5.0, sync_inventory_to_sheets)
    _sync_timer.start()

# ── AUTO-POLLER: סנכרון אוטומטי הוסר ──────────────────────────────────────────
def _auto_import_loop():
    """לולאה ברקע: מושבתת"""
    import time
    global _last_sheets_import
    # המתן 30 שניות לפני הריצה הראשונה (לתת ל-Flask להתייצב)
    time.sleep(30)
    while True:
        try:
            # הסנכרון הוסר
            pass
        except Exception as ex:
            print(f"[AUTO-SYNC] ❌ שגיאה: {ex}", flush=True)
        time.sleep(180)  # 3 דקות

# הפעל את הפולר בתחילת האפליקציה
_poller_thread = threading.Thread(target=_auto_import_loop, daemon=True, name="SheetsAutoPoller")
_poller_thread.start()
print("[AUTO-SYNC] 🔄 Auto-poller הופעל — סנכרון אוטומטי הוסר", flush=True)

# ── DB STORAGE MONITOR: checks DB size every hour ──────────────────────────
DB_SIZE_LIMIT_GB = 1.0  # plan limit in GB (Basic-256mb instance ships 1 GB of storage)

def _db_storage_monitor_loop():
    import time
    time.sleep(60)  # wait 1 min for Flask to stabilize
    while True:
        try:
            _db_url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
            if _db_url:
                import psycopg2 as _pg2
                _mc = _pg2.connect(_db_url)
                _cur2 = _mc.cursor()
                _cur2.execute('SELECT pg_database_size(current_database())')
                _sz = _cur2.fetchone()[0]
                _cur2.close()
                _mc.close()
                _gb = _sz / (1024 ** 3)
                _mb = _sz / (1024 ** 2)
                _pct = (_gb / DB_SIZE_LIMIT_GB) * 100
                if _pct >= 95:
                    print(f'[DB-MONITOR] CRITICAL: DB at {_pct:.1f}% ({_mb:.0f} MB / {DB_SIZE_LIMIT_GB:.0f} GB) - ACT NOW!', flush=True)
                elif _pct >= 85:
                    print(f'[DB-MONITOR] WARNING: DB at {_pct:.1f}% ({_mb:.0f} MB / {DB_SIZE_LIMIT_GB:.0f} GB) - start cleaning.', flush=True)
                elif _pct >= 70:
                    print(f'[DB-MONITOR] NOTICE: DB at {_pct:.1f}% ({_mb:.0f} MB / {DB_SIZE_LIMIT_GB:.0f} GB).', flush=True)
                else:
                    print(f'[DB-MONITOR] OK: {_mb:.1f} MB used ({_pct:.1f}% of {DB_SIZE_LIMIT_GB:.0f} GB)', flush=True)
        except Exception as _mon_ex:
            print(f'[DB-MONITOR] Error checking DB size: {_mon_ex}', flush=True)
        time.sleep(3600)  # check every hour

_db_monitor_thread = threading.Thread(target=_db_storage_monitor_loop, daemon=True, name='DBStorageMonitor')
_db_monitor_thread.start()
print('[DB-MONITOR] DB Storage Monitor started - checks DB size every hour', flush=True)
# טעינת הגדרות
load_dotenv()

app = Flask(__name__)

# SECRET_KEY: required from the environment in production (RENDER present).
# For local dev only, fall back to a random ephemeral key (sessions reset on restart).
_secret_key = os.getenv('SECRET_KEY')
if not _secret_key:
    if os.getenv('RENDER'):
        raise RuntimeError(
            "SECRET_KEY environment variable is required in production (RENDER detected) but is not set."
        )
    _secret_key = secrets.token_hex(32)
    print("[WARNING] SECRET_KEY not set — generated an ephemeral key for local dev "
          "(sessions will not persist across restarts).", flush=True)
app.secret_key = _secret_key

app.permanent_session_lifetime = timedelta(days=365)
app.config['SESSION_COOKIE_SECURE']   = True
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['JSON_AS_ASCII']           = False  # Hebrew in JSON stays as Hebrew
app.config['MAX_CONTENT_LENGTH']      = 16 * 1024 * 1024  # cap uploads at 16 MB

# Rate limiting (Flask-Limiter) — used to throttle sensitive routes such as /login
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

limiter = Limiter(key_func=get_remote_address, app=app)

@app.after_request
def set_utf8_charset(response):
    """Force UTF-8 charset on all HTML responses for all platforms."""
    if 'text/html' in response.content_type:
        if 'charset' not in response.content_type:
            response.content_type = 'text/html; charset=utf-8'
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma'] = 'no-cache'
    return response

@app.before_request
def refresh_session():
    """מרענן את הסשן בכל בקשה כדי שלא יפוג"""
    try:
        if 'user_id' in session:
            session.permanent = True
            session.modified  = True
        # עדכן timestamp פעם ב-5 דקות לכל יותר (למנוע עומס על DB)
        now = datetime.now()
        last_ping = session.get('_last_ping')
        if not last_ping or (now - datetime.fromisoformat(last_ping)).total_seconds() > 300:
            session['_last_ping'] = now.isoformat()
            try:
                conn = get_db_connection()
                if conn:
                    cur = get_safe_cursor(conn)
                    cur.execute("UPDATE users SET timestamp = NOW() WHERE id = %s", (session['user_id'],))
                    conn.commit()
                    cur.close()
                    release_db_connection(conn)
            except Exception:
                pass
    except Exception as _ex:
        print(f'[BEFORE_REQUEST ERROR] {_ex}', flush=True)
        session.clear()

# Initialize Google AI (Gemini)
try:
    # New google-genai SDK (preferred)
    if genai is not None and hasattr(genai, 'Client'):
        genai_client = genai.Client(api_key=os.getenv("GOOGLE_API_KEY"))
    elif genai is not None:
        # Old google-generativeai SDK fallback
        genai_client = genai.GenerativeModel('gemini-2.0-flash')
        genai_client._legacy_mode = True
    else:
        genai_client = None
        print("[WARNING] Google AI (genai) not installed — AI features disabled")
except Exception as e:
    print(f"[WARNING] Google AI init failed: {e}")
    genai_client = None

# Initialize Connection Pool variables (will be created lazily on first DB access)
db_pool = None
db_pool_initialized = False
# אם RENDER=true → ענן. אחרת → מקומי (גם אם ה-DB ענני)
IS_LOCAL_MODE = bool(os.getenv('IS_LOCAL_MODE')) or not bool(os.getenv('RENDER'))

# הרשמה עצמית סגורה כברירת מחדל: /register היה פתוח לכל האינטרנט וכל אדם
# יכול היה לפתוח לעצמו חשבון טכנאי עם גישה מלאה למלאי. משתמשים חדשים נוצרים
# דרך פאנל הניהול. להחזרה: הגדר ALLOW_SELF_REGISTRATION=true במשתני הסביבה.
ALLOW_SELF_REGISTRATION = os.getenv('ALLOW_SELF_REGISTRATION', '').strip().lower() in ('1', 'true', 'yes')

def is_test_env():
    """סביבת טסט — מסומנת בכותרת כדי שלא יתבלבלו בינה לבין הפרודקשן.
    מזוהה לפי APP_ENV=test, ואם לא הוגדר — לפי 'test' בכתובת האתר."""
    if os.getenv('APP_ENV', '').strip().lower() == 'test':
        return True
    try:
        return 'test' in (request.host or '').lower()
    except Exception:
        return False

# ── מצב DEGRADED: אין חיבור ל-PostgreSQL ──────────────────────────────────
# בענן, כשהחיבור ל-PostgreSQL נכשל, האפליקציה נופלת ל-SQLite ריק. בלי סימון
# המסך פשוט נראה ריק — כאילו הנתונים נמחקו. הדגל הזה מוצג כבאנר בכל עמוד.
DB_DEGRADED = {'active': False, 'reason': '', 'since': None}

def _mark_db_degraded(reason):
    """מסמן שהחיבור ל-PostgreSQL נכשל והאפליקציה עובדת על SQLite ריק."""
    if not DB_DEGRADED['active']:
        DB_DEGRADED['since'] = datetime.now()
        print('[DEGRADED] No PostgreSQL connection — serving from empty SQLite. '
              'A warning banner is now shown on every page.', flush=True)
    DB_DEGRADED['active'] = True
    DB_DEGRADED['reason'] = str(reason).strip()[:300]

def _mark_db_healthy():
    """מנקה את מצב ה-DEGRADED אחרי שהחיבור ל-PostgreSQL חזר."""
    if DB_DEGRADED['active']:
        print('[OK] PostgreSQL connection restored — degraded banner cleared.', flush=True)
    DB_DEGRADED['active'] = False
    DB_DEGRADED['reason'] = ''
    DB_DEGRADED['since'] = None

# attendance_cache הוסר — הוחלף ב-examinee_cache (ליד resolve_exam_sheet),
# שמחזיק גם נתוני נבחן מלאים, לא רק חותמת נוכחות

class SafeCursor:
    """Wrapper for cursor to handle %s -> ? translation for SQLite"""
    def __init__(self, cursor, is_sqlite=False):
        self.cursor = cursor
        self.is_sqlite = is_sqlite

    def execute(self, query, params=None):
        if self.is_sqlite and params:
            # Handle list for IN clauses and basic %s replacements
            if "IN (" in query and isinstance(params, (list, tuple)):
                # This is a bit tricky, but common in this app
                pass # Already handled by placeholders in most cases
            query = query.replace('%s', '?')
            # Handle ILIKE -> LIKE for SQLite (SQLite LIKE is case-insensitive usually, but ILIKE is Postgres specific)
            query = query.replace('ILIKE', 'LIKE')
            # Handle NOW() -> datetime('now')
            query = query.replace('NOW()', "datetime('now', 'localtime')")
            # Handle NULLS LAST (SQLite supports it in newer versions, but let's be safe)
            # query = query.replace('NULLS LAST', '') 
        
        try:
            if params:
                return self.cursor.execute(query, params)
            else:
                return self.cursor.execute(query)
        except Exception as e:
            print(f"[DB ERROR] Query: {query}")
            print(f"[DB ERROR] Params: {params}")
            raise e

    def fetchone(self): return self.cursor.fetchone()
    def fetchall(self): return self.cursor.fetchall()
    def close(self): return self.cursor.close()
    def __getattr__(self, name): return getattr(self.cursor, name)

def dict_factory(cursor, row):
    d = {}
    for idx, col in enumerate(cursor.description):
        d[col[0]] = row[idx]
    return d

def _resolve_db_url():
    """מחזיר את מחרוזת החיבור ל-PostgreSQL ממשתני הסביבה.
    אין ברירת מחדל קשיחה — בענן (RENDER) חובה להגדיר RENDER_DB_URL או DATABASE_URL."""
    url = os.getenv('RENDER_DB_URL') or os.getenv('DATABASE_URL')
    if not url and not IS_LOCAL_MODE:
        raise RuntimeError(
            "No database URL configured. RENDER is set, so PostgreSQL is required — "
            "set RENDER_DB_URL or DATABASE_URL in the environment."
        )
    return url

def get_db_connection():
    global db_pool, db_pool_initialized, IS_LOCAL_MODE
    
    if not db_pool_initialized:
        if IS_LOCAL_MODE:
            # מצב מקומי — דלג על PostgreSQL, עבור ישירות ל-SQLite
            db_pool = None
            db_pool_initialized = True
            print("[LOCAL] Running in local mode — using SQLite directly", flush=True)
        else:
            db_url = _resolve_db_url()
            if db_url:
                if 'connect_timeout' not in db_url:
                    db_url += ('&' if '?' in db_url else '?') + 'connect_timeout=15'
                try:
                    import urllib.parse
                    parsed = urllib.parse.urlparse(db_url)
                    host = parsed.hostname
                    if host:
                        import socket
                    
                    db_pool = psycopg2.pool.ThreadedConnectionPool(2, 20, db_url)
                    print("[OK] Database connection pool created successfully (lazy)", flush=True)
                    _mark_db_healthy()
                    # IS_LOCAL_MODE נקבע לפי RENDER env var — לא משנים כאן
                except Exception as e:
                    print(f"[FALLBACK] Cloud DB init failed: {e}. Switching to Local SQLite.", flush=True)
                    _mark_db_degraded(e)
                    db_pool = None
                    # IS_LOCAL_MODE נקבע לפי RENDER env var — לא משנים כאן
            else:
                db_pool = None
                # IS_LOCAL_MODE נקבע לפי RENDER env var — לא משנים כאן
            db_pool_initialized = True

    if IS_LOCAL_MODE:
        # מצב מקומי — SQLite בלבד
        try:
            conn = sqlite3.connect('system_data.db', check_same_thread=False)
            conn.row_factory = dict_factory
            return conn
        except Exception as e:
            print(f"[CRITICAL] SQLite connection failed: {e}", flush=True)
            return None

    if not db_pool:
        # ב-Render אבל הפול לא אותחל — ננסה חיבור ישיר לפוסטגרס
        print("[WARNING] db_pool is None on Render — attempting direct Postgres connection.", flush=True)
        try:
            _db_url = _resolve_db_url()
            direct_conn = psycopg2.connect(_db_url, connect_timeout=10)
            direct_conn.autocommit = False
            print("[OK] Direct Postgres connection established (pool was None).", flush=True)
            _mark_db_healthy()
            return direct_conn
        except Exception as e:
            print(f"[CRITICAL] Direct Postgres also failed: {e} — DATA WILL BE EMPTY (SQLite fallback)!", flush=True)
            _mark_db_degraded(e)
            try:
                conn = sqlite3.connect('system_data.db', check_same_thread=False)
                conn.row_factory = dict_factory
                return conn
            except Exception as e2:
                print(f"[CRITICAL] SQLite also failed: {e2}", flush=True)
                return None
            
    try:
        conn = db_pool.getconn()
        try:
            # Ping connection to ensure it's alive
            with conn.cursor() as c:
                c.execute('SELECT 1')
        except Exception:
            # Connection is dead, throw it away and get a new one
            db_pool.putconn(conn, close=True)
            conn = db_pool.getconn()
        _mark_db_healthy()
        return conn
    except Exception as e:
        print(f"[WARNING] Pool getconn failed ({e}), trying direct Postgres connection.", flush=True)
        # נסה חיבור ישיר לפוסטגרס במקום לנפול ל-SQLite ריק
        try:
            db_url = _resolve_db_url()
            direct_conn = psycopg2.connect(db_url, connect_timeout=10)
            direct_conn.autocommit = False
            print("[OK] Direct Postgres connection established as fallback.", flush=True)
            _mark_db_healthy()
            return direct_conn
        except Exception as e2:
            print(f"[CRITICAL] All Postgres connections failed: {e2}. Falling back to empty SQLite — DATA WILL BE EMPTY!", flush=True)
            _mark_db_degraded(e2)
            try:
                conn = sqlite3.connect('system_data.db', check_same_thread=False)
                conn.row_factory = dict_factory
                return conn
            except Exception as e3:
                print(f"[CRITICAL] SQLite also failed: {e3}")
                return None

def get_safe_cursor(conn):
    if isinstance(conn, sqlite3.Connection):
        return SafeCursor(conn.cursor(), is_sqlite=True)
    else:
        return conn.cursor(cursor_factory=RealDictCursor)

def release_db_connection(conn):
    if not conn:
        return
    if isinstance(conn, sqlite3.Connection):
        conn.close()
        return
    if not db_pool:
        # החיבור נפתח ישירות (הפול לא אותחל) — חייבים לסגור אותו בעצמנו,
        # אחרת כל בקשה מדליפה חיבור עד שנגמרות ההרשאות במסד.
        try:
            conn.close()
        except Exception as e:
            print(f"[ERROR] Failed to close direct connection: {e}", flush=True)
        return
    try:
        db_pool.putconn(conn)
    except Exception as e:
        # ייתכן שזה חיבור ישיר שלא שייך לפול — putconn ייכשל, ואז נסגור ידנית.
        print(f"[ERROR] Failed to return connection to pool: {e}", flush=True)
        try:
            conn.close()
        except Exception:
            pass

def run_startup_migrations():
    """הוספת עמודות חדשות למסד אם עדיין לא קיימות"""
    conn = get_db_connection()
    if not conn:
        return
    is_sqlite = isinstance(conn, sqlite3.Connection)
    try:
        cur = get_safe_cursor(conn)
        
        # ALWAYS initialize SQLite fallback DB to prevent crashes on Render if Postgres goes down
        try:
            with sqlite3.connect('system_data.db', check_same_thread=False) as sl_conn:
                sl_conn.row_factory = dict_factory
                sl_cur = sl_conn.cursor()
                sl_cur.execute("""CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL,
                    role TEXT NOT NULL,
                    timestamp TIMESTAMP
                )""")
                sl_cur.execute("""CREATE TABLE IF NOT EXISTS computers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    barcode TEXT,
                    case_number TEXT,
                    cage_number TEXT,
                    cage_name TEXT,
                    status TEXT,
                    location TEXT,
                    exam_appeal TEXT,
                    specs TEXT,
                    project TEXT,
                    ministry TEXT,
                    notes TEXT,
                    scan_time TEXT,
                    last_technician TEXT,
                    sheets_delete_request INTEGER DEFAULT 0
                )""")
                sl_cur.execute("""CREATE TABLE IF NOT EXISTS inventory_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    computer_id INTEGER,
                    technician TEXT,
                    change_type TEXT,
                    old_value TEXT,
                    new_value TEXT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                )""")
                sl_cur.execute("""CREATE TABLE IF NOT EXISTS projects (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    keywords TEXT NOT NULL,
                    sheets_id TEXT DEFAULT '',
                    drive_url TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")
                sl_cur.execute("""CREATE TABLE IF NOT EXISTS examinees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    exam_name TEXT DEFAULT '',
                    id_number TEXT DEFAULT '',
                    full_name TEXT DEFAULT '',
                    username TEXT DEFAULT '',
                    password TEXT DEFAULT '',
                    row TEXT DEFAULT '',
                    seat TEXT DEFAULT '',
                    hall TEXT DEFAULT '',
                    adaptations TEXT DEFAULT '',
                    computer TEXT DEFAULT '',
                    is_present INTEGER DEFAULT 0,
                    scan_time TEXT DEFAULT '',
                    technician TEXT DEFAULT '',
                    pc_status TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")
                
                try: sl_cur.execute("ALTER TABLE computers ADD COLUMN sheets_delete_request INTEGER DEFAULT 0")
                except: pass
                try: sl_cur.execute("ALTER TABLE users ADD COLUMN timestamp TIMESTAMP")
                except: pass

                sl_cur.execute("SELECT COUNT(*) as cnt FROM users")
                if sl_cur.fetchone()['cnt'] == 0:
                    from werkzeug.security import generate_password_hash
                    _seed_user = os.getenv('SEED_ADMIN_USERNAME', 'admin_uri')
                    _seed_pass = os.getenv('SEED_ADMIN_PASSWORD')
                    if not _seed_pass:
                        _seed_pass = secrets.token_urlsafe(18)
                        print("=" * 64, flush=True)
                        print(f"[SEED] Created initial admin user '{_seed_user}'.", flush=True)
                        print(f"[SEED] Generated password: {_seed_pass}", flush=True)
                        print("[SEED] Save this now — it is printed only once. "
                              "Set SEED_ADMIN_USERNAME / SEED_ADMIN_PASSWORD to control it.", flush=True)
                        print("=" * 64, flush=True)
                    sl_cur.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", (_seed_user, generate_password_hash(_seed_pass), "admin"))
                sl_conn.commit()
        except Exception as e_sl:
            print(f"[ERROR] SQLite fallback initialization failed: {e_sl}", flush=True)

        # Now do the migrations for the ACTIVE database connection
        try:
            if is_sqlite:
                cur.execute("ALTER TABLE computers ADD COLUMN sheets_delete_request INTEGER DEFAULT 0")
            else:
                cur.execute("ALTER TABLE computers ADD COLUMN IF NOT EXISTS sheets_delete_request BOOLEAN DEFAULT FALSE")
            conn.commit()
            print("[OK] Migration: added sheets_delete_request column")
        except Exception:
            conn.rollback()
        # הוסף עמודת timestamp לטבלת users אם לא קיימת
        try:
            if is_sqlite:
                cur.execute("ALTER TABLE users ADD COLUMN timestamp TIMESTAMP")
            else:
                cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS timestamp TIMESTAMP")
            conn.commit()
            print("[OK] Migration: added timestamp column to users table")
        except Exception:
            conn.rollback()
        # צור טבלת פרויקטים אם לא קיימת
        try:
            if not is_sqlite:
                cur.execute("""CREATE TABLE IF NOT EXISTS projects (
                    id SERIAL PRIMARY KEY,
                    name TEXT NOT NULL,
                    keywords TEXT NOT NULL,
                    sheets_id TEXT DEFAULT '',
                    drive_url TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")
                conn.commit()
                print("[OK] Migration: projects table ready")
        except Exception as e:
            conn.rollback()
            print(f"[WARNING] projects table: {e}")

        # צור טבלת נבחנים אם לא קיימת
        try:
            if not is_sqlite:
                cur.execute("""CREATE TABLE IF NOT EXISTS examinees (
                    id SERIAL PRIMARY KEY,
                    exam_name TEXT DEFAULT '',
                    id_number TEXT DEFAULT '',
                    full_name TEXT DEFAULT '',
                    username TEXT DEFAULT '',
                    password TEXT DEFAULT '',
                    row TEXT DEFAULT '',
                    seat TEXT DEFAULT '',
                    hall TEXT DEFAULT '',
                    adaptations TEXT DEFAULT '',
                    computer TEXT DEFAULT '',
                    is_present INTEGER DEFAULT 0,
                    scan_time TEXT DEFAULT '',
                    technician TEXT DEFAULT '',
                    pc_status TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")
                conn.commit()
                print("[OK] Migration: examinees table ready")
        except Exception as e:
            conn.rollback()
            print(f"[WARNING] examinees table: {e}")
        cur.close()
    finally:
        release_db_connection(conn)

def _run_migrations_background():
    """מריץ migrations ברקע כדי שgunicorn לא יחכה"""
    import time
    time.sleep(3)  # תן ל-gunicorn להתחיל קודם
    with app.app_context():
        try:
            run_startup_migrations()
            print("[OK] Startup migrations completed in background", flush=True)
        except Exception as _me:
            print(f"[WARNING] Startup migrations error: {_me}", flush=True)

_migration_thread = threading.Thread(target=_run_migrations_background, daemon=True, name="StartupMigrations")
_migration_thread.start()

def get_google_creds(scopes):
    """
    Load Google credentials — delegates to drive_manager.get_google_credentials,
    שהוא עכשיו המקור היחיד (תומך גם ב-OAuth של חשבון אישי, לא רק
    service account שאין לו מכסת אחסון משלו).
    """
    from drive_manager import get_google_credentials
    return get_google_credentials(scopes)

# ── מיפוי טווחי ברקוד לדגמים — מקור אמת יחיד ──────────────────────────────
# עד 10/09/2026 הכלל הזה היה משוכפל בארבעה מקומות שלא הסכימו ביניהם:
# HP נחשב 2001-2400 בשלושה מקומות ו-2001-2200 בלוח הבקרה, Lenovo 4001-4300
# מול 4001-4400, ו-HP 2023 (6001-6400) נעדר משניים מהם. התוצאה: מחשבים
# נספרו תחת יצרן שגוי או לא נספרו כלל. כל שינוי טווח נעשה כאן בלבד.
COMPUTER_MODELS = [
    {'key': 'dell2018', 'name': 'Dell 2018', 'manufacturer': 'Dell',
     'cpu': 'i7-8550U @ 1.80GHz', 'ram': '32GB', 'spec_label': 'i5/i7', 'color': '#4A90D9',
     'ranges': [(1, 600), (1001, 1600)],
     'auto_spec': 'Dell | i7-8550U @ 1.80GHz | 32GB RAM'},
    {'key': 'hp2023', 'name': 'HP 2023', 'manufacturer': 'HP',
     'cpu': 'i7-1195G7 @ 2.90GHz', 'ram': '', 'spec_label': 'i7-1195G7', 'color': '#34C759',
     'ranges': [(6001, 6400)],
     'auto_spec': '11th Gen Intel(R) Core(TM) i7-1195G7 @ 2.90GHz'},
    {'key': 'hp2018', 'name': 'HP 2018', 'manufacturer': 'HP',
     'cpu': 'i5-7200U @ 2.50GHz', 'ram': '8GB', 'spec_label': 'i5-7200U', 'color': '#FF9500',
     'ranges': [(2001, 2400)],
     'auto_spec': 'HP | i5-7200U @ 2.50GHz | 8GB RAM'},
    {'key': 'dell_barut', 'name': 'Dell Bagrut', 'manufacturer': 'Dell',
     'cpu': 'i7-8550U @ 1.80GHz', 'ram': '16GB', 'spec_label': 'i7-8550U', 'color': '#AF52DE',
     'ranges': [(3001, 3200)],
     'auto_spec': 'Dell | i7-8550U @ 1.80GHz | 16GB RAM'},
    {'key': 'lenovo', 'name': 'Lenovo', 'manufacturer': 'Lenovo',
     'cpu': 'i5-7200U @ 2.50GHz', 'ram': '8GB', 'spec_label': 'i5-7200U', 'color': '#FF2D55',
     'ranges': [(4001, 4400)],
     'auto_spec': 'Lenovo | i5-7200U @ 2.50GHz | 8GB RAM'},
]

def _model_capacity(model):
    """כמה מספרים יש בטווחי הדגם."""
    return sum(hi - lo + 1 for lo, hi in model['ranges'])

def _model_range_label(model):
    return ', '.join(f"{lo}-{hi}" for lo, hi in model['ranges'])

def get_model_for_barcode(barcode):
    """מחזיר את הגדרת הדגם לפי מספר המחשב, או None אם המספר מחוץ לכל טווח."""
    try:
        num = int(str(barcode).strip())
    except (ValueError, TypeError):
        return None
    for model in COMPUTER_MODELS:
        for lo, hi in model['ranges']:
            if lo <= num <= hi:
                return model
    return None

@app.context_processor
def utility_processor():
    def get_cage_color(cage):
        if not cage: return "inherit"
        val = sum(ord(c) for c in str(cage))
        hue = (val * 137) % 360
        return f"hsl({hue}, 70%, 65%)"

    def get_computer_spec(computer_number):
        """מחזיר מפרט לפי מספר מחשב — לפי COMPUTER_MODELS."""
        model = get_model_for_barcode(computer_number)
        if not model:
            return None
        return {'manufacturer': model['manufacturer'], 'cpu': model['cpu'],
                'ram': model['ram'], 'icon': '💻'}

    return dict(get_cage_color=get_cage_color, IS_LOCAL_MODE=IS_LOCAL_MODE, IS_TEST_ENV=is_test_env(), get_computer_spec=get_computer_spec, db_degraded=DB_DEGRADED, allow_self_registration=ALLOW_SELF_REGISTRATION, APP_VERSION="v2.7.3")

@app.template_filter('format_history')
def format_history_filter(val_str):
    return format_history(val_str)

@app.template_filter('summarize_history')
def summarize_history_filter(entry):
    return summarize_history(entry)

@app.template_filter('israel_time')
def israel_time_filter(dt):
    """ממיר datetime מ-UTC לשעון ישראל (UTC+3)"""
    if not dt:
        return '—'
    
    if isinstance(dt, str):
        try:
            # Handle possible ISO format or SQL format
            dt = datetime.fromisoformat(dt.replace('Z', '+00:00'))
        except ValueError:
            try:
                dt = datetime.strptime(dt.split('.')[0], '%Y-%m-%d %H:%M:%S')
            except ValueError:
                return dt  # Return as-is if unparseable

    if hasattr(dt, 'tzinfo') and dt.tzinfo is not None:
        il = dt.astimezone(timezone(timedelta(hours=3)))
    else:
        il = dt + timedelta(hours=3)
    return il.strftime('%H:%M %d/%m/%Y')

def get_auto_spec(barcode):
    """מחזיר מפרט אוטומטי לפי מספר מחשב — לפי COMPUTER_MODELS."""
    model = get_model_for_barcode(barcode)
    return model['auto_spec'] if model else ''

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session: return redirect(url_for('login', next=request.url))
        session.permanent = True
        return f(*args, **kwargs)
    return decorated_function

def scan1_only(f):
    """
    מגביל גישה למשתמש 'scan1' בדיוק — כולל admin. לפי בקשה מפורשת:
    רק חשבון הסריקה הייעודי יכול לסרוק נוכחות, אף אחד אחר.
    שים לב: אין כאן חריגת מנהל — אם scan1 ננעל בחוץ (סיסמה נשכחה
    וכו') צריך לאפס את הסיסמה שלו דרך ניהול משתמשים, לא לעקוף כאן.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get('username') != 'scan1':
            if request.path.startswith('/api/'):
                return jsonify({"error": "גישה לסריקה מוגבלת למשתמש scan1 בלבד"}), 403
            flash("גישה לסריקת נוכחות מוגבלת למשתמש scan1 בלבד", "danger")
            return redirect(url_for('exam_attendance'))
        return f(*args, **kwargs)
    return decorated_function

def local_only(f):
    """חוסם גישה ל-route כאשר רצים על Render (אינטרנט) — רק לשימוש מקומי"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not IS_LOCAL_MODE:
            abort(404)
        return f(*args, **kwargs)
    return decorated_function

@app.route('/api/projects')
@login_required
def api_projects():
    """מחזיר רשימת פרויקטים + Sheet IDs לסורק"""
    conn = get_db_connection()
    projects = []
    if conn:
        try:
            cur = get_safe_cursor(conn)
            cur.execute("SELECT name, keywords, sheets_id, drive_url FROM projects WHERE sheets_id != '' ORDER BY name")
            rows = cur.fetchall()
            projects = [dict(r) for r in rows]
            cur.close()
        except Exception:
            pass
        finally:
            release_db_connection(conn)
    return jsonify(projects)

@app.route('/')
def index():
    return redirect(url_for('portal')) if 'user' in session else redirect(url_for('login'))


# ── עמודים ציבוריים (בלי login_required) ────────────────────────────
# גוגל דורש דף בית ודף מדיניות פרטיות נגישים לכל אחד, בלי התחברות,
# כתנאי להוצאת אפליקציית OAuth ממצב Testing. בלי זה ה-refresh token
# פג כל 7 ימים והמערכת מפסיקה לסנכרן לדרייב.
# חשוב: אסור להוסיף כאן @login_required — גוגל חייב להגיע לעמודים.

_LEGAL_PAGE = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} · URI System</title>
<style>
  :root {{ color-scheme: dark; }}
  body {{ margin:0; padding:32px 16px; background:#0f1b33; color:#e8eefc;
         font-family:Segoe UI,Arial,sans-serif; line-height:1.7; }}
  main {{ max-width:760px; margin:0 auto; }}
  h1 {{ color:#f5b921; margin:0 0 4px; font-size:1.8rem; }}
  h2 {{ color:#8ab4ff; font-size:1.1rem; margin:28px 0 6px; }}
  .sub {{ color:#9bb0d3; font-size:.9rem; margin-bottom:28px; }}
  ul {{ padding-inline-start:20px; }}
  a {{ color:#8ab4ff; }}
  footer {{ margin-top:40px; border-top:1px solid #24365c; padding-top:16px;
            color:#9bb0d3; font-size:.85rem; }}
</style>
</head>
<body><main>
<h1>{title}</h1>
<div class="sub">URI System — מערכת ניהול מלאי ונוכחות בחינות · עודכן: 17 בספטמבר 2026</div>
{body}
<footer>
  לשאלות: <a href="mailto:uribehr@gmail.com">uribehr@gmail.com</a> ·
  <a href="/privacy">מדיניות פרטיות</a> · <a href="/terms">תנאי שימוש</a> ·
  <a href="/">דף הבית</a>
</footer>
</main></body></html>"""


@app.route('/privacy')
def privacy_policy():
    """מדיניות פרטיות — ציבורי, נדרש על ידי גוגל לאישור OAuth."""
    return _LEGAL_PAGE.format(title='מדיניות פרטיות', body="""
<h2>מי אנחנו</h2>
<p>URI System היא מערכת פנימית לניהול מלאי מחשבים ולרישום נוכחות נבחנים
בבחינות. השימוש בה מוגבל לצוות מורשה בלבד, באמצעות שם משתמש וסיסמה.
המערכת אינה פתוחה לציבור ואינה מיועדת לשימוש אישי.</p>

<h2>איזה מידע נאסף</h2>
<ul>
  <li><b>פרטי משתמשי המערכת</b> — שם משתמש וסיסמה מוצפנת של אנשי הצוות.</li>
  <li><b>נתוני מלאי</b> — מספרי מחשבים, מיקום, סטטוס תקינות.</li>
  <li><b>נתוני נוכחות בבחינה</b> — שם הנבחן, מספר זהות, שיוך לאולם ולעמדה,
      ושעת הסריקה. מידע זה נמסר למפעיל המערכת על ידי הגוף המזמין את הבחינה.</li>
</ul>

<h2>איפה המידע נשמר</h2>
<p>נתוני הנוכחות של הנבחנים נכתבים ישירות לחשבון Google Drive ו-Google Sheets
של מפעיל המערכת, ונשארים שם בלבד. הם אינם נשמרים במסד הנתונים של האתר.</p>

<h2>השימוש בהרשאות גוגל</h2>
<p>המערכת מבקשת גישה ל-Google Drive ול-Google Sheets של מפעיל המערכת בלבד,
לצורך יחיד: יצירת גיליון נוכחות לכל בחינה ועדכון שורות הנוכחות בו.
המערכת אינה קוראת, מוחקת או משנה קבצים אחרים בחשבון, ואינה ניגשת לחשבונות
של אף משתמש אחר.</p>

<h2>שיתוף עם צדדים שלישיים</h2>
<p>המידע אינו נמכר, אינו מושכר ואינו מועבר לצדדים שלישיים לצורכי פרסום או
כל מטרה מסחרית אחרת. המידע נשאר אצל מפעיל המערכת ואצל הגוף המזמין את הבחינה.</p>

<h2>שמירה ומחיקה</h2>
<p>נתוני נוכחות נשמרים כל עוד הם נדרשים לגוף המזמין את הבחינה. בקשות לעיון
או למחיקה של מידע אישי יטופלו בפנייה לכתובת המייל שבתחתית העמוד.</p>

<h2>אבטחה</h2>
<p>הגישה למערכת מחייבת התחברות. סיסמאות נשמרות מוצפנות. התקשורת עם האתר
מוצפנת ב-HTTPS.</p>
""")


@app.route('/terms')
def terms_of_service():
    """תנאי שימוש — ציבורי, נדרש על ידי גוגל לאישור OAuth."""
    return _LEGAL_PAGE.format(title='תנאי שימוש', body="""
<h2>היקף השימוש</h2>
<p>השימוש ב-URI System מותר לצוות מורשה בלבד, לצורכי ניהול מלאי המחשבים
ורישום נוכחות בבחינות. כל שימוש אחר אסור.</p>

<h2>אחריות המשתמש</h2>
<ul>
  <li>לשמור על סודיות פרטי ההתחברות ולא להעבירם לאחר.</li>
  <li>להזין נתונים נכונים ומדויקים בלבד.</li>
  <li>לא לנסות לעקוף את מנגנוני ההרשאות של המערכת.</li>
  <li>לשמור על סודיות המידע האישי של הנבחנים שנחשף במהלך השימוש.</li>
</ul>

<h2>זמינות השירות</h2>
<p>המערכת מסופקת כפי שהיא. ייתכנו הפסקות שירות לצורכי תחזוקה, עדכון או
בשל תקלות בספקי התשתית. מפעיל המערכת אינו מתחייב לזמינות רציפה.</p>

<h2>הגבלת אחריות</h2>
<p>מפעיל המערכת לא יישא באחריות לנזק עקיף הנובע משימוש במערכת או
מאי-זמינותה. האחריות לנכונות הנתונים המוזנים היא על המשתמש המזין אותם.</p>

<h2>שינויים בתנאים</h2>
<p>תנאים אלה עשויים להתעדכן. המשך השימוש במערכת לאחר עדכון מהווה הסכמה
לתנאים המעודכנים.</p>
""")


@app.route('/login', methods=['GET', 'POST'])
@limiter.limit("10 per minute", methods=["POST"])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()

        # All authentication goes through the database — no hardcoded fallback.
        conn = get_db_connection()
        if conn:
            try:
                cur = get_safe_cursor(conn)
                cur.execute("SELECT * FROM users WHERE username = %s", (username,))
                user = cur.fetchone()
                
                if user:
                    db_pass = (user['password'] or '').strip()
                    is_valid = False
                    needs_migration = False
                    
                    # Check if the stored password is a hash (starts with scrypt:, pbkdf2:, bcrypt$, argon2$)
                    if any(db_pass.startswith(prefix) for prefix in ['scrypt:', 'pbkdf2:', 'bcrypt$', 'argon2$']):
                        if check_password_hash(db_pass, password):
                            is_valid = True
                    else:
                        # Plain text match fallback
                        if db_pass == password:
                            is_valid = True
                            needs_migration = True
                    
                    if is_valid:
                        # If plain text, hash and migrate it now
                        if needs_migration:
                            try:
                                hashed_pass = generate_password_hash(password)
                                cur.execute("UPDATE users SET password = %s WHERE id = %s", (hashed_pass, user['id']))
                                conn.commit()
                                print(f"[MIGRATION] Successfully hashed plain-text password for user: {username}", flush=True)
                            except Exception as em:
                                print(f"[MIGRATION ERROR] Could not hash password for user {username}: {em}", flush=True)
                                
                        # Update last active time
                        cur.execute("UPDATE users SET timestamp = NOW() WHERE id = %s", (user['id'],))
                        conn.commit()
                        
                        session.update({
                            'user': user['username'],
                            'user_id': user.get('id', 999), # In case id is missing
                            'username': user['username'],
                            'role': user['role']
                        })
                        session.permanent = True
                        print(f"[OK] User {username} logged in via DB")
                        next_page = request.args.get('next')
                        if next_page and next_page.startswith('/') and not next_page.startswith('//'):
                            return redirect(next_page)
                        return redirect(url_for('portal'))
                    else:
                        flash("שם משתמש או סיסמה שגויים", "danger")
                else:
                    flash("שם משתמש או סיסמה שגויים", "danger")
            except Exception as e:
                print(f"DB Login Error: {e}")
                flash("שגיאה בהתחברות למסד הנתונים", "danger")
            finally:
                release_db_connection(conn)
        else:
            flash("שגיאת חיבור למסד הנתונים", "danger")
            
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if not ALLOW_SELF_REGISTRATION:
        # אין הרשמה עצמית — משתמשים נוצרים רק מפאנל הניהול
        abort(404)
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        if not username or not password:
            flash("יש להזין שם משתמש וסיסמה", "warning")
            return redirect(url_for('register'))
            
        conn = get_db_connection()
        if conn:
            try:
                cur = get_safe_cursor(conn)
                cur.execute("SELECT * FROM users WHERE username = %s", (username,))
                if cur.fetchone():
                    flash("שם המשתמש כבר קיים במערכת, בחר שם אחר או התחבר", "warning")
                else:
                    hashed_pass = generate_password_hash(password)
                    cur.execute("INSERT INTO users (username, password, role) VALUES (%s, %s, 'technician')", (username, hashed_pass))
                    conn.commit()
                    flash(f"משתמש {username} נוצר בהצלחה! כעת ניתן להתחבר.", "success")
                    cur.close()
                    return redirect(url_for('login'))
                cur.close()
            except Exception as e:
                print(f"DB Register Error: {e}")
                flash("שגיאה ביצירת המשתמש", "danger")
            finally:
                release_db_connection(conn)
        else:
            flash("שגיאת חיבור למסד הנתונים", "danger")
            
    return render_template('register.html')

@app.route('/portal')
@login_required
def portal():
    return render_template('portal.html')

# ── API: שינוי פרטי כניסה של admin_uri ─────────────────────────────────────────
@app.route('/api/change-admin-credentials', methods=['POST'])
@login_required
def api_change_admin_credentials():
    if session.get('role') != 'admin' and session.get('user') not in ('uri', 'admin_uri'):
        return {"success": False, "error": "גישה מותרת לאדמין בלבד"}, 403

    data = request.json or {}
    new_username = (data.get('new_username') or '').strip()
    new_password = (data.get('new_password') or '').strip()

    if not new_username and not new_password:
        return {"success": False, "error": "לא סופק שום שינוי"}

    conn = get_db_connection()
    if not conn:
        return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        current_username = session.get('user', 'admin_uri')

        # בדוק אם קיים רשומה בDB עבור המשתמש הזה
        cur.execute("SELECT id FROM users WHERE username = %s", (current_username,))
        existing = cur.fetchone()

        if existing:
            # עדכן רשומה קיימת
            if new_username and new_password:
                hashed = generate_password_hash(new_password)
                cur.execute("UPDATE users SET username=%s, password=%s WHERE username=%s",
                            (new_username, hashed, current_username))
            elif new_username:
                cur.execute("UPDATE users SET username=%s WHERE username=%s",
                            (new_username, current_username))
            elif new_password:
                hashed = generate_password_hash(new_password)
                cur.execute("UPDATE users SET password=%s WHERE username=%s",
                            (hashed, current_username))
        else:
            # אין רשומה במסד — חובה להזין סיסמה כדי ליצור חשבון אדמין חדש (אין ברירת מחדל)
            if not new_password:
                cur.close()
                return {"success": False, "error": "כדי ליצור רשומת אדמין חדשה חובה להזין סיסמה חדשה"}, 400
            final_username = new_username or current_username
            hashed = generate_password_hash(new_password)
            cur.execute("INSERT INTO users (username, password, role) VALUES (%s, %s, 'admin')",
                        (final_username, hashed))

        conn.commit()
        cur.close()

        msg = []
        if new_username: msg.append(f"שם משתמש שונה ל-{new_username}")
        if new_password: msg.append("סיסמה עודכנה")
        return {"success": True, "message": " | ".join(msg) + ". מתנתק..."}
    except Exception as e:
        conn.rollback()
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/inventory-stats')
@login_required
def api_inventory_stats():
    conn = get_db_connection()
    if not conn:
        return {"error": "db"}, 500
    try:
        cur = get_safe_cursor(conn)
        # ה-SQL נבנה מ-COMPUTER_MODELS כדי שהספירה והמפרט לעולם לא יסתרו זה את זה.
        # כל הערכים הם מספרים שלמים מקבוע פנימי — אין כאן קלט משתמש.
        selects = []
        for _m in COMPUTER_MODELS:
            # שומרים את בדיקת הספרות בתוך כל CASE, כמו בקוד המקורי: אסור להסתמך
            # על כך ש-WHERE יורץ לפני ההשלכה, אחרת ::integer ייפול על ברקוד לא מספרי.
            _conds = " OR ".join(
                f"(barcode ~ '^[0-9]+$' AND barcode::integer BETWEEN {lo} AND {hi})"
                for lo, hi in _m['ranges']
            )
            selects.append(f"COUNT(CASE WHEN ({_conds}) THEN 1 END) AS {_m['key']}")
        cur.execute(
            "SELECT " + ", ".join(selects) +
            " FROM computers WHERE barcode ~ '^[0-9]+$'"
        )
        r = cur.fetchone()
        cur.close()
        items = []
        for _m in COMPUTER_MODELS:
            _count = r[_m['key']] or 0
            _cap = _model_capacity(_m)
            items.append({
                "name": _m['name'], "spec": _m['spec_label'], "count": _count,
                "capacity": _cap, "range": _model_range_label(_m),
                "color": _m['color'], "pct": round(_count / _cap * 100) if _cap else 0,
            })
        return {"items": items}
    except Exception as e:
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/install-cert')
def install_cert():
    """מאפשר לאייפון להוריד ולהתקין את אישור ה-SSL"""
    cert_path = os.path.join(os.path.dirname(__file__), 'server.crt')
    if not os.path.exists(cert_path):
        return "Certificate not found", 404
    return send_file(cert_path, 
                     mimetype='application/x-x509-ca-cert',
                     as_attachment=False,
                     download_name='uri-system.crt')

@app.route('/dashboard')
@app.route('/manage-computers')
@app.route('/computers') # תמיכה בשני השמות
@login_required
def computers():
    search = request.args.get('q', '').strip()
    cage_search = request.args.get('cage_q', '').strip()
    status_filter = request.args.get('status', '').strip()
    page = request.args.get('page', 1, type=int)
    sort = request.args.get('sort', 'scan_time')
    direction = request.args.get('dir', 'desc').lower()
    
    per_page = 100
    offset = (page - 1) * per_page
    
    # Whitelist of allowed sort columns
    allowed_sorts = {
        'barcode': 'barcode',
        'case_number': 'case_number',
        'cage_number': 'cage_number',
        'status': 'status',
        'location': 'location',
        'scan_time': 'scan_time',
        'exam_appeal': 'exam_appeal',
        'specs': 'specs',
        'project': 'project'
    }
    sort_col = allowed_sorts.get(sort, 'scan_time')
    sort_dir = 'ASC' if direction == 'asc' else 'DESC'
    
    conn = get_db_connection()
    if not conn: return "<h1>⚠️ המערכת לא מצליחה להתחבר לענן. בדוק חיבור אינטרנט.</h1>"
    try:
        cur = get_safe_cursor(conn)
        
        # Dashboard Stats
        cur.execute("SELECT COUNT(*) as total FROM computers")
        total_in_db = cur.fetchone()['total']
        
        cur.execute("SELECT status, COUNT(*) as count FROM computers GROUP BY status")
        stats = cur.fetchall()
        stats_dict = {row['status']: row['count'] for row in stats}
        faulty_count = stats_dict.get('תקול', 0)
        # Count computers with no cage assigned
        cur.execute("SELECT COUNT(*) as count FROM computers WHERE (cage_number IS NULL OR TRIM(cage_number) = '') AND (cage_name IS NULL OR TRIM(cage_name) = '')")
        not_in_cage_count = cur.fetchone()['count']
        
        # Base query for computers and count of total matching records
        base_where = " WHERE 1=1"
        params = []
        
        # Free search across multiple fields
        if search:
            norm_search = re.sub(r'^0+(?=\d)', '', search)
            if search.isdigit():
                # Exact barcode match for numeric input (1 = only barcode 1, not 10/11/100)
                base_where += " AND (barcode = %s OR barcode = %s OR case_number ILIKE %s OR location ILIKE %s OR notes ILIKE %s OR exam_appeal ILIKE %s)"
                params.extend([search, norm_search, f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%"])
            else:
                # Partial match for text searches
                base_where += " AND (barcode ILIKE %s OR barcode ILIKE %s OR case_number ILIKE %s OR location ILIKE %s OR notes ILIKE %s OR exam_appeal ILIKE %s)"
                search_val = f"%{search}%"
                norm_val = f"%{norm_search}%"
                params.extend([search_val, norm_val, search_val, search_val, search_val, search_val])
        else:
            norm_search = ''
            
        # Dedicated cage search
        if cage_search:
            base_where += " AND (cage_number = %s OR cage_name ILIKE %s)"
            cs = f"%{cage_search}%"
            params.extend([cage_search, cs])
            
        if status_filter:
            base_where += " AND status = %s"
            params.append(status_filter)

        # Get total matching count for pagination
        cur.execute("SELECT COUNT(*) as cnt FROM computers" + base_where, params)
        total_matching = cur.fetchone()['cnt']
        total_pages = (total_matching + per_page - 1) // per_page
            
        # Query results for current page
        query = "SELECT id, barcode, case_number, cage_name, cage_number, location, status, exam_appeal, specs, project, notes, last_technician, scan_time as last_seen FROM computers"
        query += base_where
        
        # Order by logic — exact barcode match always comes first, then normal sort
        if search:
            exact_priority = " (CASE WHEN barcode = %s OR barcode = %s THEN 0 ELSE 1 END),"
            order_params = [search, norm_search]
        else:
            exact_priority = ""
            order_params = []

        if sort_col == 'scan_time':
            query += f" ORDER BY{exact_priority} {sort_col} {sort_dir} NULLS LAST"
        else:
            query += f" ORDER BY{exact_priority} {sort_col} {sort_dir}"
            
        query += " LIMIT %s OFFSET %s"
        
        cur.execute(query, params + order_params + [per_page, offset])
        computers = cur.fetchall()
        cur.close()
        
        return render_template('computers.html', 
                               computers=computers, 
                               search=search, 
                               status_filter=status_filter,
                               total=total_in_db, 
                               faulty=faulty_count, 
                               not_in_cage=not_in_cage_count,
                               page=page,
                               total_pages=total_pages,
                               total_matching=total_matching,
                               sort=sort,
                               direction=direction,
                               cage_search=cage_search,
                               project_search='',
                               inventory=[])
    except Exception as _e:
        import traceback as _tb
        _err = _tb.format_exc()
        print(f'[COMPUTERS ERROR] {_err}', flush=True)
        return f'<pre style="direction:ltr">COMPUTERS ERROR:\n{_err}</pre>', 500
    finally:
        release_db_connection(conn)

# נתיבים נוספים שנדרשים בטמפלייט base.html
@app.route('/quick-add', methods=['GET'])
@login_required
def quick_add():
    return render_template('batch_add.html')

@app.route('/add-computer', methods=['GET', 'POST'])
@login_required
def add_computer():
    if request.method == 'POST':
        data = request.form
        conn = get_db_connection()
        if not conn: return "DB connection failed", 500
        try:
            cur = get_safe_cursor(conn)
            barcode = re.sub(r'^0+(?=\d)', '', data['barcode'].strip())
            force_update = data.get('force_update', '0') == '1'

            # Check for existing barcode
            cur.execute("SELECT id, cage_number, location FROM computers WHERE barcode = %s", (barcode,))
            existing = cur.fetchone()

            if existing and not force_update:
                existing_cage = existing['cage_number'] or 'לא ידוע'
                flash(f"⚠️ מחשב {barcode} כבר קיים בכלוב {existing_cage}", "warning")
                return render_template('computer_form.html', action='add', computer=None,
                                       existing_barcode=barcode, existing_cage=existing_cage,
                                       existing_id=existing['id'], prefill=data)

            project = data.get('project', '').strip()
            auto_spec = get_auto_spec(barcode)
            specs_val = data.get('specs', '').strip() or auto_spec

            if existing and force_update:
                # עדכן את הרשומה הקיימת
                cur.execute("""
                    UPDATE computers
                    SET case_number=%s, cage_number=%s, status=%s, location=%s,
                        specs=%s, project=%s, notes=%s, last_technician=%s, scan_time=NOW()
                    WHERE id=%s
                """, (data.get('case_number',''), data.get('cage_number',''),
                       data.get('status','תקין'), data.get('location',''),
                       specs_val, project, data.get('notes',''),
                       session.get('username'), existing['id']))
                cur.execute("""
                    INSERT INTO inventory_history (computer_id, technician, change_type, new_value)
                    VALUES (%s, %s, 'Updated via Add Form', %s)
                """, (existing['id'], session.get('username'),
                       f"כלוב שונה ל: {data.get('cage_number','')}" ))
                conn.commit()
                cur.close()
                flash(f"מחשב {barcode} עודכן בהצלחה!", "success")
                return redirect(url_for('computers'))

            cur.execute("""
                INSERT INTO computers (barcode, case_number, cage_number, status, location, specs, project, notes, scan_time, last_technician)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW(), %s)
            """, (barcode, data.get('case_number',''), data.get('cage_number',''),
                  data.get('status','תקין'), data.get('location',''),
                  specs_val, project, data.get('notes',''), session.get('username')))
            conn.commit()
            cur.close()
            flash("מחשב נוסף בהצלחה!", "success")
            return redirect(url_for('computers'))
        except Exception as e:
            conn.rollback()
            flash(f"שגיאה בהוספת מחשב: {e}", "danger")
        finally:
            release_db_connection(conn)

    return render_template('computer_form.html', action='add', computer=None)

@app.route('/edit-computer/<int:cid>', methods=['GET', 'POST'])
@login_required
def edit_computer(cid):
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    try:
        cur = get_safe_cursor(conn)
        if request.method == 'POST':
            data = request.form
            cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
            old_val = cur.fetchone()
            
            project = data.get('project', '').strip()

            cur.execute("""
                UPDATE computers 
                SET case_number=%s, cage_number=%s, status=%s, location=%s, exam_appeal=%s, specs=%s, project=%s, notes=%s, last_technician=%s
                WHERE id=%s
            """, (data.get('case_number',''), data.get('cage_number',''), data.get('status','תקין'), data.get('location',''), data.get('exam_appeal',''), data.get('specs',''), project, data.get('notes',''), session.get('username'), cid))
            
            cur.execute("""
                INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                VALUES (%s, %s, 'Manual Edit', %s, %s)
            """, (
                cid, 
                session.get('username'), 
                json.dumps(dict(old_val), default=str) if old_val else None, 
                json.dumps(dict(data), default=str)
            ))
            
            conn.commit()
            cur.close()
            flash("פרטי המחשב עודכנו!", "success")
            return redirect(url_for('computers'))
            
        cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
        computer = cur.fetchone()
        cur.close()
        return render_template('computer_form.html', action='edit', computer=computer)
    finally:
        release_db_connection(conn)

@app.route('/delete-computer/<int:cid>', methods=['POST'])
@login_required
def delete_computer(cid):
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    try:
        cur = get_safe_cursor(conn)
        # Only admin_uri can hard delete
        if session.get('user') == 'admin_uri':
            cur.execute("DELETE FROM computers WHERE id = %s", (cid,))
            flash("המחשב נמחק מהמערכת סופית (admin_uri)", "warning")
        else:
            cur.execute("UPDATE computers SET status = 'ממתין למחיקה' WHERE id = %s", (cid,))
            flash("הבקשה למחיקת המחשב הועברה לאישור מנהל העל (admin_uri)", "info")
            
        conn.commit()
        cur.close()
    except Exception as e:
        flash(f"שגיאה במחיקה: {e}", "danger")
    finally:
        release_db_connection(conn)
    return redirect(url_for('computers'))

@app.route('/scanner')
@login_required
def scanner():
    return render_template('scanner.html')

@app.route('/exam')
@login_required
def exam_page():
    conn = get_db_connection()
    if not conn: return redirect(url_for('dashboard'))
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT * FROM computers WHERE exam_appeal IS NOT NULL AND TRIM(exam_appeal) != '' AND LOWER(TRIM(exam_appeal)) != 'none'")
        computers = cur.fetchall()
        cur.close()
        return render_template('exam.html', computers=computers)
    finally:
        release_db_connection(conn)

@app.route('/history')
@login_required
def history_page():
    conn = get_db_connection()
    if not conn: return redirect(url_for('dashboard'))
    try:
        cur = get_safe_cursor(conn)
        cur.execute("""
            SELECT h.*, c.barcode 
            FROM inventory_history h
            LEFT JOIN computers c ON h.computer_id = c.id
            ORDER BY h.timestamp DESC
            LIMIT 100
        """)
        history = cur.fetchall()
        cur.close()
        
        # Safe processing for templates
        processed_history = []
        for h in history:
            h_dict = dict(h)
            ts = h_dict['timestamp']
            if ts and isinstance(ts, str):
                try:
                    # SQLite default format
                    h_dict['timestamp'] = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
                except:
                    pass
            processed_history.append(h_dict)

        return render_template('history.html', history=processed_history)
    finally:
        release_db_connection(conn)

# ── SHARED SCAN SESSION ────────────────────────────────────────────────────────
# מאפשר לטכנאי אחד להגדיר הגדרות, והשאר מצטרפים עם קוד קצר
import random
import string
import time as _time

_shared_sessions = {}  # { code: { settings, owner, created_at } }

def _generate_code():
    """מייצר קוד קצר ייחודי בן 4 תווים"""
    while True:
        code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
        if code not in _shared_sessions:
            return code

def _cleanup_old_sessions():
    """מוחק sessions ישנות (מעל 8 שעות)"""
    now = _time.time()
    expired = [c for c, s in _shared_sessions.items() if now - s['created_at'] > 28800]
    for c in expired:
        del _shared_sessions[c]

@app.route('/api/shared-session/create', methods=['POST'])
@login_required
def create_shared_session():
    """יוצר session חדש עם הגדרות — מחזיר קוד קצר לשיתוף"""
    _cleanup_old_sessions()
    data = request.json or {}
    settings = {
        'location': data.get('location', ''),
        'cage': data.get('cage', ''),
        'status': data.get('status', 'תקין'),
        'exam': data.get('exam', ''),
        'specs': data.get('specs', ''),
        'project': data.get('project', ''),
        'ministry': data.get('ministry', ''),
    }
    code = _generate_code()
    _shared_sessions[code] = {
        'settings': settings,
        'owner': session.get('username', ''),
        'created_at': _time.time()
    }
    print(f"[SharedSession] Created session {code} by {session.get('username')}")
    return {'success': True, 'code': code}

@app.route('/api/shared-session/<code>', methods=['GET'])
@login_required
def get_shared_session(code):
    """מחזיר את הגדרות ה-session לפי קוד"""
    sess = _shared_sessions.get(code.upper())
    if not sess:
        return {'success': False, 'error': 'קוד לא נמצא או פג תוקף'}, 404
    return {'success': True, 'settings': sess['settings'], 'owner': sess['owner']}

@app.route('/api/shared-session/<code>/update', methods=['POST'])
@login_required
def update_shared_session(code):
    """מעדכן הגדרות session קיים (רק הבעלים)"""
    sess = _shared_sessions.get(code.upper())
    if not sess:
        return {'success': False, 'error': 'קוד לא נמצא'}, 404
    if sess['owner'] != session.get('username'):
        return {'success': False, 'error': 'רק הבעלים יכול לעדכן'}, 403
    data = request.json or {}
    for key in ['location', 'cage', 'status', 'exam', 'specs', 'project', 'ministry']:
        if key in data:
            sess['settings'][key] = data[key]
    return {'success': True}

@app.route('/api/shared-session/<code>/close', methods=['POST'])
@login_required
def close_shared_session(code):
    """סוגר session"""
    code = code.upper()
    if code in _shared_sessions:
        if _shared_sessions[code]['owner'] == session.get('username'):
            del _shared_sessions[code]
    return {'success': True}
# ─────────────────────────────────────────────────────────────────────────────

@app.route('/api/process-scan', methods=['POST'])
@login_required
def process_scan():
    barcode = request.json.get('barcode', '').strip()
    if not barcode:
        return {"error": "No barcode provided"}, 400
    barcode = re.sub(r'^0+(?=\d)', '', barcode)

    # ── חסימת QR קודים של נבחנים ────────────────────────────────────────────────
    # מזהים: מכיל | (pipe) עם נתוני בחינה, מילות מפתח של מבחן/ערעור/נבחן
    EXAM_KEYWORDS = ['במבחן', 'ערעור', 'נבחן', 'בחינה', 'תעודה']
    is_exam_qr = (
        '|' in barcode and len(barcode) > 20  # מבנה QR נבחן: data|data|data
        or any(kw in barcode for kw in EXAM_KEYWORDS)
    )
    if is_exam_qr:
        print(f"[BLOCKED] Examinee QR rejected: {barcode[:30]}...")
        return {
            "error": "❌ QR זה שייך לנבחן — לא ניתן לסרוק אותו כמחשב!",
            "blocked": True
        }, 400
    # ─────────────────────────────────────────────────────────────────────────────

    # ── ולידציה: ברקוד חייב להיות מספרי בלבד ובטווח 1-9999 ────────────────────────
    if not barcode.isdigit() or not (1 <= int(barcode) <= 9999):
        print(f"[BLOCKED] Invalid barcode rejected: '{barcode}'")
        return {
            "error": "❌ הברקוד דפוק",
            "blocked": True
        }, 400
    # ─────────────────────────────────────────────────────────────────────────────

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        # Check if computer exists
        cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
        computer = cur.fetchone()
        
        if computer:
            # Update the existing record instead of inserting a duplicate
            auto_spec = get_auto_spec(barcode)
            cur.execute("""
                UPDATE computers 
                SET scan_time = NOW(), last_technician = %s
                    {spec_part}
                    {case_part}
                WHERE id = %s
            """.format(
                spec_part=", specs = %s" if auto_spec else "",
                case_part=", case_number = %s" if not (computer.get('case_number') or '').strip() else ""
            ),
                (session.get('username'),)
                + ((auto_spec,) if auto_spec else ())
                + ((barcode,) if not (computer.get('case_number') or '').strip() else ())
                + (computer['id'],))
            
            # Fetch the updated record (SQLite doesn't support RETURNING)
            cur.execute("SELECT * FROM computers WHERE id = %s", (computer['id'],))
            new_computer = dict(cur.fetchone())
            
            # Fetch previous technician and time (skip current scan = LIMIT 1 OFFSET 1)
            cur.execute("SELECT technician, timestamp FROM inventory_history WHERE computer_id = %s ORDER BY timestamp DESC LIMIT 1 OFFSET 1", (new_computer['id'],))
            hist = cur.fetchone()
            new_computer['last_technician'] = hist['technician'] if hist and hist['technician'] else ""
            
            ts = hist['timestamp'] if hist else None
            if ts and not isinstance(ts, str):
                new_computer['last_scan_time'] = ts.strftime("%d/%m/%Y %H:%M")
            else:
                new_computer['last_scan_time'] = str(ts) if ts else ''
            
            # Serialize scan_time for JSON
            if new_computer.get('scan_time') and not isinstance(new_computer['scan_time'], str):
                new_computer['scan_time'] = new_computer['scan_time'].strftime("%d/%m/%Y %H:%M")
            
            conn.commit()
            cur.close()
            # Trigger Google Sheets sync in background
            trigger_debounced_sync()
            
            # מחזירים את המידע הקיים כדי שהטופס יתמלא נכון
            return {"exists": True, "computer": new_computer}
        else:
            # Create completely new record
            auto_spec = get_auto_spec(barcode)
            cur.execute("""
                INSERT INTO computers (barcode, case_number, status, scan_time, specs, notes, last_technician) 
                VALUES (%s, %s, 'תקין', NOW(), %s, %s, %s) 
            """, (barcode, barcode, auto_spec or None, None, session.get('username')))
            
            last_id = cur.lastrowid if hasattr(cur, 'lastrowid') else None
            if last_id:
                cur.execute("SELECT * FROM computers WHERE id = %s", (last_id,))
                new_computer = dict(cur.fetchone())
            else:
                # Fallback if lastrowid fails
                cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
                new_computer = dict(cur.fetchone())

            conn.commit()
            cur.close()
            # Trigger Google Sheets sync in background
            trigger_debounced_sync()
            
            return {"exists": False, "computer": new_computer}
            
    except Exception as e:
        print(f"Error in process_scan: {e}")
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/update-computer', methods=['POST'])
@login_required
def api_update_computer():
    data = request.json
    cid = data.get('id')
    if not cid: return {"error": "No ID provided"}, 400

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        # Get old values for history
        cur.execute("SELECT * FROM computers WHERE id = %s", (cid,))
        old_val = cur.fetchone()
        
        # Update
        updates = []
        params = []
        for key in ['case_number', 'cage_number', 'status', 'location', 'exam_appeal', 'specs', 'project']:
            if key in data:
                val = data[key]
                updates.append(f"{key} = %s")
                params.append(val)

        # Only update notes if a non-empty value was explicitly sent
        notes_val = data.get('notes')
        if isinstance(notes_val, str) and notes_val.strip():
            updates.append("notes = %s")
            params.append(notes_val.strip())

        
        # Always update last_technician on scan update
        updates.append("last_technician = %s")
        params.append(session.get('username'))

        if updates:
            params.append(cid)
            cur.execute(f"UPDATE computers SET {', '.join(updates)} WHERE id = %s", params)
            
            # Record history
            cur.execute("""
                INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                VALUES (%s, %s, 'Update via Scan', %s, %s)
            """, (
                cid, 
                session.get('username'), 
                json.dumps(dict(old_val), default=str) if old_val else None, 
                json.dumps(data, default=str)
            ))
            
            conn.commit()
            cur.close()
            # Trigger Google Sheets sync in background
            trigger_debounced_sync()
            
            return {"success": True}
        return {"success": False, "message": "No fields to update"}
        
    except Exception as e:
        print(f"Error in api_update_computer: {e}")
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

# ── AI ASSISTANT (GEMINI) ────────────────────────────────────────────────────
@app.route('/api/ai-chat', methods=['POST'])
@login_required
def api_ai_chat():
    data = request.json
    user_msg = data.get('message', '').strip()
    if not user_msg:
        return {"error": "No message provided"}, 400

    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        # Fetch stats for context
        cur.execute("SELECT COUNT(*) as total FROM computers")
        total = cur.fetchone()['total']
        
        cur.execute("SELECT status, COUNT(*) as count FROM computers GROUP BY status")
        status_stats = cur.fetchall()
        status_desc = ", ".join([f"{row['status']}: {row['count']}" for row in status_stats])
        
        # Build system context
        system_prompt = f"""
        אתה עוזר ה-AI של מערכת URI לניהול מלאי מחשבים. 
        הנתונים הנוכחיים במערכת הם:
        - סה"כ מחשבים: {total}
        - סטטוסים: {status_desc}
        
        ענה למשתמש בעברית בצורה עוזרת, מקצועית וקצרה. 
        אם המשתמש שואל על המצב, השתמש בנתונים שלעיל.
        משתמש נוכחי: {session.get('username')}
        """
        
        # Try Claude Sonnet first, fallback to Gemini
        anthropic_key = os.getenv('ANTHROPIC_API_KEY')
        if anthropic_key:
            try:
                import anthropic
                client = anthropic.Anthropic(api_key=anthropic_key)
                msg = client.messages.create(
                    model='claude-sonnet-4-5',
                    max_tokens=1024,
                    system=system_prompt,
                    messages=[{'role': 'user', 'content': user_msg}]
                )
                return {'response': msg.content[0].text}
            except Exception as anthropic_err:
                print(f"[AI] Anthropic failed ({anthropic_err}), falling back to Gemini")

        # Fallback: Gemini
        if not genai_client:
            return {'response': 'מערכת ה-AI אינה זמינה כרגע (מפתח API חסר)'}, 503
        if getattr(genai_client, '_legacy_mode', False):
            # Old SDK
            response = genai_client.generate_content([system_prompt, user_msg])
        else:
            response = genai_client.models.generate_content(
                model='gemini-2.0-flash',
                contents=[system_prompt, user_msg]
            )
        return {'response': response.text}
        
    except Exception as e:
        print(f"AI Error: {e}")
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

# ── API: Batch Operations ──────────────────────────────────────────────────
@app.route('/api/batch-update', methods=['POST'])
@login_required
def api_batch_update():
    data = request.json
    ids = data.get('ids', [])
    updates = data.get('updates', {})
    
    if not ids or not updates:
        return {"success": False, "error": "Missing ids or updates"}, 400
        
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        
        set_clauses = []
        params = []
        for key in ['location', 'cage_number', 'cage_name', 'status', 'exam_appeal', 'specs', 'project', 'ministry', 'notes']:
            if key in updates:
                set_clauses.append(f"{key} = %s")
                params.append(updates[key])
                
        if not set_clauses:
            return {"success": False, "error": "No valid fields provided"}, 400
            
        params.extend(ids)
        placeholders = ','.join(['%s'] * len(ids))
        
        query = f"UPDATE computers SET {', '.join(set_clauses)} WHERE id IN ({placeholders})"
        cur.execute(query, params)
        
        # Log to history for each (simplified to avoid mass select first, assuming identical change)
        for cid in ids:
            cur.execute("""
                INSERT INTO inventory_history (computer_id, technician, change_type, new_value)
                VALUES (%s, %s, 'Batch Update', %s)
            """, (cid, session.get('username'), json.dumps(updates, default=str)))
            
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error in batch-update: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/batch-delete', methods=['POST'])
@login_required
def api_batch_delete():
    # Only admin should be able to trigger this in the UI, verify on server too
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        return {"success": False, "error": "Unauthorized"}, 403
        
    data = request.json
    ids = data.get('ids', [])
    if not ids: return {"success": False, "error": "No ids provided"}, 400
    
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        placeholders = ','.join(['%s'] * len(ids))
        cur.execute(f"DELETE FROM computers WHERE id IN ({placeholders})", ids)
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error in batch-delete: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

# â”€â”€ API: Google Sheets Sync â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@app.route('/api/sync-to-sheets', methods=['POST'])
@login_required
def api_sync_to_sheets():
    """
    Manual trigger for Google Sheets sync.
    """
    # Only admin or experienced technicians should sync
    # For now, allow all logged in users as requested "connect table"
    success, message = sync_inventory_to_sheets()
    if success:
        return {"success": True, "message": message}
    else:
        return {"success": False, "error": message}, 500

@app.route('/api/find-duplicates', methods=['GET'])
@login_required
def api_find_duplicates():
    conn = get_db_connection()
    if not conn:
        return jsonify({"success": False, "error": "DB error"}), 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("""
            SELECT barcode, COUNT(*) as cnt
            FROM computers
            GROUP BY barcode
            HAVING COUNT(*) > 1
            ORDER BY cnt DESC
        """)
        rows = cur.fetchall()
        duplicates = []
        for r in rows:
            if hasattr(r, 'keys'):
                duplicates.append({'barcode': r['barcode'], 'count': r['cnt']})
            else:
                duplicates.append({'barcode': r[0], 'count': r[1]})
        cur.close()
        return jsonify({"success": True, "duplicates": duplicates, "total": len(duplicates)})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        release_db_connection(conn)

@app.route('/api/delete-duplicates', methods=['POST'])
@login_required
def api_delete_duplicates():
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        return jsonify({"success": False, "error": "Admin only"}), 403
    conn = get_db_connection()
    if not conn:
        return jsonify({"success": False, "error": "DB error"}), 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("""
            DELETE FROM computers
            WHERE id NOT IN (
                SELECT DISTINCT ON (barcode) id
                FROM computers
                ORDER BY barcode, scan_time DESC NULLS LAST
            )
        """)
        deleted = cur.rowcount
        conn.commit()
        cur.close()
        return jsonify({"success": True, "deleted": deleted, "message": f"נמחקו {deleted} כפולים"})
    except Exception as e:
        conn.rollback()
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        release_db_connection(conn)


# -- API: import-from-sheets removed --
@app.route('/api/import-from-sheets', methods=['POST'])
@login_required
def api_import_from_sheets():
    return {"success": False, "message": "removed"}, 410

# â”€â”€ API: אישור מחיקה סופית (רק admin) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@app.route('/api/approve-sheets-delete', methods=['POST'])
@login_required
def api_approve_sheets_delete():
    """רק admin_uri יכול לאשר מחיקה סופית של מחשבים שמסומנים כ-sheets_delete_request"""
    if session.get('user') != 'admin_uri' and session.get('role') != 'admin':
        return {"success": False, "error": "גישה מותרת לאדמין בלבד"}, 403

    data = request.json or {}
    action = data.get('action')  # 'approve' or 'cancel'
    computer_ids = data.get('ids', [])

    if not computer_ids:
        # אם לא נשלחו IDs â€” פעל על כולם המסומנים
        conn = get_db_connection()
        if not conn:
            return {"success": False, "error": "DB connection failed"}, 500
        try:
            cur = get_safe_cursor(conn)
            cur.execute("SELECT id FROM computers WHERE sheets_delete_request = TRUE")
            rows = cur.fetchall()
            computer_ids = [r['id'] for r in rows]
            cur.close()
        finally:
            release_db_connection(conn)

    if not computer_ids:
        return {"success": True, "message": "אין מחשבים הממתינים לאישור מחיקה", "count": 0}

    conn = get_db_connection()
    if not conn:
        return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        placeholders = ','.join(['%s'] * len(computer_ids))

        if action == 'approve':
            # מחיקה סופית
            cur.execute(f"DELETE FROM computers WHERE id IN ({placeholders}) AND sheets_delete_request = TRUE", computer_ids)
            conn.commit()
            count = cur.rowcount
            cur.close()
            return {"success": True, "message": f"נמחקו {count} מחשבים לצמיתות", "count": count}
        else:
            # ביטול סימון â€” המחשב נשאר במסד
            cur.execute(f"UPDATE computers SET sheets_delete_request = FALSE WHERE id IN ({placeholders})", computer_ids)
            conn.commit()
            count = cur.rowcount
            cur.close()
            return {"success": True, "message": f"בוטל סימון המחיקה עבור {count} מחשבים", "count": count}
    except Exception as e:
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

# â”€â”€ API: שליפת מידע כלוב â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@app.route('/api/cage/<cage_id>', methods=['GET'])
@login_required
def api_get_cage(cage_id):
    """מחזיר מידע על כלוב + רשימת המחשבים בו"""
    conn = get_db_connection()
    if not conn: return {"error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)

        # שליפת פרטי הכלוב (אם קיים בטבלת cages)
        cur.execute("SELECT * FROM cages WHERE cage_id = %s", (cage_id,))
        cage = cur.fetchone()

        # שליפת מחשבים בכלוב זה
        cur.execute("""
            SELECT id, barcode, status, location, specs, scan_time, notes
            FROM computers
            WHERE cage_number = %s OR cage_name = %s
            ORDER BY scan_time DESC NULLS LAST
        """, (cage_id, cage_id))
        computers_in_cage = cur.fetchall()

        # סטטיסטיקות
        total = len(computers_in_cage)
        status_counts = {}
        for c in computers_in_cage:
            s_val = c.get('status')
            s = s_val if s_val is not None else 'לא ידוע'
            status_counts[s] = status_counts.get(s, 0) + 1  # type: ignore

        cur.close()
        return {
            "cage": dict(cage) if cage else {"cage_id": cage_id, "name": f"כלוב {cage_id}"},
            "computers": [dict(c) for c in computers_in_cage],
            "total": total,
            "status_counts": status_counts
        }
    except Exception as e:
        print(f"Error in api_get_cage: {e}")
        return {"error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/cage_manage/<cage_id>')
@login_required
def cage_manage(cage_id):
    normalized = normalize_cage_id(cage_id)
    if normalized != cage_id:
        return redirect(url_for('cage_manage', cage_id=normalized))
    """UI דף ניהול כלוב"""
    conn = get_db_connection()
    if not conn:
        flash("שגיאת חיבור למסד הנתונים", "danger")
        return redirect(url_for('computers'))
        
    try:
        cur = get_safe_cursor(conn)
        cur.execute("""
            SELECT barcode, status, specs
            FROM computers
            WHERE cage_number = %s OR cage_name = %s
            ORDER BY scan_time DESC NULLS LAST
        """, (cage_id, cage_id))
        computers_in_cage = cur.fetchall()
        
        total = len(computers_in_cage)
        stats = {'Dell': 0, 'Lenovo': 0, 'HP': 0}
        
        enriched_computers = []
        for c in computers_in_cage:
            # נסה להסיק יצרן מהברקוד או מהמפרט
            bc = str(c.get('barcode', ''))
            specs = c.get('specs', '') or ''
            mfg = 'אחר'

            # מספר המחשב הוא הקובע. קודם היה נבדק כאן טקסט המפרט לפני הברקוד,
            # ולכן HP 2023 (המפרט שלו מכיל "i7-1195G7") נספר כ-Dell.
            model = get_model_for_barcode(bc)
            if model:
                mfg = model['manufacturer']
            elif 'Dell' in specs:
                mfg = 'Dell'
            elif 'HP' in specs:
                mfg = 'HP'
            elif 'Lenovo' in specs:
                mfg = 'Lenovo'
            if mfg in stats:
                stats[mfg] += 1
            
            enriched_computers.append({
                'barcode': bc,
                'status': c.get('status', ''),
                'mfg': mfg
            })
            
        cur.close()
        return render_template('cage_manage.html', cage_id=cage_id, computers=enriched_computers, total=total, stats=stats)
    except Exception as e:
        print(f"Error in cage_manage: {e}")
        flash("שגיאה בטעינת נתוני כלוב", "danger")
        return redirect(url_for('computers'))
    finally:
        release_db_connection(conn)

@app.route('/api/cage/<cage_id>/add', methods=['POST'])
@login_required
def api_cage_add(cage_id):
    cage_id = normalize_cage_id(cage_id)
    barcode = request.form.get('barcode', '').strip()
    confirm_move = request.form.get('confirm_move', '')
    from_edit = request.form.get('from_edit', '')
    session_count = request.form.get('session_count', '0')
    
    try:
        session_count = int(session_count)
    except:
        session_count = 0
        
    def redirect_back(**kwargs):
        if from_edit == '1':
            return redirect(url_for('cage_edit', cage_id=cage_id, session_count=session_count, **kwargs))
        return redirect(url_for('cage_manage', cage_id=cage_id, **kwargs))

    if not barcode:
        flash("נא להזין ברקוד", "warning")
        return redirect_back()
        
    conn = get_db_connection()
    if not conn:
        flash("שגיאת חיבור למסד הנתונים", "danger")
        return redirect_back()
        
    try:
        cur = get_safe_cursor(conn)
        # Check if computer exists
        cur.execute("SELECT id, cage_number FROM computers WHERE barcode = %s", (barcode,))
        existing = cur.fetchone()
        
        if existing:
            old_cage = existing['cage_number']
            if old_cage and str(old_cage) != str(cage_id) and confirm_move != 'yes':
                # Needs confirmation
                return redirect_back(confirm_barcode=barcode, old_cage=old_cage)
                
            # Update its cage
            cur.execute("UPDATE computers SET cage_number = %s, cage_name = %s, status = 'בכלוב', scan_time = NOW() WHERE barcode = %s", 
                        (cage_id, cage_id, barcode))
            if confirm_move != 'yes':
                session_count += 1
            flash(f"המחשב {barcode} שויך לכלוב {cage_id} בהצלחה.", "success")
            last_added_bc = barcode
        else:
            # Create new computer and assign to cage
            cur.execute("""
                INSERT INTO computers (barcode, cage_number, cage_name, status, scan_time, last_technician)
                VALUES (%s, %s, %s, 'בכלוב', NOW(), %s)
            """, (barcode, cage_id, cage_id, session.get('username', '')))
            if confirm_move != 'yes':
                session_count += 1
            flash(f"המחשב {barcode} נוצר במערכת ושויך לכלוב {cage_id}.", "success")
            last_added_bc = barcode
            
        conn.commit()
        trigger_debounced_sync()
        cur.close()
    except Exception as e:
        print(f"Error in api_cage_add: {e}")
        flash("שגיאה בהוספת המחשב", "danger")
        last_added_bc = ""

    finally:
        release_db_connection(conn)
        
    return redirect_back(last_added=last_added_bc)

@app.route('/api/cage/<cage_id>/batch_add', methods=['POST'])
@login_required
def api_cage_batch_add(cage_id):
    cage_id = normalize_cage_id(cage_id)
    data = request.get_json()
    if not data or 'barcodes' not in data:
        return jsonify({"success": False, "error": "Invalid payload"}), 400
        
    barcodes = data['barcodes']
    if not isinstance(barcodes, list):
        return jsonify({"success": False, "error": "Barcodes must be a list"}), 400
        
    conn = get_db_connection()
    if not conn:
        return jsonify({"success": False, "error": "DB connection failed"}), 500
        
    results = {"success": True, "updated": 0, "new": 0, "transferred": 0, "details": []}
    technician = session.get('username', 'unknown')
    
    try:
        cur = get_safe_cursor(conn)
        
        for barcode in barcodes:
            barcode = str(barcode).strip()
            if not barcode: continue
            
            cur.execute("SELECT * FROM computers WHERE barcode = %s", (barcode,))
            existing = cur.fetchone()
            
            if existing:
                old_val = dict(existing)
                prev_cage = old_val.get('cage_number', '')
                is_transfer = (prev_cage and str(prev_cage) != str(cage_id))
                
                cur.execute("UPDATE computers SET cage_number = %s, cage_name = %s, status = 'בכלוב', scan_time = NOW(), last_technician = %s WHERE barcode = %s", 
                            (cage_id, cage_id, technician, barcode))
                            
                cur.execute("""
                    INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                    VALUES (%s, %s, 'Batch Pack', %s, %s)
                """, (
                    old_val['id'],
                    technician,
                    json.dumps(old_val, default=str),
                    json.dumps({"cage_number": cage_id, "cage_name": cage_id}, default=str)
                ))
                
                if is_transfer:
                    results["transferred"] += 1
                    results["details"].append({"barcode": barcode, "status": "transferred"})
                else:
                    results["updated"] += 1
                    results["details"].append({"barcode": barcode, "status": "updated"})
                    
            else:
                cur.execute("""
                    INSERT INTO computers (barcode, cage_number, cage_name, status, scan_time, last_technician)
                    VALUES (%s, %s, %s, 'בכלוב', NOW(), %s) RETURNING id
                """, (barcode, cage_id, cage_id, technician))
                new_id = cur.fetchone()['id']
                
                cur.execute("""
                    INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                    VALUES (%s, %s, 'Batch Pack (New)', %s, %s)
                """, (
                    new_id,
                    technician,
                    None,
                    json.dumps({"cage_number": cage_id, "cage_name": cage_id, "status": "בכלוב"}, default=str)
                ))
                
                results["new"] += 1
                results["details"].append({"barcode": barcode, "status": "new"})
                
        conn.commit()
        trigger_debounced_sync()
        cur.close()
    except Exception as e:
        print(f"Error in api_cage_batch_add: {e}")
        if conn: conn.rollback()
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        release_db_connection(conn)
        
    return jsonify(results)

def normalize_cage_id(cid):
    if not cid: return cid
    cid = cid.strip()
    if cid.isdigit():
        return str(int(cid))
    return cid

@app.route('/cage_entry', methods=['GET', 'POST'])
@login_required
def cage_entry():
    if request.method == 'POST':
        cage_id = request.form.get('cage_id', '').strip()
        cage_id = normalize_cage_id(cage_id)
        if cage_id:
            return redirect(url_for('cage_manage', cage_id=cage_id))
        else:
            flash("נא להזין מספר כלוב", "warning")
    return render_template('cage_entry.html')

@app.route('/cage_manage/<cage_id>/edit')
@login_required
def cage_edit(cage_id):
    normalized = normalize_cage_id(cage_id)
    if normalized != cage_id:
        session_count = request.args.get('session_count', '0')
        return redirect(url_for('cage_edit', cage_id=normalized, session_count=session_count))
    session_count = request.args.get('session_count', '0')
    try:
        session_count = int(session_count)
    except:
        session_count = 0
    return render_template('cage_edit.html', cage_id=cage_id, session_count=session_count)

@app.route('/api/cage/<cage_id>/verify', methods=['POST'])
@login_required
def api_cage_verify(cage_id):
    # This just marks the cage as verified by user, then returns them to portal
    flash(f"כלוב {cage_id} אומת בהצלחה! תודה.", "success")
    return redirect(url_for('portal'))

@app.route('/api/cage/<cage_id>/mark_returned', methods=['POST'])
@login_required
def api_cage_mark_returned(cage_id):
    # get all barcodes in the form, and the returned ones
    returned_barcodes = request.form.getlist('returned_barcodes')
    all_barcodes = request.form.getlist('all_barcodes')
    
    if not all_barcodes:
        return redirect(url_for('cage_manage', cage_id=cage_id))
        
    conn = get_db_connection()
    if not conn:
        flash("שגיאת חיבור למסד הנתונים", "danger")
        return redirect(url_for('cage_manage', cage_id=cage_id))
        
    try:
        cur = get_safe_cursor(conn)
        updated_count = 0
        
        for bc in all_barcodes:
            if bc in returned_barcodes:
                status = 'חזר'
            else:
                status = 'חסר'  # Or we can just leave it as it was if we don't want to mark missing
                
            cur.execute("UPDATE computers SET status = %s WHERE barcode = %s", (status, bc))
            updated_count += 1
            
        conn.commit()
        trigger_debounced_sync()
        cur.close()
        flash(f"עודכן סטטוס עבור {updated_count} מחשבים.", "success")
    except Exception as e:
        print(f"Error in api_cage_mark_returned: {e}")
        flash("שגיאה בעדכון הסטטוס", "danger")
    finally:
        release_db_connection(conn)
        
    return redirect(url_for('cage_manage', cage_id=cage_id))

# â”€â”€ API: סריקה מהירה (ללא חלון) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@app.route('/api/fast-scan', methods=['POST'])
@login_required
def api_fast_scan():
    """
    סריקה מהירה - מעדכן מחשב ישירות ללא תצוגת UI מפורטת.
    user שולח: barcode, location, cage_number, cage_name, status
    """
    data = request.json
    barcode = data.get('barcode', '').strip()
    if not barcode:
        return {"success": False, "error": "ברקוד חסר"}, 400
    barcode = re.sub(r'^0+(?=\d)', '', barcode)

    location   = data.get('location', '')
    cage_number = data.get('cage_number', '')
    cage_name  = data.get('cage_name', cage_number)
    status     = data.get('status', 'תקין')
    specs      = data.get('specs', '') or get_auto_spec(barcode)
    project    = data.get('project', '')
    ministry   = data.get('ministry', '')
    technician = session.get('username', 'unknown')

    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
        computer = cur.fetchone()

        old_val = dict(computer) if computer else None
        last_technician = "לא ידוע"
        notes_val = ""

        if old_val:
            notes_val = old_val.get('notes') or ""
            fields_to_update = []
            params = []
            
            for key in ['location', 'cage_number', 'cage_name', 'status', 'specs', 'project', 'ministry', 'exam_appeal']:
                if key in data:
                    fields_to_update.append(f"{key} = %s")
                    params.append(data[key])
            
            # Special auto-ministry logic
            if 'project' in data and 'ministry' not in data:
                project_name = data['project']
                ministry = ''
                if 'רופאי' in project_name or 'שיניים' in project_name: ministry = 'משרד הבריאות'
                elif 'משפטים' in project_name: ministry = 'משרד המשפטים'
                elif 'עבודה' in project_name: ministry = 'משרד העבודה'
                elif 'חינוך' in project_name: ministry = 'משרד החינוך'
                fields_to_update.append("ministry = %s")
                params.append(ministry)
                
            fields_to_update.extend(["scan_time = NOW()", "last_technician = %s"])
            params.extend([technician, old_val['id']])
            
            query = f"UPDATE computers SET {', '.join(fields_to_update)} WHERE id = %s"
            # RETURNING * is not supported in SQLite; fetch separately
            cur.execute(query, params)
            cur.execute("SELECT * FROM computers WHERE id = %s", (old_val['id'],))
            new_computer = cur.fetchone()
            
            cur.execute("SELECT technician, timestamp FROM inventory_history WHERE computer_id = %s ORDER BY timestamp DESC LIMIT 1", (new_computer['id'],))
            hist = cur.fetchone()
            if hist and hist['technician']:
                last_technician = hist['technician']
                last_scan_time = hist['timestamp'].strftime("%d/%m/%Y %H:%M") if hist['timestamp'] else ""
        else:
            fields = ['barcode', 'scan_time', 'last_technician']
            params = [barcode, technician]
            
            for key in ['location', 'cage_number', 'cage_name', 'status', 'specs', 'project', 'ministry', 'exam_appeal']:
                if key in data:
                    fields.append(key)
                    params.append(data[key])
                    
            if 'project' in data and 'ministry' not in data:
                fields.append('ministry')
                project_name = data['project']
                ministry = ''
                if 'רופאי' in project_name or 'שיניים' in project_name: ministry = 'משרד הבריאות'
                elif 'משפטים' in project_name: ministry = 'משרד המשפטים'
                elif 'עבודה' in project_name: ministry = 'משרד העבודה'
                elif 'חינוך' in project_name: ministry = 'משרד החינוך'
                params.append(ministry)
                
            if 'status' not in data:
                fields.append('status')
                params.append('תקין')
                
            placeholders = ', '.join(['%s'] * len(fields))
            query = f"INSERT INTO computers ({', '.join(fields)}) VALUES ({placeholders})"
            cur.execute(query, params)
            # RETURNING * is not supported in SQLite; use lastrowid fallback
            last_id = cur.lastrowid if hasattr(cur, 'lastrowid') else None
            if last_id:
                cur.execute("SELECT * FROM computers WHERE id = %s", (last_id,))
            else:
                cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
            new_computer = cur.fetchone()

        cur.execute("""
            INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
            VALUES (%s, %s, 'Fast Scan', %s, %s)
        """, (
            new_computer['id'],
            technician,
            json.dumps(old_val, default=str) if old_val else None,
            json.dumps(data, default=str)
        ))

        conn.commit()
        cur.close()
        
        # טריגר לסנכרון אוטומטי (מתוזמן) עבור סריקה מהירה
        trigger_debounced_sync()
        
        return {
            "success": True, 
            "barcode": barcode, 
            "is_new": old_val is None,
            "previous_cage": old_val.get('cage_number', '') if old_val else "",
            "notes": notes_val
        }

    except Exception as e:
        print(f"Error in api_fast_scan: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/admin_approve_delete', methods=['POST'])
@login_required
def admin_approve_delete():
    # Only admin_uri can approve deletes
    if session.get('user') != 'admin_uri':
        return {"success": False, "error": "Unauthorized: Only admin_uri can perform this action"}, 403
        
    data = request.json
    barcode = data.get('barcode')
    action = data.get('action')
    
    if not barcode or not action:
        return {"success": False, "error": "Missing parameters"}, 400
        
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB Error"}, 500
    try:
        cur = get_safe_cursor(conn)
        if action == 'hard_delete':
            cur.execute("DELETE FROM computers WHERE barcode = %s", (barcode,))
        elif action == 'restore':
            cur.execute("UPDATE computers SET status = 'תקין' WHERE barcode = %s", (barcode,))
            
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error in admin_approve_delete: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/manage-users')
@login_required
def manage_users():
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        flash("אין לך הרשאה לגשת לעמוד זה", "danger")
        return redirect(url_for('portal'))
        
    conn = get_db_connection()
    if not conn: return "DB connection failed", 500
    try:
        cur = get_safe_cursor(conn)
        # 1. Fetch users
        cur.execute("SELECT username, role, timestamp FROM users ORDER BY timestamp DESC NULLS LAST")
        users_raw = cur.fetchall()

        # קבע מי מחובר (פעיל ב-30 דקות האחרונות)
        now = datetime.now()
        users = []
        for u in users_raw:
            ts = u['timestamp'] if hasattr(u, '__getitem__') else u[2]
            is_online = False
            if ts:
                if hasattr(ts, 'replace'):
                    ts_naive = ts.replace(tzinfo=None) if hasattr(ts, 'tzinfo') and ts.tzinfo else ts
                    is_online = (now - ts_naive).total_seconds() < 1800  # 30 דקות
            users.append({'username': u['username'] if hasattr(u, '__getitem__') else u[0],
                          'role': u['role'] if hasattr(u, '__getitem__') else u[1],
                          'timestamp': ts,
                          'is_online': is_online})

        # 2. Fetch pending deletions
        cur.execute("SELECT barcode as computer_number, cage_number FROM computers WHERE status = 'ממתין למחיקה'")
        pending_raw = cur.fetchall()

        # Format pending for template
        pending = []
        for p in pending_raw:
            pending.append({
                'computer_number': p['computer_number'],
                'cage_number': p['cage_number'] or 'לא ידוע',
                'scanned_by': 'טכנאי'
            })

        cur.close()
        return render_template('manage_users.html', users=users, pending=pending)
    finally:
        release_db_connection(conn)

@app.route('/api/add_user', methods=['POST'])
@login_required
def api_add_user():
    # Only admin should be able to add users
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        return {"success": False, "error": "Unauthorized"}, 403
        
    data = request.json
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    role = data.get('role', 'technician')
    
    if not username or not password:
        return {"success": False, "error": "Missing username or password"}, 400
        
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        hashed_pass = generate_password_hash(password)
        cur.execute("INSERT INTO users (username, password, role) VALUES (%s, %s, %s)", (username, hashed_pass, role))
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error adding user: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/update_user', methods=['POST'])
@login_required
def api_update_user():
    if session.get('role') != 'admin' and session.get('user') not in ('uri', 'admin_uri'):
        return {"success": False, "error": "Unauthorized"}, 403
    data = request.json or {}
    orig_username = data.get('orig_username', '').strip()
    new_username  = data.get('new_username', '').strip()
    new_password  = (data.get('password') or '').strip()
    new_role      = data.get('role', '').strip()
    allowed_roles = ['technician', 'scanner', 'manager', 'logistics', 'admin']
    if not orig_username or not new_username:
        return {"success": False, "error": "שם משתמש לא תקין"}, 400
    if new_role and new_role not in allowed_roles:
        return {"success": False, "error": "הרשאה לא תקינה"}, 400
    if orig_username == 'admin_uri':
        return {"success": False, "error": "לא ניתן לערוך admin_uri מכאן"}, 400
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        if new_password:
            hashed = generate_password_hash(new_password)
            cur.execute("UPDATE users SET username=%s, password=%s, role=%s WHERE username=%s",
                        (new_username, hashed, new_role or 'technician', orig_username))
        else:
            cur.execute("UPDATE users SET username=%s, role=%s WHERE username=%s",
                        (new_username, new_role or 'technician', orig_username))
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        conn.rollback()
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/update_user_password', methods=['POST'])
@login_required
def api_update_user_password():
    if session.get('role') != 'admin' and session.get('user') not in ('uri', 'admin_uri'):
        return {"success": False, "error": "Unauthorized"}, 403
    data = request.json or {}
    username = data.get('username', '').strip()
    new_password = data.get('password', '').strip()
    if not username or len(new_password) < 4:
        return {"success": False, "error": "נתונים לא תקינים"}, 400
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        hashed = generate_password_hash(new_password)
        cur.execute("UPDATE users SET password = %s WHERE username = %s", (hashed, username))
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/update_user_role', methods=['POST'])
@login_required
def api_update_user_role():
    if session.get('role') != 'admin' and session.get('user') not in ('uri', 'admin_uri'):
        return {"success": False, "error": "Unauthorized"}, 403
    data = request.json or {}
    username = data.get('username', '').strip()
    new_role = data.get('role', '').strip()
    allowed_roles = ['technician', 'scanner', 'manager', 'logistics', 'admin']
    if not username or new_role not in allowed_roles:
        return {"success": False, "error": "נתונים לא תקינים"}, 400
    if username == 'admin_uri':
        return {"success": False, "error": "לא ניתן לשנות את הרשאות admin_uri"}, 400
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("UPDATE users SET role = %s WHERE username = %s", (new_role, username))
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/reset_user_password/<username>', methods=['POST'])
@login_required
def api_reset_user_password(username):
    """Reset a user's password to the default: username* (e.g. daniel*)"""
    if session.get('role') != 'admin' and session.get('user') not in ('uri', 'admin_uri'):
        return {"success": False, "error": "Unauthorized"}, 403
    if username == 'admin_uri':
        return {"success": False, "error": "לא ניתן לאפס את admin_uri"}, 400
    default_password = username + '*'
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        hashed = generate_password_hash(default_password)
        cur.execute("UPDATE users SET password = %s WHERE username = %s", (hashed, username))
        conn.commit()
        cur.close()
        return {"success": True, "message": f"סיסמא אופסה ל: {default_password}"}
    except Exception as e:
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/api/delete_user/<username>', methods=['DELETE'])
@login_required
def api_delete_user(username):
    # Only admin should be able to delete users
    if session.get('role') != 'admin' and session.get('user') != 'admin_uri':
        return {"success": False, "error": "Unauthorized"}, 403
        
    if username == 'admin_uri':
        return {"success": False, "error": "Cannot delete super-admin"}, 400
        
    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("DELETE FROM users WHERE username = %s", (username,))
        conn.commit()
        cur.close()
        return {"success": True}
    except Exception as e:
        print(f"Error deleting user: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/cages')
@login_required
def cages_page():
    conn = get_db_connection()
    if not conn: return "DB Error", 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("""
            SELECT
                c.cage_id, c.name, c.location, c.notes,
                COUNT(comp.id) AS computer_count,
                SUM(CASE WHEN comp.status = 'תקין' THEN 1 ELSE 0 END) AS ok_count,
                SUM(CASE WHEN comp.status = 'תקול' THEN 1 ELSE 0 END) AS broken_count,
                SUM(CASE WHEN comp.status NOT IN ('תקין','תקול') AND comp.status IS NOT NULL THEN 1 ELSE 0 END) AS other_count
            FROM cages c
            LEFT JOIN computers comp ON comp.cage_number = c.cage_id OR comp.cage_name = c.cage_id
            GROUP BY c.id, c.cage_id, c.name, c.location, c.notes
            ORDER BY computer_count DESC
        """)
        cages = cur.fetchall()
        cur.close()
        return render_template('cages.html', cages=cages)
    except Exception as e:
        return f"<h1>Error: {e}</h1>", 500
    finally:
        release_db_connection(conn)

@app.route('/api/cage/save', methods=['POST'])
@login_required
def api_save_cage():
    data = request.json
    cage_id = data.get('cage_id', '').strip()
    existing_id = data.get('existing_id', '').strip()
    if not cage_id: return {'success': False, 'error': 'cage_id is required'}, 400
    conn = get_db_connection()
    if not conn: return {'success': False, 'error': 'DB Error'}, 500
    try:
        cur = get_safe_cursor(conn)
        if existing_id:
            cur.execute("UPDATE cages SET name=%s, location=%s, notes=%s, updated_at=NOW() WHERE cage_id=%s", 
                        (data.get('name',''), data.get('location',''), data.get('notes',''), existing_id))
        else:
            cur.execute("INSERT INTO cages (cage_id, name, location, notes) VALUES (%s, %s, %s, %s) ON CONFLICT (cage_id) DO UPDATE SET name=EXCLUDED.name, location=EXCLUDED.location, notes=EXCLUDED.notes, updated_at=NOW()", 
                        (cage_id, data.get('name',''), data.get('location',''), data.get('notes','')))
        conn.commit()
        cur.close()
        return {'success': True}
    finally: release_db_connection(conn)

@app.route('/api/pack-cage-photo', methods=['POST'])
@login_required
def api_pack_cage_photo():
    data = request.json
    cage_id = data.get('cage_id', '').strip()
    image_b64 = data.get('image', '')
    
    if not cage_id or not image_b64:
        return {"success": False, "error": "Missing cage_id or image"}, 400
        
    # Extract base64 part
    if ',' in image_b64:
        image_b64 = image_b64.split(',', 1)[1]
        
    try:
        image_data = base64.b64decode(image_b64)
    except Exception as e:
        return {"success": False, "error": "Invalid image data"}, 400

    conn = get_db_connection()
    if not conn: return {"success": False, "error": "DB connection failed"}, 500
    cur = None
    try:
        # Call Gemini Vision
        prompt = "Look at the handwritten numbers written in white on the edges of the laptops in the cage. Extract all of them. Return ONLY a JSON array of strings (e.g. [\"1064\", \"366\", \"1480\"]). Do not add any markdown, comments, or other text."

        if not genai_client:
            return {"success": False, "error": "מפתח GOOGLE_API_KEY חסר — תכונת זיהוי תמונה לא זמינה"}, 503

        if getattr(genai_client, '_legacy_mode', False):
            # Old google-generativeai SDK
            import PIL.Image
            img_pil = PIL.Image.open(io.BytesIO(image_data))
            response = genai_client.generate_content([prompt, img_pil])
        else:
            # New google-genai SDK
            response = genai_client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[
                    genai.types.Part.from_bytes(data=image_data, mime_type="image/jpeg"),
                    prompt
                ]
            )
        
        try:
            # Clean response text just in case Gemini adds markdown
            text = response.text.strip()
            if text.startswith('```json'): text = text[7:]
            if text.startswith('```'): text = text[3:]
            if text.endswith('```'): text = text[:-3]
            text = text.strip()
            extracted_numbers = json.loads(text)
            if not isinstance(extracted_numbers, list):
                extracted_numbers = []
        except Exception as e:
            print(f"Gemini parse error: {e}. Raw response: {response.text}")
            return {"success": False, "error": "Could not parse AI response as JSON list."}, 500
            
        cur = get_safe_cursor(conn)
        
        results = {
            "success_count": 0,
            "transferred_count": 0,
            "new_count": 0,
            "details": []
        }
        
        technician = session.get('username', 'unknown')
        
        # Process each extracted number
        for barcode in extracted_numbers:
            barcode = str(barcode).strip()
            barcode = re.sub(r'^0+(?=\d)', '', barcode)
            if not barcode: continue
            
            cur.execute("SELECT * FROM computers WHERE barcode = %s ORDER BY id DESC LIMIT 1", (barcode,))
            computer = cur.fetchone()
            
            if computer:
                old_val = dict(computer)
                prev_cage = old_val.get('cage_number', '')
                
                # Check if it's already in this cage
                is_transfer = (prev_cage and prev_cage != cage_id)
                
                # Update
                cur.execute("""
                    UPDATE computers 
                    SET cage_number = %s, cage_name = %s, scan_time = NOW(), last_technician = %s
                    WHERE id = %s
                """, (cage_id, cage_id, technician, old_val['id']))
                
                # Log history
                cur.execute("""
                    INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                    VALUES (%s, %s, 'Photo Pack', %s, %s)
                """, (
                    old_val['id'],
                    technician,
                    json.dumps(old_val, default=str),
                    json.dumps({"cage_number": cage_id, "cage_name": cage_id}, default=str)
                ))
                
                if is_transfer:
                    results["transferred_count"] += 1
                    results["details"].append({"barcode": barcode, "status": "transferred", "prev_cage": prev_cage})
                else:
                    results["success_count"] += 1
                    results["details"].append({"barcode": barcode, "status": "updated"})
                    
            else:
                # New computer
                cur.execute("""
                    INSERT INTO computers (barcode, cage_number, cage_name, status, scan_time, last_technician)
                    VALUES (%s, %s, %s, 'תקין', NOW(), %s) RETURNING id
                """, (barcode, cage_id, cage_id, technician))
                new_id = cur.fetchone()['id']
                
                cur.execute("""
                    INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                    VALUES (%s, %s, 'Photo Pack (New)', %s, %s)
                """, (
                    new_id,
                    technician,
                    None,
                    json.dumps({"cage_number": cage_id, "cage_name": cage_id, "status": "תקין"}, default=str)
                ))
                
                results["new_count"] += 1
                results["details"].append({"barcode": barcode, "status": "new"})
                
        conn.commit()
        cur.close()
        
        # Trigger async sync
        trigger_debounced_sync()
        
        results["total_extracted"] = len(extracted_numbers)
        results["success"] = True
        return results

    except Exception as e:
        if conn:
            conn.rollback()
        print(f"Error in api_pack_cage_photo: {e}")
        traceback.print_exc()
        return {"success": False, "error": str(e)}, 500
    finally:
        if cur:
            try:
                cur.close()
            except Exception:
                pass
        release_db_connection(conn)

@app.route('/scan-dashboard')
@login_required
def scan_dashboard():
    return render_template('scan_dashboard.html')

@app.route('/api/scan-stats')
@login_required
def api_scan_stats():
    conn = get_db_connection()
    if not conn: return {'error': 'DB Error'}, 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT COUNT(*) as cnt FROM inventory_history WHERE timestamp::date = CURRENT_DATE AND change_type IN ('Fast Scan', 'Update via Scan')")
        today_total = cur.fetchone()['cnt']
        cur.execute("SELECT COUNT(*) as cnt FROM computers")
        total_computers = cur.fetchone()['cnt']
        cur.execute("SELECT COUNT(*) as cnt FROM computers WHERE status = 'תקול'")
        broken = cur.fetchone()['cnt']
        cur.execute("SELECT technician, COUNT(*) as count, MAX(timestamp) as last_scan FROM inventory_history WHERE timestamp::date = CURRENT_DATE AND change_type IN ('Fast Scan', 'Update via Scan') GROUP BY technician ORDER BY count DESC")
        workers = [dict(r) for r in cur.fetchall()]
        cur.close()
        return {'today_total': today_total, 'total_computers': total_computers, 'broken': broken, 'workers': workers}
    finally: release_db_connection(conn)

# ── API: ייבוא קובץ Excel תקולים ────────────────────────────────────────────────
@app.route('/api/upload-faulty-excel', methods=['POST'])
@login_required
def api_upload_faulty_excel():
    """
    מקבל קובץ Excel של מחשבים תקולים ומעדכן את ה-DB:
    עמודות: סוג מחשב | מספר מחשב | תקלה | בטיפול | במעבדה | תוקן
    """
    if 'file' not in request.files:
        return {"success": False, "error": "לא נשלח קובץ"}, 400
    file = request.files['file']
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        return {"success": False, "error": "יש להעלות קובץ Excel (.xlsx / .xls)"}, 400

    try:
        import io
        wb = openpyxl.load_workbook(io.BytesIO(file.read()))
        ws = wb.active
    except Exception as e:
        return {"success": False, "error": f"שגיאה בפתיחת הקובץ: {e}"}, 400

    conn = get_db_connection()
    if not conn:
        return {"success": False, "error": "DB connection failed"}, 500

    updated = 0
    not_found = 0
    errors = []

    try:
        cur = get_safe_cursor(conn)
        # דלג על שורת כותרת
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[1] is None:
                continue
            try:
                barcode = str(int(row[1])).strip()
            except (ValueError, TypeError):
                continue

            fault_desc  = str(row[2] or '').strip()   # תקלה
            in_repair   = bool(row[3])                  # בטיפול אצלינו
            in_lab      = bool(row[4])                  # במעבדה
            fixed       = bool(row[5])                  # תוקן

            # קבע סטטוס ומיקום
            if fixed:
                new_status   = 'תקין'
                new_location = ''
            elif in_lab:
                new_status   = 'בתיקון'
                new_location = 'במעבדה'
            elif in_repair:
                new_status   = 'בתיקון'
                new_location = 'בטיפול'
            else:
                new_status   = 'תקול'
                new_location = ''

            # בדוק קיום מחשב
            cur.execute("SELECT id, status, location, notes FROM computers WHERE barcode = %s", (barcode,))
            computer = cur.fetchone()
            if not computer:
                not_found += 1
                continue

            cid = computer['id']
            # עדכן
            cur.execute("""
                UPDATE computers
                SET status = %s, location = %s, notes = %s, last_technician = %s
                WHERE id = %s
            """, (new_status, new_location, fault_desc, session.get('username'), cid))

            # רשום היסטוריה
            old_val = json.dumps({'status': computer['status'], 'location': computer['location'], 'notes': computer['notes']}, ensure_ascii=False)
            new_val = json.dumps({'status': new_status, 'location': new_location, 'notes': fault_desc}, ensure_ascii=False)
            cur.execute("""
                INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                VALUES (%s, %s, 'Excel Import', %s, %s)
            """, (cid, session.get('username'), old_val, new_val))

            updated += 1

        conn.commit()
        cur.close()
        trigger_debounced_sync()
        return {
            "success": True,
            "updated": updated,
            "not_found": not_found,
            "message": f"עודכנו {updated} מחשבים | לא נמצאו: {not_found}"
        }
    except Exception as e:
        print(f"Error in upload_faulty_excel: {e}")
        return {"success": False, "error": str(e)}, 500
    finally:
        release_db_connection(conn)

@app.route('/export/computers')
@login_required
def export_computers():
    conn = get_db_connection()
    if not conn: return 'DB Error', 500
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT barcode, cage_number, cage_name, location, status, case_number, exam_appeal, notes, scan_time FROM computers ORDER BY cage_number, barcode")
        rows = cur.fetchall()
        cur.close()
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'מחשבים'
        ws.sheet_view.rightToLeft = True
        headers = ['ברקוד', 'כלוב', 'שם כלוב', 'מיקום', 'סטטוס', 'מספר תיק', 'מבחן/ערעור', 'הערות', 'נסרק לאחרונה']
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
        for row_num, row in enumerate(rows, 2):
            values = [row.get('barcode',''), row.get('cage_number',''), row.get('cage_name',''), row.get('location',''), row.get('status',''), row.get('case_number',''), row.get('exam_appeal',''), row.get('notes',''), str(row.get('scan_time',''))[0:16] if row.get('scan_time') else '']
            for col_num, val in enumerate(values, 1): ws.cell(row=row_num, column=col_num, value=val)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True, download_name=f"inventory_{datetime.now().strftime('%Y%m%d')}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    finally: release_db_connection(conn)

# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# מערכת נוכחות נבחנים
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

@app.route('/exam-attendance')
@login_required
def exam_attendance():
    """
    דשבורד נוכחות נבחנים.
    אין כאן שאילתת נתוני נבחנים — Google Drive הוא מקור האמת היחיד
    לשמות/ת.ז/סיסמאות, ולא נשמר מהם דבר במסד. 'projects' מחזיקה רק
    מטא-דאטה (שם מבחן + קישור לגיליון), בלי נתון אישי כלשהו.
    """
    exams = []
    conn = get_db_connection()
    if conn:
        try:
            cur = get_safe_cursor(conn)
            cur.execute("SELECT name, drive_url FROM projects WHERE sheets_id != '' ORDER BY name")
            exams = [dict(r) for r in cur.fetchall()]
            cur.close()
        except Exception as e:
            print(f"[DASHBOARD] projects lookup failed: {e}", flush=True)
        finally:
            release_db_connection(conn)
    return render_template('exam_attendance.html', exams=exams)

@app.route('/exam-attendance/add', methods=['GET', 'POST'])
@login_required
def exam_attendance_add():
    """הוספת נבחן ידנית — נכתב ישירות לגיליון ב-Drive, אין כאן מסד"""
    if request.method == 'POST':
        data = request.form
        from exam_naming import parse_exam_title
        from drive_manager import get_exam_sheet, merge_examinees_into_sheet

        raw_exam = data.get('exam_name', '').strip()
        name = data.get('name', '').strip()
        if not raw_exam or not name:
            flash("יש למלא שם מבחן ושם נבחן", "danger")
            return render_template('exam_attendance_add.html')

        record = {
            'full_name':   name,
            'id_number':   data.get('id_number', '').strip(),
            'username':    data.get('username', '').strip(),
            'password':    data.get('password', '').strip(),
            'hall':        data.get('location', '').strip(),
            'computer':    data.get('computer', '').strip(),
            'row':         data.get('row', '').strip(),
            'seat':        data.get('seat', '').strip(),
            'adaptations': data.get('notes', '').strip(),
        }
        try:
            info = parse_exam_title(raw_exam)
            target = get_exam_sheet(raw_exam, create=True)
            if not target:
                flash("לא הצלחתי ליצור/למצוא את גיליון המבחן ב-Drive", "danger")
                return render_template('exam_attendance_add.html')
            merge_examinees_into_sheet(target['worksheet'], [record], title_text=info['exam_name'])
            invalidate_examinee_cache(raw_exam)
            flash(f"נבחן {name} נוסף בהצלחה ל-Drive! ✅", "success")
            return redirect(url_for('exam_attendance'))
        except Exception as e:
            flash(f"שגיאה: {e}", "danger")
    return render_template('exam_attendance_add.html')

# ══ מטמון-תהליך לנתוני נבחנים (RAM בלבד — לא מסד, לא דיסק) ══════════
# Google Drive הוא מקור האמת היחיד לנתוני נבחנים (שם/ת.ז/סיסמה).
# המטמון הזה קיים רק כדי שסריקה שנייה של אותו מבחן לא תצטרך לקרוא
# מחדש את כל הגיליון מגוגל — הוא נמחק בכל הפעלה מחדש של השרת, ואינו
# נשמר בשום מקום קבוע.
examinee_cache = {}  # canonical_exam_name -> {id_number_or_name: record}


def load_examinee_cache(exam_name, force=False):
    """
    טוען את רשימת הנבחנים של מבחן מהגיליון שלו ב-Drive לזיכרון.
    מחזיר dict (ריק אם הגיליון לא נמצא), ולא זורק חריגה.
    """
    global examinee_cache
    from exam_naming import parse_exam_title
    canonical = parse_exam_title(exam_name)['exam_name'] if exam_name else ''
    if not canonical:
        return {}
    if not force and canonical in examinee_cache:
        return examinee_cache[canonical]

    from drive_manager import load_examinee_records
    try:
        target = resolve_exam_sheet(exam_name, create=False)
    except Exception as e:
        print(f"[CACHE] sheet lookup failed for '{canonical}': {e}", flush=True)
        target = None
    if not target:
        examinee_cache[canonical] = {}
        return {}
    try:
        records = load_examinee_records(target['worksheet'])
    except Exception as e:
        print(f"[CACHE] Failed to load examinees for '{canonical}': {e}", flush=True)
        records = {}
    examinee_cache[canonical] = records
    print(f"[CACHE] Loaded {len(records)} examinees for '{canonical}'", flush=True)
    return records


def invalidate_examinee_cache(exam_name):
    """קורא שוב מהגיליון בפעם הבאה — משתמשים בזה אחרי ייבוא/הוספה."""
    from exam_naming import parse_exam_title
    canonical = parse_exam_title(exam_name)['exam_name'] if exam_name else ''
    examinee_cache.pop(canonical, None)


def find_examinee(exam_name, id_number=None, full_name=None):
    """מחפש נבחן במטמון (טוען מהגיליון אם צריך). מחזיר dict או None."""
    records = load_examinee_cache(exam_name)
    if id_number and id_number in records:
        return records[id_number]
    if full_name and full_name in records:
        return records[full_name]
    return None


def mark_examinee_scanned(exam_name, id_number, full_name, computer, technician,
                          pc_status='', is_present=1):
    """
    מעדכן את המטמון בזיכרון בלבד (לא כותב לגיליון — זה תפקידו של
    write_examinee_scan/הקורא). משמש כדי שסריקות עוקבות של אותו מבחן
    יראו את הסטטוס העדכני בלי לקרוא שוב את הגיליון.
    מחזיר True אם הנבחן נמצא במטמון (כלומר רשום למבחן הזה).
    """
    records = load_examinee_cache(exam_name)
    key = id_number if id_number in records else (full_name if full_name in records else None)
    if key is None:
        return False
    rec = records[key]
    rec['is_present'] = str(is_present)
    rec['scan_time'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    rec['technician'] = technician
    rec['pc_status'] = pc_status
    if computer:
        rec['computer'] = computer
    return True


def resolve_exam_sheet(raw_exam_title, create=False):
    """
    מאתר את הגיליון של מבחן — מאותו מקור בדיוק שהייבוא השתמש בו.

    raw_exam_title: הכותרת הגולמית (parts[0] של ה-QR / שדה exam_name בטופס).
    מחזיר dict של drive_manager.get_exam_sheet, או None.

    סדר החיפוש:
      1. טבלת projects (מהיר — נשמר בייבוא)
      2. Drive: תיקיית המשרד ← גיליון המבחן
    """
    from exam_naming import parse_exam_title
    from drive_manager import get_exam_sheet

    info = parse_exam_title(raw_exam_title)
    exam_name = info['exam_name']

    # 1. מיפוי שמור — חוסך סריקת Drive
    sheet_id = None
    conn = get_db_connection()
    if conn:
        try:
            cur = get_safe_cursor(conn)
            cur.execute("SELECT sheets_id FROM projects WHERE name = %s", (exam_name,))
            row = cur.fetchone()
            if row and row['sheets_id']:
                sheet_id = row['sheets_id']
            cur.close()
        except Exception as e:
            print(f"[SCAN] project lookup failed: {e}", flush=True)
        finally:
            release_db_connection(conn)

    if sheet_id:
        try:
            from drive_manager import _get_clients
            gs, _ = _get_clients()
            sh = gs.open_by_key(sheet_id)
            result = dict(info)
            result.update({
                'sheet_id': sheet_id,
                'spreadsheet': sh,
                'worksheet': sh.sheet1,
                'url': f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
            })
            return result
        except Exception as e:
            print(f"[SCAN] stored sheet {sheet_id} unusable ({e}) — falling back to Drive", flush=True)

    # 2. חיפוש/יצירה ב-Drive. raise_errors=True: כשל אמיתי (לא "לא
    # נמצא") נזרק החוצה כדי שהקורא יציג את הסיבה האמיתית למשתמש,
    # ולא רק "לא הצלחתי לפתוח/ליצור גיליון" בלי שום הסבר
    return get_exam_sheet(raw_exam_title, create=create, raise_errors=True)


def convert_docx_to_pdf(docx_bytes, timeout=60):
    """
    ממיר bytes של docx ל-PDF דרך LibreOffice headless.
    מחזיר bytes של ה-PDF, או None אם soffice לא מותקן/נכשל — כדי שהקורא
    ייפול חזרה ל-Word במקום להחזיר שגיאה למשתמש.

    שים לב: זה דורש LibreOffice מותקן בשרת. סביבת Render הרגילה (runtime:
    python, ללא Dockerfile) בדרך כלל לא כוללת אותו — הפונקציה נכשלת בשקט
    במקרה הזה, וזה בכוונה.
    """
    import shutil
    import subprocess
    import tempfile

    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        print("[PDF] soffice not found — PDF conversion unavailable on this server", flush=True)
        return None

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            src_path = os.path.join(tmpdir, 'input.docx')
            with open(src_path, 'wb') as f:
                f.write(docx_bytes)

            result = subprocess.run(
                [soffice, '--headless', '--norestore', '--convert-to', 'pdf',
                 '--outdir', tmpdir, src_path],
                capture_output=True, timeout=timeout
            )
            pdf_path = os.path.join(tmpdir, 'input.pdf')
            if result.returncode != 0 or not os.path.exists(pdf_path):
                print(f"[PDF] soffice conversion failed (rc={result.returncode}): "
                      f"{result.stderr.decode(errors='replace')[:500]}", flush=True)
                return None
            with open(pdf_path, 'rb') as f:
                return f.read()
    except Exception as e:
        print(f"[PDF] conversion error: {e}", flush=True)
        return None


def save_project_mapping(exam_name, sheet_id, sheet_url, office=''):
    """שומר/מעדכן את מיפוי המבחן → גיליון בטבלת projects (משמש את הסורק)."""
    conn = get_db_connection()
    if not conn:
        return
    try:
        cur = get_safe_cursor(conn)
        cur.execute("SELECT id FROM projects WHERE name = %s", (exam_name,))
        existing = cur.fetchone()
        keywords = ' '.join(filter(None, [exam_name, office]))
        if existing:
            cur.execute("UPDATE projects SET sheets_id=%s, drive_url=%s, keywords=%s WHERE name=%s",
                        (sheet_id, sheet_url, keywords, exam_name))
        else:
            cur.execute("INSERT INTO projects (name, keywords, sheets_id, drive_url) VALUES (%s,%s,%s,%s)",
                        (exam_name, keywords, sheet_id, sheet_url))
        conn.commit()
        cur.close()
    except Exception as e:
        conn.rollback()
        print(f"[IMPORT] Failed to save project mapping: {e}", flush=True)
    finally:
        release_db_connection(conn)


@app.route('/exam-attendance/import', methods=['POST'])
@login_required
def exam_attendance_import():
    """ייבוא נבחנים מ-Excel"""
    file = request.files.get('excel_file')
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        flash("יש להעלות קובץ Excel (.xlsx / .xls)", "danger")
        return redirect(url_for('exam_attendance'))

    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        
        # קריאת כותרת הבחינה משורה 1
        exam_title_row1 = ''
        for cell in ws[1]:
            if cell.value and str(cell.value).strip():
                exam_title_row1 = str(cell.value).strip()
                break

        headers = []
        header_row_idx = 1
        for row_idx, r in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
            row_strs = [str(c).strip() if c else '' for c in r]
            if any('שם' in s or 'תעודת' in s or 'ת.ז' in s or 'id' in s.lower() or 'פרטי' in s for s in row_strs):
                headers = row_strs
                header_row_idx = row_idx
                break
        if not headers:
            headers = [str(cell.value).strip() if cell.value else '' for cell in ws[1]]

        # מיפוי עמודות גמיש â€” כולל שם פרטי + שם משפחה
        col_map = {}
        for i, h in enumerate(headers):
            h_lower = h.lower()
            if 'פרטי' in h:                                          col_map['first_name'] = i
            elif 'משפחה' in h or 'family' in h_lower:               col_map['last_name']  = i
            elif ('שם' in h or 'name' in h_lower) and 'משתמש' not in h: col_map['name'] = i
            elif 'זהות' in h or 'ת.ז' in h or 'id' in h_lower:     col_map['id_number']  = i
            elif 'משתמש' in h or 'user' in h_lower or 'קוד' in h:  col_map['username']   = i
            elif 'סיסמ' in h or 'pass' in h_lower:                  col_map['password']   = i
            elif 'טור' in h or 'עמודה' in h:                        col_map['row_number'] = i
            elif 'כסא' in h or 'כיסא' in h or 'seat' in h_lower or 'מושב' in h:   col_map['seat_number']= i
            elif 'מיקום' in h or 'כיתה' in h or 'location' in h_lower or 'אולם' in h: col_map['hall']= i
            elif 'בחינה' in h or 'גרסה' in h or 'exam' in h_lower:  col_map['exam_name']  = i
            elif 'מחשב' in h or 'computer' in h_lower:              col_map['computer']   = i
            elif 'התאמות' in h or 'notes' in h_lower:               col_map['adaptations'] = i

        def safe_val(val):
            """המרת ערך תא â€” מטפל במספרים גדולים/נוטציה מדעית"""
            if val is None: return ''
            if isinstance(val, float):
                if val == int(val): return str(int(val))
                return str(val)
            return str(val).strip()

        # איסוף כל השורות
        all_rows = []
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if not any(row): continue
            def get_val(key, r=row):
                idx = col_map.get(key)
                if idx is None: return ''
                return safe_val(r[idx])

            # שם מלא: אם יש שם פרטי + משפחה â€” מאחד
            first = get_val('first_name')
            last  = get_val('last_name')
            if first or last:
                name = (first + ' ' + last).strip()
            else:
                name = get_val('name')
            if not name: continue

            all_rows.append({
                'full_name':   name,
                'id_number':   get_val('id_number'),
                'username':    get_val('username'),
                'password':    get_val('password'),
                'hall':        get_val('hall'),
                'version':     get_val('exam_name'),
                'computer':    get_val('computer'),
                'adaptations': get_val('adaptations'),
                'row':         get_val('row_number'),
                'seat':        get_val('seat_number'),
            })

        # מיון לפי טור â†’ כסא
        def sort_key(r):
            try:    row_n  = int(float(r['row']))  if r['row']  else 9999
            except: row_n  = 9999
            try:    seat_n = int(float(r['seat'])) if r['seat'] else 9999
            except: seat_n = 9999
            return (row_n, seat_n)
        all_rows.sort(key=sort_key)

        if not all_rows:
            flash("לא נמצאו שורות נבחנים בקובץ", "warning")
            return redirect(url_for('exam_attendance'))

        # ── שם המשרד/מבחן/תאריך — מכותרת שורה 1, ובהיעדרה משם הקובץ ──
        # כשהכותרות עצמן בשורה 1 אין שורת כותרת ממוזגת, ו-exam_title_row1
        # מחזיק את שם העמודה הראשונה ("שם פרטי") — לא שם מבחן.
        if header_row_idx <= 1:
            exam_title_row1 = ''
        from exam_naming import parse_exam_title
        info = parse_exam_title(exam_title_row1, file.filename)
        exam_name = info['exam_name']

        # גרסה/אולם חסרים בשורה? משלימים משם המבחן
        for rec in all_rows:
            if not rec.get('version'):
                rec['version'] = exam_name
            if info['hall'] and not rec.get('hall'):
                rec['hall'] = info['hall']

        print(f"[IMPORT] office='{info['office']}' sheet='{info['sheet_name']}' "
              f"exam='{exam_name}' hall='{info['hall']}' rows={len(all_rows)}", flush=True)

        # ── Google Drive: תיקיית משרד ← גיליון המבחן (מקור האמת היחיד) ──
        from drive_manager import get_exam_sheet, merge_examinees_into_sheet

        # מעבירים כותרת ושם קובץ בנפרד — כך סיומת .xlsx נחתכת ושם הגיליון
        # יוצא זהה ל-exam_name הקנוני שהסריקה תחפש
        target = get_exam_sheet(exam_title_row1, filename=file.filename)
        if not target:
            flash("❌ יצירת/איתור הגיליון ב-Drive נכשלה — הייבוא לא הושלם", "danger")
            return redirect(url_for('exam_attendance'))

        # מיזוג — לא מחיקה. מבחן מחולק לאולמות מיובא בכמה קבצים לאותו גיליון.
        added, updated = merge_examinees_into_sheet(target['worksheet'], all_rows,
                                                     title_text=exam_name)
        sheet_url = target['url']
        save_project_mapping(exam_name, target['sheet_id'], sheet_url, info['office'])
        invalidate_examinee_cache(exam_name)  # הסריקה תקרא מחדש מהגיליון המעודכן

        summary = f"✅ {added} נבחנים חדשים"
        if updated:
            summary += f", {updated} עודכנו"
        summary += f" — {info['office']} ← {info['sheet_name']}"
        if info['hall']:
            summary += f" ({info['hall']})"
        flash(summary, "success")
        flash(f"🔗 <a href='{sheet_url}' target='_blank'>פתח את הגיליון ב-Google Drive</a>", "info")

    except Exception as e:
        import traceback; traceback.print_exc()
        flash(f"שגיאה בייבוא: {e}", "danger")
    return redirect(url_for('exam_attendance'))

@app.route('/api/generate-word-docs', methods=['POST'])
@login_required
def generate_word_docs():
    """מחולל דפי נבחנים מקובץ אקסל ותבנית וורד"""
    excel_file = request.files.get('excel_file')
    word_template = request.files.get('word_template')
    
    if not excel_file:
        flash("יש להעלות קובץ Excel", "danger")
        return redirect(url_for('exam_attendance'))
        
    try:
        # Load Excel data
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        
        # קריאת שם הבחינה משורה 1 (כותרת הדף) אם קיים
        exam_title_from_header = ''
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        for cell in first_row:
            if cell and str(cell).strip():
                exam_title_from_header = str(cell).strip()
                break

        headers = []
        header_row_idx = 1
        # זיהוי שורת כותרות: צריך לפחות 2 עמודות מוכרות (מונע שגיאות כמו "חשמלאים"âŠƒ"שם")
        HEADER_KEYS = ['שם', 'ת.ז', 'תעודת', 'סיסמ', 'משתמש', 'טור', 'כסא', 'מחשב', 'התאמות', 'פרטי', 'משפחה', 'id', 'pass', 'user', 'seat', 'name']
        for row_idx, r in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
            row_strs = [str(c).strip() if c else '' for c in r]
            matches = sum(1 for s in row_strs if any(k in s.lower() for k in HEADER_KEYS))
            if matches >= 2:
                headers = row_strs
                header_row_idx = row_idx
                break
        if not headers:
            headers = [str(cell.value).strip() if cell.value else '' for cell in ws[1]]

        def safe_val(val):
            """המרת ערך תא â€” מטפל במספרים גדולים/נוטציה מדעית"""
            if val is None: return ''
            if isinstance(val, float):
                if val == int(val): return str(int(val))
                return str(val)
            return str(val).strip()

        col_map = {}
        for i, h in enumerate(headers):
            h_lower = h.lower()
            if 'פרטי' in h:                                              col_map['first_name'] = i
            elif 'משפחה' in h or 'family' in h_lower:                   col_map['last_name']  = i
            elif ('שם' in h or 'name' in h_lower) and 'משתמש' not in h: col_map['name']       = i
            elif 'זהות' in h or 'ת.ז' in h or 'id' in h_lower:         col_map['id_number']  = i
            elif 'משתמש' in h or 'user' in h_lower or 'קוד' in h:      col_map['username']   = i
            elif 'סיסמ' in h or 'pass' in h_lower:                      col_map['password']   = i
            elif 'מיקום' in h or 'כיתה' in h or 'location' in h_lower or 'אולם' in h: col_map['location']   = i
            elif 'בחינה' in h or 'exam' in h_lower:                     col_map['exam_name']  = i
            elif 'מחשב' in h or 'computer' in h_lower:                  col_map['computer']   = i
            elif 'התאמות' in h or 'notes' in h_lower:                   col_map['notes']      = i
            elif 'טור' in h or 'עמודה' in h:                            col_map['row']        = i
            elif 'כסא' in h or 'כיסא' in h or 'seat' in h_lower or 'מושב' in h:       col_map['seat']       = i
            elif 'תקין' in h or 'valid' in h_lower:                      col_map['is_valid']   = i
            elif 'נוכחות' in h or 'attendance' in h_lower:               col_map['attendance'] = i

        examinees = []
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if not any(row): continue
            def get_val(key, r=row):
                idx = col_map.get(key)
                if idx is None: return ''
                return safe_val(r[idx])

            # DEBUG â€” שומר מיפוי לקובץ (למניעת שגיאת charmap בטרמינל)
            if len(examinees) == 0:
                try:
                    with open('debug_colmap.txt', 'w', encoding='utf-8') as dbf:
                        dbf.write(f"col_map = {col_map}\n")
                        dbf.write(f"headers = {headers}\n")
                        dbf.write(f"first data row = {[str(v) for v in row]}\n")
                except Exception:
                    pass

            # שם מלא: שם פרטי + שם משפחה אם קיימים
            first = get_val('first_name')
            last  = get_val('last_name')
            if first or last:
                name = (first + ' ' + last).strip()
            else:
                name = get_val('name')
            if not name: continue
            
            examinees.append({
                'name': name,
                'id_number': get_val('id_number'),
                'username': get_val('username'),
                'password': get_val('password'),
                'location': get_val('location'),
                'row': get_val('row'),
                'seat': get_val('seat'),
                'exam_name': get_val('exam_name'),
                'computer': get_val('computer'),
                'notes': get_val('notes'),
                'is_valid': get_val('is_valid'),
                'attendance': get_val('attendance'),
            })
            
        if not examinees:
            flash("לא נמצאו נתונים בקובץ האקסל", "warning")
            return redirect(url_for('exam_attendance'))

        # ── שמירה אוטומטית ל-Drive — בלי זה מי שמודפס כאן לעולם לא
        # נמצא בסריקה, כי הסריקה מחפשת בגיליון, לא בקובץ ה-Word ──
        from exam_naming import parse_exam_title
        from drive_manager import get_exam_sheet, merge_examinees_into_sheet

        groups = {}  # שם מבחן קנוני -> רשימת רשומות (יש תמיכה בעמודת exam_name שונה לכל שורה)
        for e in examinees:
            raw_title = e['exam_name'] or exam_title_from_header or 'EXAMINEE'
            canonical = parse_exam_title(raw_title, excel_file.filename)['exam_name']
            e['_canonical_exam'] = canonical  # אותו שם בדיוק ייכנס ל-QR בהמשך
            groups.setdefault(canonical, []).append({
                'full_name':   e['name'],
                'id_number':   e['id_number'],
                'username':    e['username'],
                'password':    e['password'],
                'hall':        e['location'],
                'computer':    e['computer'],
                'row':         e['row'],
                'seat':        e['seat'],
                'adaptations': e['notes'],
            })

        drive_saved = 0
        for canonical, records in groups.items():
            try:
                target = get_exam_sheet(canonical, create=True)
                if target:
                    merge_examinees_into_sheet(target['worksheet'], records, title_text=canonical)
                    drive_saved += len(records)
            except Exception as e_drive:
                print(f"[WORD-GEN] Drive save failed for '{canonical}': {e_drive}", flush=True)

        if drive_saved:
            flash(f"📁 {drive_saved} נבחנים נשמרו אוטומטית גם ב-Google Drive — הסריקה תמצא אותם", "info")
        else:
            flash("⚠️ השמירה האוטומטית ל-Drive נכשלה — הדפים יופקו, אך הסריקה לא תמצא אותם עד לייבוא ידני", "warning")

        # Load template
        if word_template and word_template.filename.endswith('.docx'):
            word_bytes = word_template.read()
        else:
            template_path = os.path.join(os.path.dirname(__file__), 'default_template.docx')
            with open(template_path, 'rb') as f:
                word_bytes = f.read()
        
        all_docs = []  # רשימה של (שם_קובץ, bytes) לכל נבחן
        
        for i, e in enumerate(examinees):
            # Load template first so we can use it for InlineImage
            tpl = DocxTemplate(io.BytesIO(word_bytes))
            
            # Format: exam_name|id_number|name|username|password|row|seat
            # exam_name: from column or from Excel title row
            exam_title = e['exam_name'] if e['exam_name'] else (exam_title_from_header or 'EXAMINEE')
            # Clean trailing "- נוכחות" or "נוכחות"
            exam_title_clean = re.sub(r'[-â€“]\s*נוכחות\s*$', '', exam_title).strip()
            exam_title_clean = exam_title_clean.replace('|', ' ').strip()  # מונע שבירת פורמט ה-QR
            qr_data = f"{exam_title_clean}|{e['id_number']}|{e['name']}|{e['username']}|{e['password']}|{e['row']}|{e['seat']}"
            # QR מ-API מבוטל â€” משתמשים רק ב-QR מקומי (ראה Step 2 בהמשך)
            qr_inline = ""
            
            # Context
            context = {
                'full_name': e['name'],
                'id_number': e['id_number'],
                'username': e['username'],
                'password': e['password'],
                'location': e['location'],
                'row': e['row'],
                'seat': e['seat'],
                'exam_name': e['exam_name'],
                'computer': e['computer'],
                'notes': e['notes'],
                'qr_code': qr_inline
            }
            
            tpl.render(context)
            
            # Save rendered to memory
            rendered_buf = io.BytesIO()
            tpl.save(rendered_buf)
            rendered_buf.seek(0)
            
            doc = Document(rendered_buf)

            # ---- Step 0.5: Replace static old title and location to match the imported exam ----
            def replace_static_text(p, new_title, new_loc, row_val):
                txt = p.text
                title_keywords = ["משרד הבריאות", "מינהל האחיות", "מומחיות", "טיפול תומך", "רישוי חשמלאים", "חשמלאים", "כיתות חשמל", "כיתת חשמל", "חשמל"]
                loc_keywords = ["בניין REIT1", "פתח תקווה", "מליאה", "אפעל 6", "בניין", "קומה", "כיתה"]
                
                is_title = any(kw in txt for kw in title_keywords) and "טופס התחברות" not in txt
                is_loc = any(kw in txt for kw in loc_keywords) and not is_title and "טופס התחברות" not in txt
                
                if '[1]' in txt and row_val:
                    if p.runs:
                        for r_item in p.runs:
                            if '[1]' in r_item.text:
                                r_item.text = r_item.text.replace('[1]', row_val)
                    else:
                        p.text = p.text.replace('[1]', row_val)

                if is_title:
                    if p.runs:
                        p.runs[0].text = new_title
                        for r_item in p.runs[1:]:
                            r_item.text = ""
                    else:
                        p.text = new_title
                elif is_loc and new_loc:
                    if p.runs:
                        p.runs[0].text = new_loc
                        for r_item in p.runs[1:]:
                            r_item.text = ""
                    else:
                        p.text = new_loc

            # Process sections (headers and footers)
            for section in doc.sections:
                # 1. Process headers with both heuristic and keywords
                for hf_name in ['header', 'first_page_header', 'even_page_header']:
                    hf = getattr(section, hf_name, None)
                    if hf:
                        # Heuristic index-based replacement for direct paragraphs in header
                        non_empty_ps = [p for p in hf.paragraphs if p.text.strip()]
                        title_p = None
                        loc_p = None
                        for p in non_empty_ps:
                            if "טופס התחברות" not in p.text:
                                if not title_p:
                                    title_p = p
                                elif not loc_p:
                                    loc_p = p
                                    break
                        if title_p:
                            if title_p.runs:
                                title_p.runs[0].text = exam_title_clean
                                for r in title_p.runs[1:]:
                                    r.text = ""
                            else:
                                title_p.text = exam_title_clean
                        if loc_p and e['location']:
                            if loc_p.runs:
                                loc_p.runs[0].text = e['location']
                                for r in loc_p.runs[1:]:
                                    r.text = ""
                            else:
                                loc_p.text = e['location']
                        
                        # Also apply keyword replacement for paragraphs and tables inside header
                        for p in hf.paragraphs:
                            replace_static_text(p, exam_title_clean, e['location'], str(e.get('row', '')))
                        for table in hf.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for p in cell.paragraphs:
                                        replace_static_text(p, exam_title_clean, e['location'], str(e.get('row', '')))

                # 2. Process footers with keywords replacement only (no heuristic to prevent corruption)
                for hf_name in ['footer', 'first_page_footer', 'even_page_footer']:
                    hf = getattr(section, hf_name, None)
                    if hf:
                        for p in hf.paragraphs:
                            replace_static_text(p, exam_title_clean, e['location'], str(e.get('row', '')))
                        for table in hf.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for p in cell.paragraphs:
                                        replace_static_text(p, exam_title_clean, e['location'], str(e.get('row', '')))

            # Process main body paragraphs and tables
            for p in doc.paragraphs:
                replace_static_text(p, exam_title_clean, e['location'], str(e.get('row', '')))
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_static_text(p, exam_title_clean, e['location'], str(e.get('row', '')))
            
            # Generate QR locally
            qr_gen = qrcode.QRCode(version=1, box_size=10, border=1)
            qr_gen.add_data(qr_data)
            qr_gen.make(fit=True)
            qr_img = qr_gen.make_image(fill_color="black", back_color="white")
            qr_buf_local = io.BytesIO()
            # PyPNGImage.save() doesn't accept format= kwarg (it's always PNG)
            # For Pillow-backed images it does, so handle both cases
            try:
                qr_img.save(qr_buf_local, **{'format': 'PNG'})  # type: ignore
            except TypeError:
                qr_img.save(qr_buf_local)
            qr_buf_local.seek(0)

            from docx.shared import Inches, Pt, Cm
            from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement

            # ---- Step 1: Fill data tables by label matching ----
            label_to_value = {
                'שם נבחן': e['name'],
                'תעודת זהות': e['id_number'],
                'קוד משתמש': e['username'],
                'סיסמה': e['password'],
                'התאמות': e['notes'],
                'אולם': e['location'],
                'כיתה': e['location'],
                'אולם/כיתה': e['location'],
                'מחשב': e['computer'],
                'מספר מחשב': e['computer'],
                'טור': e['row'],
                'כסא': e['seat'],
                'מושב': e['seat'],
                'תקין': e['is_valid'],
                'נוכחות': e['attendance'],
            }
            for t_idx, t in enumerate(doc.tables):
                # דלג על טבלת "למילוי על ידי הנבחן/ת" â€” נשארת ריקה לכתב יד
                table_text = ' '.join(cell.text for row in t.rows for cell in row.cells)
                if 'חתימה' in table_text or 'למילוי' in table_text:
                    continue
                for row_idx, r in enumerate(t.rows):
                    if len(r.cells) >= 2:
                        label_right = r.cells[1].text.strip()
                        label_left  = r.cells[0].text.strip()
                        for key, val_text in label_to_value.items():
                            val_str = str(val_text) if val_text is not None else ''
                            if key in label_right:
                                para = r.cells[0].paragraphs[0]
                                orig_run = para.runs[0] if para.runs else None
                                para.clear()
                                new_run = para.add_run(val_str)
                                if orig_run:
                                    new_run.font.name = orig_run.font.name
                                    new_run.font.size = orig_run.font.size
                                    new_run.font.bold = orig_run.font.bold
                                break
                            elif key in label_left:
                                para = r.cells[1].paragraphs[0]
                                orig_run = para.runs[0] if para.runs else None
                                para.clear()
                                new_run = para.add_run(val_str)
                                if orig_run:
                                    new_run.font.name = orig_run.font.name
                                    new_run.font.size = orig_run.font.size
                                    new_run.font.bold = orig_run.font.bold
                                break

            # ---- Step 2: Find seat paragraph in BODY and add QR ----
            seat_display = e.get('seat', '') or str(i + 1)  # fallback: מספר סידורי
            target_p = None
            for p in doc.paragraphs:
                txt = p.text.strip()
                if txt == seat_display or txt == '1':   # '1' הוא ה-placeholder בתבנית
                    target_p = p
                    break

            if target_p:
                # Force LTR on this specific paragraph (so QR is left, number is right)
                pPr = target_p._p.get_or_add_pPr()
                for bidi_el in pPr.findall(qn('w:bidi')):
                    pPr.remove(bidi_el)
                bidi = OxmlElement('w:bidi')
                bidi.set(qn('w:val'), '0')
                pPr.append(bidi)

                target_p.text = ''
                target_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                # 1. Add QR Code first (will be on the LEFT in LTR)
                qr_buf_local.seek(0)
                run_qr = target_p.add_run()
                run_qr.add_picture(qr_buf_local, width=Inches(1.2))

                # 2. Add a RIGHT tab stop to push the number to the far right
                tab_stops = target_p.paragraph_format.tab_stops
                tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
                target_p.add_run('\t')

                # 3. Add Seat Number (will be on the RIGHT)
                run_seat = target_p.add_run(seat_display)
                run_seat.font.size = Pt(85)
                run_seat.font.bold = True
                run_seat.font.name = 'Tahoma'

            # DEBUG: שמור כל מסמך נפרד לבדיקה (רק בהרצה מקומית, כדי למנוע קריסה ב-Render)
            if IS_LOCAL_MODE:
                try:
                    import os as _os
                    _debug_dir = r'C:\Users\uri\OneDrive\Desktop\test\debug_docs'
                    _os.makedirs(_debug_dir, exist_ok=True)
                    _safe = e['name'].replace(' ','_')[:20]
                    _dbuf = io.BytesIO()
                    doc.save(_dbuf)
                    _dbuf.seek(0)
                    with open(f'{_debug_dir}\\{i+1:02d}_{_safe}.docx','wb') as _f:
                        _f.write(_dbuf.read())
                    print(f"[DEBUG] Saved person {i+1}: {e['name']} | qr_data={qr_data[:60]}")
                except Exception as _ex:
                    print(f"[DEBUG ERROR] Failed to save debug doc: {_ex}", flush=True)

            # שמור כל דוק ברשימה
            all_docs.append(doc)

        # מזג עם docxcompose â€” מטפל נכון בקשרי תמונות
        from docxcompose.composer import Composer
        from docx.enum.text import WD_BREAK

        master_doc = all_docs[0]
        composer = Composer(master_doc)

        for doc in all_docs[1:]:
            # הוסף מעבר עמוד בסוף המסמך הנוכחי לפני הצירוף
            last_para = master_doc.paragraphs[-1] if master_doc.paragraphs else master_doc.add_paragraph()
            last_para.add_run().add_break(WD_BREAK.PAGE)
            composer.append(doc)

        final_buf = io.BytesIO()
        master_doc.save(final_buf)
        final_buf.seek(0)

        wants_pdf = (request.form.get('format', 'docx') or 'docx').strip().lower() == 'pdf'
        if wants_pdf:
            pdf_bytes = convert_docx_to_pdf(final_buf.getvalue())
            if pdf_bytes:
                return send_file(
                    io.BytesIO(pdf_bytes),
                    as_attachment=True,
                    download_name='טפסי_נבחנים.pdf',
                    mimetype='application/pdf'
                )
            # LibreOffice לא זמין בשרת הזה — לא מפילים את הבקשה, שולחים Word במקום
            flash("המרה ל-PDF לא זמינה בשרת הזה כרגע — נשלח קובץ Word במקום", "warning")
            final_buf.seek(0)

        return send_file(
            final_buf,
            as_attachment=True,
            download_name='טפסי_נבחנים.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

        
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f"שגיאה ביצירת המסמכים: {e}", "danger")
        return redirect(url_for('exam_attendance'))

@app.route('/exam-attendance/print')
@login_required
def exam_attendance_print():
    """דף הדפסת טפסים עם QR — קורא ישירות מגיליון המבחן ב-Drive"""
    exam_filter = request.args.get('exam', '')
    if not exam_filter:
        flash("יש לבחור מבחן להדפסה", "warning")
        return redirect(url_for('exam_attendance'))

    records = load_examinee_cache(exam_filter, force=True)  # תמיד עדכני להדפסה
    examinees = []
    for key, rec in sorted(records.items(), key=lambda kv: kv[1].get('full_name', '')):
        e = dict(rec)
        e.setdefault('id_number', key if key != rec.get('full_name') else '')
        e['exam_name'] = exam_filter
        qr_data = (f"EXAMINEE|{e.get('id_number','')}|{e.get('full_name','')}|"
                   f"{e.get('username','')}|{e.get('password','')}|{e.get('hall','')}|"
                   f"{exam_filter}|{e.get('computer','')}")
        qr_img = qrcode.make(qr_data)
        buf = io.BytesIO()
        qr_img.save(buf)
        buf.seek(0)
        e['qr_b64'] = base64.b64encode(buf.read()).decode('utf-8')
        examinees.append(e)

    if not examinees:
        flash(f"לא נמצאו נבחנים למבחן '{exam_filter}' — ודא שהגיליון קיים וייבאת נבחנים", "warning")
    return render_template('exam_print.html', examinees=examinees, exam_filter=exam_filter)

@app.route('/test-scanner', methods=['GET'])
@login_required
def test_scanner():
    return render_template('test_scanner.html')

@app.route('/simple-scanner', methods=['GET'])
@login_required
def simple_scanner():
    """עמוד סורק פשוט - סורק רק נבחנים"""
    return render_template('simple_scanner.html')

@app.route('/api/simple-scan', methods=['POST'])
@login_required
@scan1_only
def api_simple_scan():
    data = request.json
    qr_text = data.get('qr', '').strip()
    if not qr_text.startswith('EXAMINEE|'):
        return {"error": "QR לא תקין"}, 400

    from exam_naming import parse_examinee_qr, parse_exam_title
    qr = parse_examinee_qr(qr_text)
    if not qr:
        return {"error": "QR לא תקין"}, 400
    id_number = qr['id_number']
    name = qr['full_name']
    canonical_exam = parse_exam_title(qr['exam_title'])['exam_name'] if qr['exam_title'] else ''

    if not canonical_exam:
        return {"error": "QR לא מכיל שם מבחן"}, 400

    technician = session.get('username', '')
    # create=True: ה-QR נושא את כל הפרטים בעצמו (הופק ע"י מחולל
    # התבניות) — הסריקה פותחת/יוצרת את הגיליון הנכון ומכניסה את
    # הנבחן, גם אם לא בוצע ייבוא נפרד מראש
    try:
        target = resolve_exam_sheet(canonical_exam, create=True)
    except Exception as e_sheet:
        return {"error": f"שגיאה בפתיחת/יצירת גיליון למבחן '{canonical_exam}': {e_sheet}"}, 500
    if not target:
        return {"error": f"לא הצלחתי לפתוח/ליצור גיליון למבחן '{canonical_exam}'"}, 500

    try:
        from drive_manager import write_examinee_scan
        write_examinee_scan(target['worksheet'], id_number, full_name=name,
                            technician=technician, is_present=1)
        mark_examinee_scanned(canonical_exam, id_number, name, '', technician, is_present=1)
        return {"success": True, "name": name, "id": id_number}
    except Exception as e:
        return {"error": str(e)}, 500

# â”€â”€ BEACON: GET endpoint לסריקות מהטלפון (עוקף בעיות SSL בכרום) â”€â”€â”€â”€â”€â”€
@app.route('/api/exam-scan-beacon', methods=['GET'])
@login_required
@scan1_only
def api_exam_scan_beacon():
    """
    מקבל נתוני סריקה דרך GET params ומחזיר pixel שקוף.
    הדפדפן תמיד שולח GET לתמונות â€” עוקף בעיות SSL בכרום מובייל.
    """
    qr_text    = (request.args.get('qr', '') or '').strip()
    computer   = (request.args.get('computer', '') or '').strip()
    pc_status  = (request.args.get('pc_status', '') or '').strip()
    is_present = int(request.args.get('is_present', 1) or 1)
    seat       = (request.args.get('seat', '') or '').strip()
    col        = (request.args.get('col', '') or '').strip()
    technician = session.get('username', '')

    from exam_naming import parse_examinee_qr
    qr = parse_examinee_qr(qr_text)
    if qr:
        id_number  = qr['id_number']
        full_name  = qr['full_name']
        # אותו פירוק כמו בסריקה הכפולה — QR של Word בלי כותרת היה
        # נקרא כאן לפי הפריסה הישנה, והמבחן זוהה כמספר כסא
        exam_name_qr = qr['exam_title']
        if not computer and qr['computer']:
            computer = qr['computer']
        if not col and qr['row']:
            col = qr['row']
        if not seat and qr['seat']:
            seat = qr['seat']
        scan_time_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        # שם המבחן חייב להגיע מה-QR — אין מסד לחפש בו ת.ז חוצה-מבחנים
        from exam_naming import parse_exam_title
        canonical_b = parse_exam_title(exam_name_qr)['exam_name'] if exam_name_qr else ''

        def _save():
            if not canonical_b:
                print(f"[BEACON] QR has no exam name — cannot save ({id_number})", flush=True)
                return
            try:
                from drive_manager import write_examinee_scan
                target_b = resolve_exam_sheet(canonical_b, create=True)
                if not target_b:
                    print(f"[BEACON] no sheet for '{canonical_b}'", flush=True)
                    return
                write_examinee_scan(target_b['worksheet'], id_number, full_name=full_name,
                                    computer=computer, col=col, seat=seat,
                                    pc_status=pc_status, scan_time=scan_time_str,
                                    technician=technician, is_present=is_present)
                mark_examinee_scanned(canonical_b, id_number, full_name, computer,
                                      technician, pc_status, is_present)
            except Exception as ex:
                print(f"[BEACON] Save error: {ex}", flush=True)
        threading.Thread(target=_save, daemon=True).start()
    else:
        print(f"[BEACON] âš ï¸  Invalid QR: {qr_text[:40]}")

    # החזר pixel שקוף 1x1
    pixel = base64.b64decode('R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==')
    return send_file(io.BytesIO(pixel), mimetype='image/gif', max_age=0)

@app.route('/api/check-computer-used', methods=['POST'])
@login_required
@scan1_only
def api_check_computer_used():
    """בדיקה אם מחשב כבר שויך לנבחן אחר בגיליון הנוכחי"""
    data = request.json or {}
    computer = (data.get('computer', '') or '').strip()
    exam_name = (data.get('exam_name', '') or '').strip()

    if not computer or not exam_name:
        return jsonify({"in_use": False})

    try:
        from drive_manager import find_header_row, _col_index

        target = resolve_exam_sheet(exam_name, create=False)
        if not target:
            return jsonify({"in_use": False})

        all_rows = target['worksheet'].get_all_values()
        hdr_idx = find_header_row(all_rows)
        if hdr_idx == -1:
            return jsonify({"in_use": False})

        headers = [str(h).strip() for h in all_rows[hdr_idx]]
        computer_col = _col_index(headers, ['מ.מחשב', 'מחשב', 'computer'])
        presence_col = _col_index(headers, ['נוכחות', 'הגיע', 'attendance'])
        name_col     = _col_index(headers, ['שם'])
        if computer_col is None or presence_col is None:
            return jsonify({"in_use": False})

        for r in all_rows[hdr_idx + 1:]:
            if computer_col < len(r) and str(r[computer_col]).strip() == str(computer).strip():
                # משויך רק אם גם סומנה נוכחות באותה שורה
                if presence_col < len(r) and r[presence_col].strip():
                    existing_name = (r[name_col].strip()
                                     if name_col is not None and name_col < len(r) and r[name_col].strip()
                                     else 'נבחן')
                    return jsonify({"in_use": True, "name": existing_name})
        return jsonify({"in_use": False})
    except Exception as ex:
        print(f"[check-computer-used] Error: {ex}", flush=True)
        return jsonify({"in_use": False})

@app.route('/api/exam-scan-double', methods=['POST'])
@login_required
@scan1_only
def api_exam_scan_double():
    """סריקה כפולה: נבחן + מחשב + סטטוסים — מקבל JSON או form data"""
    if request.is_json:
        data = request.json
    else:
        data = request.form
    qr_text   = (data.get('qr', '') or '').strip()
    computer  = (data.get('computer', '') or '').strip()
    seat      = (data.get('seat', '') or '').strip()
    col       = (data.get('col', '') or '').strip()
    pc_status = (data.get('pc_status', '') or '').strip()
    is_present = int(data.get('is_present', 1) or 1)

    # פירוק ה-QR — לוגיקה משותפת לכל נקודות הסריקה (exam_naming)
    from exam_naming import parse_examinee_qr
    qr = parse_examinee_qr(qr_text)
    if not qr:
        return {"error": "QR לא מזוהה כנבחן"}, 400

    id_number   = qr['id_number']
    full_name   = qr['full_name']
    qr_username = qr['username']
    qr_password = qr['password']
    # col/seat: מה-form אם נשלח, אחרת מה-QR
    if not col  and qr['row']:  col  = qr['row']
    if not seat and qr['seat']: seat = qr['seat']
    if not computer and qr['computer']: computer = qr['computer']

    # שם המבחן — מהפרמטר או מה-QR. בלי מסד אי אפשר לחפש "לאיזה מבחן
    # ת.ז זו שייכת" בלי לדעת את שם המבחן קודם — ה-QR תמיד נושא אותו
    # (exam_naming.parse_examinee_qr דואג לזה בשני הפורמטים).
    exam_name = (data.get('exam_name', '') or '').strip() or qr['exam_title']

    scan_time_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    technician = session.get('username', '')

    from exam_naming import parse_exam_title
    exam_info = parse_exam_title(exam_name)
    canonical_exam = exam_info['exam_name'] or exam_name

    if not exam_name:
        return jsonify({"success": False, "error": "לא זוהה שם מבחן ב-QR"}), 400

    # ── איתור הגיליון ב-Drive — מקור האמת היחיד ──
    # create=True: ה-QR (מדף הנבחן שהודפס) נושא את כל הפרטים בעצמו,
    # אז הסריקה פותחת/יוצרת את התיקייה/הגיליון הנכונים לפי שם המבחן
    # ומכניסה את הנבחן — בלי צורך בייבוא אקסל נפרד מראש
    target = None
    sheet_error = None
    try:
        target = resolve_exam_sheet(exam_name, create=True)
    except Exception as e_sheet:
        sheet_error = str(e_sheet)
        print(f"[SCAN] sheet lookup/creation failed: {e_sheet}", flush=True)

    if not target:
        msg = f"לא הצלחתי לפתוח/ליצור גיליון למבחן '{canonical_exam}'"
        if sheet_error:
            msg += f" ({sheet_error})"
        return jsonify({
            "success": False, "in_roster": False, "sheet": False,
            "error": msg
        }), 500

    sheet_id = target['sheet_id']

    # ── האם הנבחן בכלל רשום למבחן הזה? (מהמטמון, שנטען מהגיליון) ──
    in_roster = find_examinee(canonical_exam, id_number=id_number, full_name=full_name) is not None
    if not in_roster:
        print(f"[SCAN] {id_number} not in roster for '{canonical_exam}' — saving anyway", flush=True)

    # שמירת מפתח הגיליון והבחינה בסשן לטובת ביטול
    session['last_exam_sheet_id'] = sheet_id
    session['last_exam_name'] = exam_name
    session['last_id_number'] = id_number

    # עדכון המטמון בזיכרון מיידית — כדי שסריקה הבאה תראה את זה בלי
    # לחכות לכתיבה בפועל לגיליון (שקורית ב-thread ברקע)
    mark_examinee_scanned(canonical_exam, id_number, full_name, computer,
                          technician, pc_status, is_present)

    def save_to_exam_sheet_in_thread(ws, target_exam_name, target_id_number, target_full_name,
                                     target_computer, target_col, target_seat, target_pc_status,
                                     target_scan_time, target_technician, target_username='',
                                     target_password=''):
        # ws מגיע מוכן מהבקשה — אותו גיליון בדיוק שהייבוא כתב אליו
        try:
            from drive_manager import write_examinee_scan
            write_examinee_scan(ws, target_id_number, full_name=target_full_name,
                                computer=target_computer, col=target_col, seat=target_seat,
                                pc_status=target_pc_status, scan_time=target_scan_time,
                                technician=target_technician, is_present=is_present,
                                username=target_username, password=target_password)
        except Exception as ex:
            print(f"[THREAD ERROR] Save failed for {target_full_name}: {ex}", flush=True)
            import traceback; traceback.print_exc()

    import threading
    t = threading.Thread(
        target=save_to_exam_sheet_in_thread,
        args=(target['worksheet'], canonical_exam, id_number, full_name, computer, col, seat, pc_status, scan_time_str, technician, qr_username, qr_password),
        daemon=False
    )
    t.start()

    return jsonify({"success": True, "in_roster": in_roster, "exam": canonical_exam})

@app.route('/api/undo-last-scan', methods=['POST'])
@login_required
@scan1_only
def undo_last_scan():
    """ביטול סריקה אחרונה וניקוי סטטוס נוכחות בגיליון ב-Drive"""
    try:
        from drive_manager import find_header_row, map_columns, _a1_col, SCAN_FIELDS

        exam_name = session.get('last_exam_name')
        last_id_number = session.get('last_id_number')
        if not last_id_number:
            return {"success": False, "error": "לא נמצאה סריקה לביטול"}

        from exam_naming import parse_exam_title
        canonical_exam = parse_exam_title(exam_name)['exam_name'] if exam_name else ''

        # הסרה מהמטמון בזיכרון — בלי למחוק את מ.מחשב שהגיע מהייבוא
        # (הסריקה מבטלת נוכחות, לא הקצאת מחשב)
        if canonical_exam and canonical_exam in examinee_cache:
            rec = (examinee_cache[canonical_exam].get(last_id_number)
                   if last_id_number in examinee_cache[canonical_exam] else None)
            if rec:
                rec['is_present'] = '0'
                rec['scan_time'] = ''
                rec['technician'] = ''
                rec['pc_status'] = ''
                print(f"[UNDO CACHE] Cleared {last_id_number} in cache for {canonical_exam}", flush=True)

        target = resolve_exam_sheet(exam_name, create=False) if exam_name else None
        if not target:
            return {"success": True, "sheet": False}

        ws = target['worksheet']
        all_vals = ws.get_all_values()
        hdr_idx = find_header_row(all_vals)
        if hdr_idx == -1:
            return {"success": True, "sheet": False}

        headers = [str(h).strip() for h in all_vals[hdr_idx]]
        mapping = map_columns(headers)
        id_col = mapping.get('id_number')
        if id_col is None:
            return {"success": True, "sheet": False}

        # עמודות הסריקה לפי הכותרות בפועל — לא מיקום קבוע.
        # מ.מחשב נשמר, בדיוק כמו במסד: הוא עשוי להגיע מרשימת הייבוא,
        # וביטול סריקה מבטל נוכחות — לא הקצאת מחשב.
        scan_cols = sorted(mapping[f] for f in SCAN_FIELDS
                           if f in mapping and f != 'computer')

        for idx, r in enumerate(all_vals[hdr_idx + 1:], start=hdr_idx + 2):
            if id_col < len(r) and str(r[id_col]).strip() == str(last_id_number).strip():
                # מנקה רק את עמודות הסריקה — שורת הנבחן עצמה נשארת
                blanks = [{'range': f'{_a1_col(c)}{idx}', 'values': [['']]}
                          for c in scan_cols]
                if blanks:
                    ws.batch_update(blanks, value_input_option='USER_ENTERED')
                print(f"[UNDO OK] Cleared scan columns on row {idx} for ID {last_id_number}", flush=True)
                return {"success": True, "sheet": True}

        # לא נמצאה שורה — לא מוחקים שום דבר אחר, זה היה מוחק נבחן תמים
        return {"success": True, "sheet": False}
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.route('/api/exam-scan', methods=['POST'])
@login_required
@scan1_only
def api_exam_scan():
    """סריקת QR לנוכחות"""
    data = request.json
    qr_text = data.get('qr', '').strip()
    if not qr_text:
        return {"error": "לא התקבל QR"}, 400

    from exam_naming import parse_examinee_qr, parse_exam_title
    qr = parse_examinee_qr(qr_text)
    if not qr or not qr['id_number']:
        return {"error": "QR לא מזוהה כנבחן", "type": "unknown"}, 400

    id_number = qr['id_number']
    canonical_exam = parse_exam_title(qr['exam_title'])['exam_name'] if qr['exam_title'] else ''
    if not canonical_exam:
        return {"error": "QR לא מכיל שם מבחן"}, 400

    full_name = qr['full_name']
    examinee = find_examinee(canonical_exam, id_number=id_number, full_name=full_name) or {}
    if str(examinee.get('is_present', '')) in ('1', 'True', 'true'):
        return {"success": True, "already": True, "examinee": examinee}

    technician = session.get('username', '')
    # create=True: ה-QR נושא את כל הפרטים בעצמו — הסריקה פותחת/יוצרת
    # את הגיליון הנכון גם בלי ייבוא אקסל קודם
    try:
        target = resolve_exam_sheet(canonical_exam, create=True)
    except Exception as e_sheet:
        return {"error": f"שגיאה בפתיחת/יצירת גיליון למבחן '{canonical_exam}': {e_sheet}"}, 500
    if not target:
        return {"error": f"לא הצלחתי לפתוח/ליצור גיליון למבחן '{canonical_exam}'"}, 500

    try:
        from drive_manager import write_examinee_scan
        write_examinee_scan(target['worksheet'], id_number,
                            full_name=examinee.get('full_name') or full_name,
                            technician=technician, is_present=1)
        mark_examinee_scanned(canonical_exam, id_number, examinee.get('full_name') or full_name,
                              '', technician, is_present=1)
    except Exception as e:
        return {"error": str(e)}, 500

    examinee['is_present'] = True
    examinee['attend_time'] = datetime.now().strftime("%H:%M:%S")
    examinee.setdefault('full_name', full_name)
    examinee.setdefault('id_number', id_number)
    return {"success": True, "already": False, "examinee": examinee}

@app.route('/exam-attendance/delete/<int:eid>', methods=['POST'])
@login_required
def exam_attendance_delete(eid):
    """
    מחיקת נבחן — מנהל בלבד.
    eid נשאר לתאימות ה-URL אך אינו מפתח מסד (אין מסד לנבחנים יותר) —
    הזיהוי בפועל הוא exam_name + id_number מהטופס.
    """
    if session.get('role') != 'admin':
        flash("אין הרשאה למחיקת נבחן", "danger")
        return redirect(url_for('exam_attendance'))

    exam_name = request.form.get('exam_name', '').strip()
    id_number = request.form.get('id_number', '').strip()
    if not exam_name or not id_number:
        flash("חסר שם מבחן או ת.ז למחיקה", "danger")
        return redirect(url_for('exam_attendance'))

    try:
        from drive_manager import find_header_row, map_columns
        target = resolve_exam_sheet(exam_name, create=False)
        if not target:
            flash(f"לא נמצא גיליון למבחן '{exam_name}'", "danger")
            return redirect(url_for('exam_attendance'))
        ws = target['worksheet']
        all_vals = ws.get_all_values()
        hdr_idx = find_header_row(all_vals)
        if hdr_idx == -1:
            flash("לא נמצאה שורת כותרות בגיליון", "danger")
            return redirect(url_for('exam_attendance'))
        mapping = map_columns([str(h).strip() for h in all_vals[hdr_idx]])
        id_col = mapping.get('id_number')
        if id_col is not None:
            for idx, row in enumerate(all_vals[hdr_idx + 1:], start=hdr_idx + 2):
                if id_col < len(row) and str(row[id_col]).strip() == id_number:
                    ws.delete_rows(idx)
                    invalidate_examinee_cache(exam_name)
                    flash("הנבחן נמחק בהצלחה", "success")
                    return redirect(url_for('exam_attendance'))
        flash("הנבחן לא נמצא בגיליון", "warning")
    except Exception as e:
        flash(f"שגיאה במחיקה: {e}", "danger")
    return redirect(url_for('exam_attendance'))

@app.route('/exam-attendance/clear', methods=['POST'])
@login_required
def exam_attendance_clear():
    """
    איפוס נוכחות למבחן ספציפי (לפני בחינה חדשה) — מנהל בלבד.
    דורש exam_name — בלי מסד אין דרך מעשית לאפס נוכחות בכל המבחנים
    בכל תיקיות המשרדים בבת אחת.
    """
    if session.get('role') != 'admin':
        flash("אין הרשאה לפעולה זו", "danger")
        return redirect(url_for('exam_attendance'))

    exam_name = request.form.get('exam_name', '').strip()
    if not exam_name:
        flash("יש לבחור מבחן לאיפוס נוכחות", "danger")
        return redirect(url_for('exam_attendance'))

    try:
        from drive_manager import find_header_row, map_columns, _a1_col, SCAN_FIELDS
        target = resolve_exam_sheet(exam_name, create=False)
        if not target:
            flash(f"לא נמצא גיליון למבחן '{exam_name}'", "danger")
            return redirect(url_for('exam_attendance'))
        ws = target['worksheet']
        all_vals = ws.get_all_values()
        hdr_idx = find_header_row(all_vals)
        if hdr_idx == -1:
            flash("לא נמצאה שורת כותרות בגיליון", "danger")
            return redirect(url_for('exam_attendance'))
        mapping = map_columns([str(h).strip() for h in all_vals[hdr_idx]])
        scan_cols = [mapping[f] for f in SCAN_FIELDS if f in mapping and f != 'computer']
        blanks = []
        for r_idx in range(hdr_idx + 2, len(all_vals) + 1):
            for c in scan_cols:
                blanks.append({'range': f'{_a1_col(c)}{r_idx}', 'values': [['']]})
        if blanks:
            ws.batch_update(blanks, value_input_option='USER_ENTERED')
        invalidate_examinee_cache(exam_name)
        flash(f"✅ הנוכחות אופסה למבחן '{exam_name}' – מוכן לבחינה חדשה!", "success")
    except Exception as e:
        flash(f"שגיאה באיפוס: {e}", "danger")
    return redirect(url_for('exam_attendance'))

@app.route('/exam-attendance/scanner')
@login_required
@scan1_only
def exam_attendance_scanner():
    """עמוד סריקת נוכחות"""
    return render_template('exam_scanner.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# --- Standalone Exam Form Generator ---

@app.route('/exam-generator')
@login_required
def exam_generator():
    """דף העלאת אקסל ליצירת טפסי בחינה"""
    return render_template('exam_generator.html')

@app.route('/generate-forms', methods=['POST'])
@login_required
def generate_forms():
    """מעבד אקסל ומחזיר דף מוכן להדפסה עם QR"""
    excel_file = request.files.get('excel_file')
    if not excel_file:
        return "לא הועלה קובץ", 400

    try:
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active

        headers = [str(cell.value).strip() if cell.value else '' for cell in ws[1]]

        col_map = {}
        for i, h in enumerate(headers):
            h_l = h.lower()
            if 'שם' in h and 'נבחן' in h: col_map['name'] = i
            elif 'שם' in h and col_map.get('name') is None: col_map['name'] = i
            elif 'זהות' in h or 'ת.ז' in h or 'id' in h_l: col_map['id_number'] = i
            elif 'משתמש' in h or 'קוד' in h or 'user' in h_l: col_map['username'] = i
            elif 'סיסמ' in h or 'pass' in h_l: col_map['password'] = i
            elif 'התאמ' in h or 'notes' in h_l: col_map['notes'] = i
            elif 'מחשב' in h or 'computer' in h_l or 'מספר' in h: col_map['computer'] = i
            elif 'מיקום' in h or 'כיתה' in h or 'location' in h_l: col_map['location'] = i

        students = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row): continue
            def get_val(key):
                idx = col_map.get(key)
                if idx is None: return ''
                v = row[idx]
                return str(v).strip() if v is not None else ''

            name = get_val('name')
            if not name or name == 'None': continue

            students.append({
                'name': name,
                'id_number': get_val('id_number'),
                'username': get_val('username'),
                'password': get_val('password'),
                'notes': get_val('notes'),
                'computer': get_val('computer'),
                'location': get_val('location'),
            })

        return render_template('exam_forms_print.html', students=students)

    except Exception as e:
        import traceback; traceback.print_exc()
        return f"שגיאה: {e}", 500

# â”€â”€ FAULT REPORT: טופס תקלות מחשב â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@app.route('/fault-report', methods=['GET'])
@login_required
def fault_report_page():
    """עמוד טופס דיווח תקלה"""
    # שלוף רשימת מחשבים לרשימת auto-complete
    conn = get_db_connection()
    barcodes = []
    if conn:
        try:
            cur = get_safe_cursor(conn)
            cur.execute("SELECT barcode, case_number, location FROM computers ORDER BY barcode")
            barcodes = [dict(r) for r in cur.fetchall()]
            cur.close()
        except Exception:
            pass
        finally:
            release_db_connection(conn)
    return render_template('fault_report.html', barcodes=barcodes)

@app.route('/api/submit-fault', methods=['POST'])
@login_required
def api_submit_fault():
    """קבלת דיווח תקלה â€” שומר לשיטס + להיסטוריה"""
    data = request.json or {}
    barcode     = (data.get('barcode', '') or '').strip()
    fault_type  = (data.get('fault_type', '') or '').strip()
    description = (data.get('description', '') or '').strip()
    location    = (data.get('location', '') or '').strip()
    technician  = session.get('username', '')
    report_time = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    if not barcode:
        return jsonify({"success": False, "error": "חובה להזין מספר מחשב"}), 400
    if not description:
        return jsonify({"success": False, "error": "חובה לתאר את התקלה"}), 400

    # â”€â”€ שמירה לגיליון 'תקלות' בשיטס â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def save_fault_to_sheets():
        import traceback
        print(f"[FAULT] â–¶ï¸ שומר תקלה: מחשב {barcode} | {fault_type}", flush=True)
        try:
            import gspread
            scopes   = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            sheet_id = os.getenv('GOOGLE_SHEETS_ID')
            if not sheet_id:
                print("[FAULT] âš ï¸ GOOGLE_SHEETS_ID לא מוגדר", flush=True)
                return
            creds  = get_google_creds(scopes)
            client = gspread.authorize(creds)
            sh     = client.open_by_key(sheet_id)
            # פתח/צור גיליון 'תקלות'
            try:
                ws = sh.worksheet('תקלות')
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title='תקלות', rows='500', cols='8')
                ws.append_row(['תאריך', 'מחשב', 'סוג תקלה', 'תיאור', 'מיקום', 'טכנאי', 'סטטוס טיפול'],
                              value_input_option='USER_ENTERED')
                ws.format('A1:G1', {'textFormat': {'bold': True},
                                    'backgroundColor': {'red': 1.0, 'green': 0.85, 'blue': 0.4}})
            ws.append_row([report_time, barcode, fault_type, description, location, technician, 'ממתין לטיפול'],
                          value_input_option='USER_ENTERED')
            print(f"[FAULT] âœ… נשמר לגיליון תקלות: {barcode}", flush=True)
        except Exception as ex:
            print(f"[FAULT] âŒ שגיאה בשמירה לשיטס: {ex}", flush=True)
            traceback.print_exc()

    threading.Thread(target=save_fault_to_sheets, daemon=False).start()

    # â”€â”€ שמירה להיסטוריה בDB â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    conn = get_db_connection()
    if conn:
        try:
            cur = get_safe_cursor(conn)
            cur.execute("SELECT id FROM computers WHERE barcode = %s", (barcode,))
            comp = cur.fetchone()
            if comp:
                comp_id = comp['id']
                note_text = f"[תקלה] {fault_type}: {description}"
                cur.execute("""
                    INSERT INTO inventory_history (computer_id, technician, change_type, old_value, new_value)
                    VALUES (%s, %s, 'דיווח תקלה', %s, %s)
                """, (comp_id, technician,
                       f"מיקום: {location}",
                       note_text))
                conn.commit()
            cur.close()
        except Exception as e:
            print(f"[FAULT] DB history error: {e}", flush=True)
        finally:
            release_db_connection(conn)

    return jsonify({"success": True, "message": f"תקלה דווחה בהצלחה עבור מחשב {barcode}"})

@app.route('/api/sheets-sync-status')
@login_required
def api_sheets_sync_status():
    """מחזיר זמן הסנכרון האחרון מגיליון שיטס"""
    global _last_sheets_import
    if _last_sheets_import:
        diff = (datetime.now() - _last_sheets_import).seconds
        if diff < 60:
            ago = f"לפני {diff} שניות"
        else:
            ago = f"לפני {diff // 60} דקות"
        return jsonify({"last_sync": _last_sheets_import.strftime('%H:%M:%S'), "ago": ago})
    return jsonify({"last_sync": None, "ago": "טרם סונכרן"})

# â”€â”€ CAGE INFO PAGE (FOR MOBILE/PHONE QR SCAN) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
@app.route('/cage-info/<cage_id>')
@login_required
def cage_info_page(cage_id):
    conn = get_db_connection()
    if not conn: return "DB Error", 500
    try:
        cur = get_safe_cursor(conn)
        
        # Get cage details
        cur.execute("SELECT * FROM cages WHERE cage_id = %s", (cage_id,))
        cage = cur.fetchone()
        
        # Get computers in cage
        cur.execute("""
            SELECT id, barcode, case_number, status, location, specs, scan_time, notes
            FROM computers
            WHERE cage_number = %s OR cage_name = %s
            ORDER BY scan_time DESC NULLS LAST
        """, (cage_id, cage_id))
        computers = cur.fetchall()
        
        cur.close()
        
        if not cage:
            cage = {
                'cage_id': cage_id,
                'name': f'כלוב {cage_id}',
                'location': '',
                'notes': 'כלוב זה נוצר אוטומטית בעת סריקת מחשבים.'
            }
            
        return render_template('cage_info.html', cage=cage, computers=computers, total=len(computers))
    except Exception as e:
        print(f"Error in cage_info_page: {e}")
        return f"<h1>Error: {e}</h1>", 500
    finally:
        release_db_connection(conn)

@app.route('/cages/print/<cage_id>')
@login_required
def print_cage_page(cage_id):
    conn = get_db_connection()
    if not conn: return "DB Error", 500
    try:
        cur = get_safe_cursor(conn)
        
        # Get cage details
        cur.execute("SELECT * FROM cages WHERE cage_id = %s", (cage_id,))
        cage = cur.fetchone()
        
        # Get computers in cage
        cur.execute("""
            SELECT barcode, case_number, status, location
            FROM computers
            WHERE cage_number = %s OR cage_name = %s
            ORDER BY scan_time DESC NULLS LAST
        """, (cage_id, cage_id))
        computers = cur.fetchall()
        
        cur.close()
        
        if not cage:
            cage = {
                'cage_id': cage_id,
                'name': f'כלוב {cage_id}',
                'location': '',
                'notes': ''
            }
            
        return render_template('print_cage.html', cage=cage, computers=computers, total=len(computers))
    except Exception as e:
        print(f"Error in print_cage_page: {e}")
        return f"<h1>Error: {e}</h1>", 500
    finally:
        release_db_connection(conn)

# --- End of Routes ---

if __name__ == '__main__':
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        local_ip = "127.0.0.1"

    print("\n[START] URI SYSTEM IS LIVE! (HTTPS ENABLED FOR MOBILE SCANNER)")
    print("URL Link: https://127.0.0.1:5000")
    print(f"Mobile Link: https://{local_ip}:5000\n")
    print("[WARNING] When opening on iPhone, you will see a 'Not Private' warning. Click 'Show Details' -> 'Visit this website' to bypass and test the scanner.")
    import os
    cert_file = os.path.join(os.path.dirname(__file__), 'server.crt')
    key_file  = os.path.join(os.path.dirname(__file__), 'server.key')
    if os.path.exists(cert_file) and os.path.exists(key_file):
        ssl_ctx = (cert_file, key_file)
        print("[OK] Using custom SSL certificate (valid for local network)")
    else:
        ssl_ctx = 'adhoc'
        print("[WARN] Custom cert not found, using adhoc SSL")
    app.run(host='0.0.0.0', debug=IS_LOCAL_MODE, port=5000, ssl_context=ssl_ctx)
