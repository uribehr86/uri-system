import sys, io, urllib.request, urllib.parse, json, ssl, http.cookiejar
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

ctx = ssl._create_unverified_context()
jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(
    urllib.request.HTTPSHandler(context=ctx),
    urllib.request.HTTPCookieProcessor(jar)
)

# 1. Login
login_data = urllib.parse.urlencode({'username': 'Admin_uri', 'password': 'uri*'}).encode()
try:
    r = opener.open('https://127.0.0.1:5000/login', login_data, timeout=5)
    print("Login status:", r.url)
except Exception as e:
    print("Login error:", e)

# 2. Test scan
scan_data = json.dumps({
    'qr': 'בדיקה|123456789|ישראל ישראלי|test_u|test_p|1|1',
    'computer': '9999',
    'pc_status': 'תקין',
    'is_present': 1,
    'col': '1',
    'seat': '1'
}).encode('utf-8')

req = urllib.request.Request(
    'https://127.0.0.1:5000/api/exam-scan-double',
    data=scan_data,
    headers={'Content-Type': 'application/json'},
    method='POST'
)
try:
    r = opener.open(req, timeout=15)
    body = r.read().decode('utf-8')
    print("Scan STATUS:", r.status)
    print("Scan BODY:", body)
except Exception as e:
    print("Scan ERROR:", e)
