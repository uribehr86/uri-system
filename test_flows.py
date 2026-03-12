import urllib.request
import ssl
import json
import urllib.parse
import sys

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

print("[TEST 0] Starting test suite")

# 1. Login
try:
    data = urllib.parse.urlencode({'username': 'uri_behr', 'password': '123'}).encode()
    req = urllib.request.Request('https://127.0.0.1:5000/login', data=data)
    resp = urllib.request.urlopen(req, context=ctx)
    cookie = resp.getheader('Set-Cookie')
    print("[PASS] Logged in successfully")
except Exception as e:
    print(f"[FAIL] Login failed: {e}")
    sys.exit(1)

# 2. Fetch Computers Table
try:
    req = urllib.request.Request('https://127.0.0.1:5000/computers', headers={'Cookie': cookie})
    resp = urllib.request.urlopen(req, context=ctx)
    html = resp.read().decode('utf-8')
    if 'מבחן\ערעור' in html and 'לוח בקרה ומלאי' in html and 'id="batchActionBar"' in html:
        print("[PASS] Computers Table page loads properly with search and forms")
    else:
        print("[FAIL] Table might be missing elements")
except Exception as e:
    print(f"[FAIL] computers route threw exception: {e}")

# 3. Test Search Logic
try:
    req = urllib.request.Request('https://127.0.0.1:5000/computers?q=4', headers={'Cookie': cookie})
    resp = urllib.request.urlopen(req, context=ctx)
    html_search = resp.read().decode('utf-8')
    if 'מבחן\ערעור' in html_search:
        print("[PASS] Search functionality is working and rendering table")
    else:
        print("[FAIL] Search functionality failed")
except Exception as e:
    print(f"[FAIL] Search route threw exception: {e}")

# 4. Test Scanner endpoint loads
try:
    req = urllib.request.Request('https://127.0.0.1:5000/scanner', headers={'Cookie': cookie})
    resp = urllib.request.urlopen(req, context=ctx)
    if resp.status == 200:
        print("[PASS] Scanner page loads successfully")
except Exception as e:
    print(f"[FAIL] Scanner page failed: {e}")
