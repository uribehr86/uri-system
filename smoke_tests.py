import flask_app
app = flask_app.app
app.config['TESTING'] = True
client = app.test_client()

# Default user
client.post('/login', data={'username': 'admin_uri', 'password': '111'}, follow_redirects=True)

routes_to_test = [
    '/portal',
    '/computers',
    '/scanner',
    '/history',
    '/export/computers',
    '/manage-users'
]

print("=== RUNNING INTERNAL ENDPOINT TESTS ===")
for route in routes_to_test:
    res = client.get(route, follow_redirects=True)
    print(f"Testing {route} -> Status Code: {res.status_code}")
    if res.status_code != 200:
        print(f"   [!] Failed! Output: {res.data[:200]}")
    elif b"500 Internal Server Error" in res.data or b"Exception" in res.data:
        print(f"   [!] Rendered 500 equivalent!")

print("=== DONE ===")
