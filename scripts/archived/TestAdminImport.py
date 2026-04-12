"""
TestAdminImport.py
------------------
Tests whether the GlueUp Admin UI import endpoint can be called
programmatically using the existing API token (from GetToken.py).

Usage:
    python TestAdminImport.py <path_to_import_xlsx>

    The XLS file should be a valid GlueUp import file produced by
    ICFGlueUpSync.py (even a 1-row file is fine for testing).

What this script tests:
    Attempt 1 — token header only (same as all other API scripts)
    Attempt 2 — token header + Origin/Referer headers (browser simulation)
    Attempt 3 — cookie-based session (requires browser session capture)

    Each attempt reports the HTTP status code and response body so you
    can see exactly what GlueUp returns, whether it works or not.

Prerequisites:
    pip install requests openpyxl
    Run GetToken.py first to produce glueup_token.json
"""

import sys
import os
import json
import hmac
import hashlib
import time
import requests

# ── Config ─────────────────────────────────────────────────────────────────────

BASE_URL   = 'https://api-services.glueup.com'
ADMIN_URL  = 'https://app.glueup.com'          # Admin UI base (may differ)
ORG_ID     = '7912'
MEMBER_TYPE_ID = '37600'                        # ICF Chapter Member
TOKEN_FILE = 'glueup_token.json'

PK = 'icfwshts'
SK = 'MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPfXb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfUCCAP67vZiAQ55'

# ── Helpers ────────────────────────────────────────────────────────────────────

def build_a_header(method: str) -> str:
    ts  = str(int(time.time() * 1000))
    msg = method + PK + '1.0' + ts
    d   = hmac.new(SK.encode(), msg.encode(), hashlib.sha256).hexdigest()
    return f'v=1.0;k={PK};ts={ts};d={d}'


def load_token() -> str:
    if not os.path.exists(TOKEN_FILE):
        print(f'✗ {TOKEN_FILE} not found. Run GetToken.py first.')
        sys.exit(1)
    with open(TOKEN_FILE) as f:
        return json.load(f)['token']


def print_response(label: str, resp: requests.Response):
    print(f'\n{"─"*60}')
    print(f'  {label}')
    print(f'{"─"*60}')
    print(f'  Status : {resp.status_code} {resp.reason}')
    print(f'  Headers: {dict(resp.headers)}')
    body = resp.text.strip()
    if body:
        # Try to pretty-print JSON
        try:
            print(f'  Body   : {json.dumps(json.loads(body), indent=4)}')
        except Exception:
            # Truncate HTML responses (login redirect pages etc.)
            print(f'  Body   : {body[:800]}{"..." if len(body) > 800 else ""}')
    else:
        print('  Body   : (empty)')


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print('Usage: python TestAdminImport.py <path_to_import_xlsx>')
        print('Example: python TestAdminImport.py glueup_import_20260330_1200.xlsx')
        sys.exit(1)

    xls_path = sys.argv[1]
    if not os.path.exists(xls_path):
        print(f'✗ File not found: {xls_path}')
        sys.exit(1)

    token = load_token()
    print(f'✓ Token loaded from {TOKEN_FILE}')
    print(f'✓ Import file: {xls_path} ({os.path.getsize(xls_path):,} bytes)')

    # The endpoint the Admin UI uses when you click Import
    # Observed via browser DevTools — POST with multipart/form-data
    endpoint_v1 = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/ajax'
    endpoint_v2 = f'https://app.glueup.com/admin/memberships/import/{MEMBER_TYPE_ID}/ajax'

    # ── Attempt 1: token header only (same as all working API calls) ───────────
    print(f'\n{"═"*60}')
    print('ATTEMPT 1: API token header only')
    print(f'Endpoint: {endpoint_v1}')
    print(f'{"═"*60}')

    with open(xls_path, 'rb') as f:
        files = {'file': (os.path.basename(xls_path), f,
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        headers = {
            'a':                     build_a_header('POST'),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                                     'AppleWebKit/537.36 (KHTML, like Gecko) '
                                     'Chrome/120.0.0.0 Safari/537.36',
        }
        try:
            resp = requests.post(endpoint_v1, headers=headers, files=files, timeout=30)
            print_response('API base URL + token header', resp)
        except requests.exceptions.RequestException as e:
            print(f'  ✗ Request failed: {e}')

    # ── Attempt 2: token + browser-like headers on app.glueup.com domain ──────
    print(f'\n{"═"*60}')
    print('ATTEMPT 2: token + Origin/Referer (browser simulation)')
    print(f'Endpoint: {endpoint_v2}')
    print(f'{"═"*60}')

    with open(xls_path, 'rb') as f:
        files = {'file': (os.path.basename(xls_path), f,
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        headers = {
            'a':                     build_a_header('POST'),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'Origin':                'https://app.glueup.com',
            'Referer':               f'https://app.glueup.com/admin/memberships/import/{MEMBER_TYPE_ID}',
            'X-Requested-With':      'XMLHttpRequest',
            'User-Agent':            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                                     'AppleWebKit/537.36 (KHTML, like Gecko) '
                                     'Chrome/120.0.0.0 Safari/537.36',
        }
        try:
            resp = requests.post(endpoint_v2, headers=headers, files=files, timeout=30)
            print_response('app.glueup.com + browser headers', resp)
        except requests.exceptions.RequestException as e:
            print(f'  ✗ Request failed: {e}')

    # ── Attempt 3: probe endpoint without file (GET) to see what it returns ───
    print(f'\n{"═"*60}')
    print('ATTEMPT 3: GET probe — does the endpoint exist at all?')
    print(f'{"═"*60}')

    for url in [endpoint_v1, endpoint_v2]:
        headers = {
            'a':                     build_a_header('GET'),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0',
        }
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            print_response(f'GET {url}', resp)
        except requests.exceptions.RequestException as e:
            print(f'  ✗ {url} — {e}')

    print(f'\n{"═"*60}')
    print('INTERPRETATION GUIDE')
    print(f'{"═"*60}')
    print('''
  200 + JSON success  → It works! Update the design spec and build
                        AdminImport.py as a production script.

  200 + HTML page     → Redirected to login page. Endpoint requires
                        browser session cookies, not API token.
                        Next step: capture cookies via Playwright.

  401 Unauthorized    → Auth rejected. Token header not accepted here.
                        Try cookie-based session (Playwright path).

  403 Forbidden       → Endpoint exists but access denied for this
                        account/method. May need different credentials
                        or CSRF token.

  404 Not Found       → Wrong URL. Check DevTools on the live Admin UI
                        to confirm the exact endpoint path.

  405 Method Not      → Same as the write API — endpoint locked down.
      Allowed           This approach will not work.

  Connection error    → Domain unreachable from this machine. Try the
                        other base URL (api-services vs app.glueup.com).
''')


if __name__ == '__main__':
    main()