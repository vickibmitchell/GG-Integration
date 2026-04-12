"""
AdminImportCookie.py
--------------------
Imports a GlueUp-formatted XLS file via the Admin UI import endpoint
using a PHP session cookie captured from your real browser.

This approach avoids Cloudflare bot detection entirely since you log
in once manually in your normal browser, then hand the session cookie
to this script.

HOW TO GET YOUR SESSION COOKIE (one-time setup, ~1 minute):
    1. Log in to https://app.glueup.com in Chrome or Safari.
    2. Open DevTools: Cmd+Option+I (Mac) or F12 (Windows).
    3. Click the "Application" tab (Chrome) or "Storage" tab (Safari).
    4. In the left panel expand Cookies → https://app.glueup.com
    5. Find the cookie named PHPSESSID and copy its Value.
    6. Paste it when this script prompts for it, or set the
       GLUEUP_PHPSESSID environment variable.

HOW LONG DOES THE COOKIE LAST?
    GlueUp PHP sessions typically last until the browser is closed or
    the server expires them (usually 24 hours to a few days).
    If the script reports a login redirect, the cookie has expired —
    just grab a fresh one from your browser.

Usage:
    python AdminImportCookie.py <path_to_import_xlsx>

Environment variables (optional — script prompts if not set):
    GLUEUP_PHPSESSID   your PHPSESSID cookie value

Prerequisites:
    pip install requests
"""

import sys
import os
import json
import getpass
import requests
from datetime import datetime

# ── Config ─────────────────────────────────────────────────────────────────────

ADMIN_BASE      = 'https://app.glueup.com'
ORG_ID          = '7912'
MEMBER_TYPE_ID  = '37600'
IMPORT_ENDPOINT = f'{ADMIN_BASE}/admin/memberships/import/{MEMBER_TYPE_ID}/ajax'

USER_AGENT = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
              'AppleWebKit/537.36 (KHTML, like Gecko) '
              'Chrome/120.0.0.0 Safari/537.36')

# ── Helpers ────────────────────────────────────────────────────────────────────

def log(msg: str):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f'[{ts}] {msg}')


def die(msg: str):
    print(f'\n✗ {msg}')
    sys.exit(1)


# ── Core ───────────────────────────────────────────────────────────────────────

def run_import(xls_path: str, phpsessid: str) -> bool:
    file_name = os.path.basename(xls_path)
    file_size = os.path.getsize(xls_path)

    log(f'Import file : {file_name} ({file_size:,} bytes)')
    log(f'Endpoint    : {IMPORT_ENDPOINT}')
    log(f'Session     : PHPSESSID={phpsessid[:8]}...{phpsessid[-4:]}')

    with open(xls_path, 'rb') as f:
        files = {
            'file': (
                file_name,
                f,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        }
        headers = {
            'Origin':           ADMIN_BASE,
            'Referer':          f'{ADMIN_BASE}/admin/memberships/import/{MEMBER_TYPE_ID}',
            'X-Requested-With': 'XMLHttpRequest',
            'User-Agent':       USER_AGENT,
        }
        cookies = {
            'PHPSESSID': phpsessid
        }

        log('Uploading ...')
        resp = requests.post(
            IMPORT_ENDPOINT,
            headers=headers,
            cookies=cookies,
            files=files,
            timeout=60
        )

    log(f'HTTP status : {resp.status_code} {resp.reason}')

    try:
        data = resp.json()
    except Exception:
        print(f'\nRaw response (non-JSON):\n{resp.text[:800]}')
        die('Could not parse server response as JSON.')

    code     = data.get('code')
    errors   = data.get('data', {}).get('errors', [])
    warnings = data.get('data', {}).get('warnings', [])
    redirect = data.get('redirect', '')

    if 'login' in redirect:
        print('\n✗ GlueUp returned a login redirect — session cookie has expired.')
        print('\n  Get a fresh PHPSESSID cookie from your browser:')
        print('  1. Log in to https://app.glueup.com')
        print('  2. DevTools → Application → Cookies → app.glueup.com → PHPSESSID')
        print('  3. Re-run this script with the new value.')
        return False

    if code == 200 and not errors:
        log('✓ Import accepted by GlueUp!')
        if warnings:
            log(f'  Warnings ({len(warnings)}):')
            for w in warnings:
                print(f'    - {w}')
        print('\n✓ Done. Verify in GlueUp Admin → Members that the records appear.')
        return True

    if errors:
        print(f'\n✗ Import returned errors:')
        for e in errors:
            print(f'  - {e}')
        print(f'\nFull response:\n{json.dumps(data, indent=2)}')
        return False

    print(f'\n? Unexpected response (code={code}):')
    print(json.dumps(data, indent=2))
    return False


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]

    if not args:
        print('Usage: python AdminImportCookie.py <path_to_import_xlsx>')
        print()
        print('HOW TO GET YOUR PHPSESSID COOKIE:')
        print('  1. Log in to https://app.glueup.com in Chrome or Safari')
        print('  2. Open DevTools: Cmd+Option+I')
        print('  3. Application tab → Cookies → https://app.glueup.com')
        print('  4. Copy the Value of the PHPSESSID cookie')
        print()
        print('Environment variable (to avoid being prompted each run):')
        print('  export GLUEUP_PHPSESSID=<your cookie value>')
        sys.exit(1)

    xls_path = args[0]
    if not os.path.exists(xls_path):
        die(f'File not found: {xls_path}')

    phpsessid = os.environ.get('GLUEUP_PHPSESSID')
    if not phpsessid:
        print('Paste your PHPSESSID cookie value below.')
        print('(DevTools → Application → Cookies → app.glueup.com → PHPSESSID)')
        phpsessid = getpass.getpass('PHPSESSID: ').strip()

    if not phpsessid:
        die('No PHPSESSID provided.')

    print()
    success = run_import(xls_path, phpsessid)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()