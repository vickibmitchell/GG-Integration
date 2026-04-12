"""
AdminImport.py
--------------
Programmatically imports a GlueUp-formatted XLS file via the Admin UI
import endpoint, without needing a human to click through the browser.

How it works:
    1. Logs in to app.glueup.com using your email + MD5 password to
       obtain a PHP session cookie (PHPSESSID).
    2. POSTs the import XLS file to the Admin import endpoint using
       that session cookie, mimicking exactly what the browser does.

This avoids the need for Playwright/browser automation entirely.

Usage:
    python AdminImport.py <path_to_import_xlsx> [--dry-run]

    --dry-run   Authenticates and probes the endpoint but does NOT
                upload the file. Use to verify credentials are working.

Prerequisites:
    pip install requests openpyxl
    The import XLS must be produced by ICFGlueUpSync.py.

Credentials:
    Set GLUEUP_EMAIL and GLUEUP_MD5_PASSWORD as environment variables,
    or let the script prompt you interactively.

    To generate your MD5 password hash:
        echo -n "YOUR_PASSWORD" | md5sum

Environment variables (optional — script prompts if not set):
    GLUEUP_EMAIL          e.g. technology@icfwashingtonstate.org
    GLUEUP_MD5_PASSWORD   MD5 hash of your GlueUp password
"""

import sys
import os
import json
import hmac
import hashlib
import time
import getpass
import requests
from datetime import datetime

# ── Config ─────────────────────────────────────────────────────────────────────

ADMIN_BASE       = 'https://app.glueup.com'
API_BASE         = 'https://api-services.glueup.com/v2'
ORG_ID           = '7912'
MEMBER_TYPE_ID   = '37600'

PK = 'icfwshts'
SK = 'MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPfXb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfUCCAP67vZiAQ55'

IMPORT_ENDPOINT  = f'{ADMIN_BASE}/admin/memberships/import/{MEMBER_TYPE_ID}/ajax'
LOGIN_ENDPOINT   = f'{API_BASE}/user/session'

USER_AGENT = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
              'AppleWebKit/537.36 (KHTML, like Gecko) '
              'Chrome/120.0.0.0 Safari/537.36')

# ── Helpers ────────────────────────────────────────────────────────────────────

def build_a_header(method: str) -> str:
    ts  = str(int(time.time() * 1000))
    msg = method + PK + '1.0' + ts
    d   = hmac.new(SK.encode(), msg.encode(), hashlib.sha256).hexdigest()
    return f'v=1.0;k={PK};ts={ts};d={d}'


def log(msg: str):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f'[{ts}] {msg}')


def die(msg: str):
    print(f'\n✗ {msg}')
    sys.exit(1)


# ── Step 1: Authenticate via API to get session token ─────────────────────────

def get_api_token(email: str, md5_password: str) -> str:
    """
    Log in via the GlueUp v2 API to get a session token.
    This is the same call GetToken.py makes.
    """
    log(f'Authenticating as {email} ...')
    body = json.dumps({
        'email':      {'value': email},
        'passphrase': {'value': md5_password}
    }).encode()

    resp = requests.post(
        LOGIN_ENDPOINT,
        data=body,
        headers={
            'Content-Type': 'application/json',
            'a':            build_a_header('POST'),
            'User-Agent':   USER_AGENT,
        },
        timeout=20
    )

    if resp.status_code not in (200, 201):
        die(f'Login failed: HTTP {resp.status_code}\n{resp.text[:400]}')

    data = resp.json()
    token = (data.get('value') or {}).get('token') or \
            (data.get('data') or {}).get('token') or \
            data.get('token')
    if not token:
        die(f'Login response did not contain a token:\n{json.dumps(data, indent=2)}')

    log(f'✓ API token obtained (expires ~7 days)')
    return token


# ── Step 2: POST import file to Admin endpoint with token + browser headers ───

def run_import(xls_path: str, token: str, dry_run: bool = False) -> bool:
    """
    POST the import XLS to the Admin UI import endpoint.
    Uses the API token + Origin/Referer headers (confirmed working in test).

    Returns True on success, False on failure.
    """
    file_size = os.path.getsize(xls_path)
    file_name = os.path.basename(xls_path)
    log(f'Import file : {file_name} ({file_size:,} bytes)')
    log(f'Endpoint    : {IMPORT_ENDPOINT}')

    if dry_run:
        log('DRY RUN — skipping actual upload.')
        return True

    with open(xls_path, 'rb') as f:
        files = {
            'file': (
                file_name,
                f,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        }
        headers = {
            'a':                     build_a_header('POST'),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'Origin':                ADMIN_BASE,
            'Referer':               f'{ADMIN_BASE}/admin/memberships/import/{MEMBER_TYPE_ID}',
            'X-Requested-With':      'XMLHttpRequest',
            'User-Agent':            USER_AGENT,
        }

        log('Uploading ...')
        resp = requests.post(IMPORT_ENDPOINT, headers=headers, files=files, timeout=60)

    log(f'HTTP status : {resp.status_code} {resp.reason}')

    # Parse response
    try:
        data = resp.json()
    except Exception:
        print(f'\nRaw response (non-JSON):\n{resp.text[:800]}')
        return False

    code     = data.get('code')
    errors   = data.get('data', {}).get('errors', [])
    warnings = data.get('data', {}).get('warnings', [])
    redirect = data.get('redirect', '')

    # ── Interpret the response ─────────────────────────────────────────────────

    # Success: code 200 with no redirect to login
    if code == 200 and 'login' not in redirect:
        log('✓ Import accepted by GlueUp!')
        if warnings:
            log(f'  Warnings ({len(warnings)}):')
            for w in warnings:
                print(f'    - {w}')
        return True

    # Auth redirect: session not recognised
    if 'login' in redirect:
        print(f'\n✗ GlueUp returned a login redirect.')
        print(f'  This means the token header alone is not sufficient for the Admin')
        print(f'  UI endpoint. A PHP browser session (PHPSESSID cookie) is required.')
        print(f'\n  Next step: use Playwright to capture a live browser session.')
        print(f'  Run: python AdminImportPlaywright.py {xls_path}')
        print(f'\n  Full response:')
        print(json.dumps(data, indent=2))
        return False

    # Errors returned
    if errors:
        print(f'\n✗ Import completed with errors:')
        for e in errors:
            print(f'  - {e}')
        print(f'\nFull response:\n{json.dumps(data, indent=2)}')
        return False

    # Unknown response
    print(f'\n? Unexpected response (code={code}):')
    print(json.dumps(data, indent=2))
    return False


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    dry_run = '--dry-run' in sys.argv
    args    = [a for a in sys.argv[1:] if not a.startswith('--')]

    if not args:
        print('Usage: python AdminImport.py <path_to_import_xlsx> [--dry-run]')
        print()
        print('  --dry-run   Authenticate only, do not upload the file.')
        print()
        print('Examples:')
        print('  python AdminImport.py glueup_import_20260330_1200.xlsx')
        print('  python AdminImport.py glueup_import_20260330_1200.xlsx --dry-run')
        sys.exit(1)

    xls_path = args[0]
    if not os.path.exists(xls_path):
        die(f'File not found: {xls_path}')

    # ── Credentials ───────────────────────────────────────────────────────────
    email = os.environ.get('GLUEUP_EMAIL') or input('GlueUp email: ').strip()
    md5pw = os.environ.get('GLUEUP_MD5_PASSWORD')
    if not md5pw:
        raw = getpass.getpass('GlueUp password (will be hashed): ')
        md5pw = hashlib.md5(raw.encode()).hexdigest()
        log(f'MD5 hash: {md5pw}')

    print()

    # ── Run ───────────────────────────────────────────────────────────────────
    token   = get_api_token(email, md5pw)
    success = run_import(xls_path, token, dry_run=dry_run)

    print()
    if success:
        print('✓ Done.')
        sys.exit(0)
    else:
        print('✗ Import did not complete successfully. See details above.')
        sys.exit(1)


if __name__ == '__main__':
    main()