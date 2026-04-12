"""
AdminImportPlaywright.py
------------------------
Imports a GlueUp-formatted XLS file via the Admin UI using a real
browser session (Playwright), bypassing the PHP session requirement
that blocks the token-only approach in AdminImport.py.

How it works:
    1. Launches a Chromium browser (headless by default).
    2. Navigates to the GlueUp Admin login page and logs in.
    3. Once the PHP session is established, posts the import file
       directly via the browser's fetch() — no separate HTTP call needed.
    4. Reports success or failure from the server response.

Usage:
    python AdminImportPlaywright.py <path_to_import_xlsx> [--headed]

    --headed    Show the browser window (useful for debugging).

Prerequisites:
    pip install playwright
    playwright install chromium

Credentials:
    Set GLUEUP_EMAIL and GLUEUP_PASSWORD as environment variables,
    or let the script prompt you interactively.

    Note: This script uses your PLAIN password (not MD5) because it
    types it into the browser login form, exactly as a human would.

Environment variables (optional — script prompts if not set):
    GLUEUP_EMAIL      e.g. technology@icfwashingtonstate.org
    GLUEUP_PASSWORD   your GlueUp plain-text password
"""

import sys
import os
import json
import getpass
import base64
from datetime import datetime

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print('✗ Playwright not installed.')
    print('  Run: pip install playwright && playwright install chromium')
    sys.exit(1)

# ── Config ─────────────────────────────────────────────────────────────────────

ADMIN_BASE      = 'https://app.glueup.com'
ORG_ID          = '7912'
MEMBER_TYPE_ID  = '37600'
LOGIN_URL       = f'{ADMIN_BASE}/account/login'
IMPORT_ENDPOINT = f'{ADMIN_BASE}/admin/memberships/import/{MEMBER_TYPE_ID}/ajax'

# ── Helpers ────────────────────────────────────────────────────────────────────

def log(msg: str):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f'[{ts}] {msg}')


def die(msg: str):
    print(f'\n✗ {msg}')
    sys.exit(1)


# ── Core ───────────────────────────────────────────────────────────────────────

def run_import(xls_path: str, email: str, password: str, headed: bool = False):
    file_name = os.path.basename(xls_path)
    file_size = os.path.getsize(xls_path)

    # Read file bytes for injection into browser
    with open(xls_path, 'rb') as f:
        file_bytes = f.read()
    file_b64 = base64.b64encode(file_bytes).decode()

    log(f'Import file : {file_name} ({file_size:,} bytes)')
    log(f'Headed mode : {headed}')

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not headed)
        context = browser.new_context(
            user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                       'Chrome/120.0.0.0 Safari/537.36'
        )
        page = context.new_page()

        # ── Step 1: Log in ─────────────────────────────────────────────────────
        log(f'Navigating to {LOGIN_URL} ...')
        page.goto(LOGIN_URL, wait_until='domcontentloaded', timeout=30000)

        # Check we actually got the login page
        if 'login' not in page.url and 'account' not in page.url:
            log(f'Warning: unexpected URL after navigation: {page.url}')

        log('Filling login form ...')
        try:
            # Try common selectors for email/password fields
            page.locator('input[type="email"], input[name="email"], input[name="username"], #email').first.fill(email)
            page.locator('input[type="password"], input[name="password"], #password').first.fill(password)
        except Exception as e:
            # Dump page for debugging if headed=False
            if not headed:
                page.screenshot(path='login_page_debug.png')
                log('Screenshot saved to login_page_debug.png for debugging.')
            die(f'Could not find login form fields: {e}\nTry --headed to see the browser.')

        log('Submitting login ...')
        page.locator('input[type="password"], input[name="password"], #password').first.press('Enter')

        # Wait for navigation away from login page
        try:
            page.wait_for_url(lambda url: 'login' not in url, timeout=15000)
            log(f'✓ Logged in. Current URL: {page.url}')
        except PWTimeout:
            if not headed:
                page.screenshot(path='login_failed_debug.png')
                log('Screenshot saved to login_failed_debug.png')
            die('Login timed out — still on login page. Check credentials.\nTry --headed to see what happened.')

        # ── Step 2: POST import file via browser fetch() ───────────────────────
        # We use JavaScript fetch() inside the browser so the PHP session
        # cookies are automatically included — no need to extract them.
        log(f'Posting import file to {IMPORT_ENDPOINT} ...')

        js_result = page.evaluate(f"""
            async () => {{
                // Reconstruct the file from base64
                const b64 = "{file_b64}";
                const binary = atob(b64);
                const bytes = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) {{
                    bytes[i] = binary.charCodeAt(i);
                }}
                const blob = new Blob([bytes], {{
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                }});

                const form = new FormData();
                form.append('file', blob, '{file_name}');

                const resp = await fetch('{IMPORT_ENDPOINT}', {{
                    method: 'POST',
                    headers: {{
                        'X-Requested-With': 'XMLHttpRequest',
                        'Origin': '{ADMIN_BASE}',
                    }},
                    body: form,
                    credentials: 'include'
                }});

                const text = await resp.text();
                return {{
                    status: resp.status,
                    statusText: resp.statusText,
                    body: text
                }};
            }}
        """)

        browser.close()

    # ── Step 3: Interpret response ─────────────────────────────────────────────
    status     = js_result.get('status')
    status_txt = js_result.get('statusText')
    body_raw   = js_result.get('body', '')

    log(f'HTTP status : {status} {status_txt}')

    try:
        data = json.loads(body_raw)
    except Exception:
        print(f'\nRaw response (non-JSON):\n{body_raw[:800]}')
        die('Could not parse server response as JSON.')

    code     = data.get('code')
    errors   = data.get('data', {}).get('errors', [])
    warnings = data.get('data', {}).get('warnings', [])
    redirect = data.get('redirect', '')

    if 'login' in redirect:
        print(f'\n✗ Still getting a login redirect after browser login.')
        print(f'  The PHP session was not established correctly.')
        print(f'  Try running with --headed to watch what happens:')
        print(f'    python AdminImportPlaywright.py {xls_path} --headed')
        print(f'\nFull response:\n{json.dumps(data, indent=2)}')
        return False

    if code == 200 and not errors:
        log('✓ Import accepted by GlueUp!')
        if warnings:
            log(f'  Warnings ({len(warnings)}):')
            for w in warnings:
                print(f'    - {w}')
        print(f'\n✓ Done. Check GlueUp Admin to confirm records were created.')
        return True

    if errors:
        print(f'\n✗ Import completed with errors:')
        for e in errors:
            print(f'  - {e}')
        print(f'\nFull response:\n{json.dumps(data, indent=2)}')
        return False

    print(f'\n? Unexpected response (code={code}):')
    print(json.dumps(data, indent=2))
    return False


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    headed = '--headed' in sys.argv
    args   = [a for a in sys.argv[1:] if not a.startswith('--')]

    if not args:
        print('Usage: python AdminImportPlaywright.py <path_to_import_xlsx> [--headed]')
        print()
        print('  --headed   Show the browser window (useful for debugging).')
        print()
        print('Examples:')
        print('  python AdminImportPlaywright.py glueup_import_20260330_1200.xlsx')
        print('  python AdminImportPlaywright.py glueup_import_20260330_1200.xlsx --headed')
        sys.exit(1)

    xls_path = args[0]
    if not os.path.exists(xls_path):
        die(f'File not found: {xls_path}')

    # ── Credentials ───────────────────────────────────────────────────────────
    email    = os.environ.get('GLUEUP_EMAIL') or input('GlueUp email: ').strip()
    password = os.environ.get('GLUEUP_PASSWORD') or getpass.getpass('GlueUp password: ')
    print()

    success = run_import(xls_path, email, password, headed=headed)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()