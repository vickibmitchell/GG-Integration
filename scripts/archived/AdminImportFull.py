"""
AdminImportFull.py
------------------
Automates the complete GlueUp Admin UI import wizard programmatically,
replicating all steps a human performs in the browser.

The 4-step import sequence:
    Step 1 — Upload file      POST https://icfwashingtonstate.glueup.com/upload/files
                               Returns internal fileName (UUID) and processId
    Step 2 — Map fields &     POST .../mapping/individual/{processId}/ajax
             assign owner      action=memberFieldsSubmit
    Step 3 — Import           POST .../mapping/individual/{processId}/ajax
                               action=continueActiveMembersSubmit
    Step 4 — Activate         POST .../mapping/individual/{processId}/ajax
                               action=notificationSettingsSubmit

Prerequisites:
    pip install requests
    A valid PHPSESSID cookie from a logged-in browser session.

Usage:
    python AdminImportFull.py <path_to_import_xlsx>

Environment variables (optional — script prompts if not set):
    GLUEUP_PHPSESSID   PHP session cookie from app.glueup.com

HOW TO GET YOUR PHPSESSID COOKIE:
    1. Log in to https://icfwashingtonstate.glueup.com in Chrome.
    2. Open DevTools: Cmd+Option+I
    3. Application tab → Cookies → https://icfwashingtonstate.glueup.com
    4. Copy the Value of the PHPSESSID cookie.
    5. export GLUEUP_PHPSESSID=<value>
"""

import sys
import os
import json
import getpass
import urllib.parse
import requests
from datetime import datetime

# ── Config ─────────────────────────────────────────────────────────────────────

BASE_URL         = 'https://icfwashingtonstate.glueup.com'
ORG_ID           = '7912'
MEMBER_TYPE_ID   = '37600'
ASSIGNEE_ID      = '2073592'                        # Vicki Mitchell (Admin)
ASSIGNEE_UUID    = 'c32a5615-57f9-4068-9f7f-2a2c26969fd8'
ASSIGNEE_GIVEN   = 'Vicki (Admin)'
ASSIGNEE_FAMILY  = 'Mitchell'

# Column → GlueUp field mapping (from captured memberFieldsSubmit payload)
COLUMN_MAPPING = {
    'A':  'startDate',
    'B':  'endDate',
    'C':  'currencyCode',
    'D':  'firstName',
    'E':  'lastName',
    'F':  'email',
    'G':  'address.zipCode',
    'H':  'address.address',
    'I':  'address.cityName',
    'J':  'properties.volunteerrole',
    'K':  'properties.coachindustry',
    'L':  'properties.coachspecialty',
    'M':  'properties.findable',
    'N':  'properties.directorylistingtext',
    'O':  'properties.icfcredential',
    'P':  'properties.icfcredentialawarddate',
    'Q':  'properties.icfcredentialexpiredate',
    'R':  'properties.icfglobalautorenewal',
    'S':  'properties.icfmemberid',
    'T':  'properties.icfglobalmembershipenddate',
    'U':  'properties.icfglobalmembershipstartdate',
    'V':  'properties.icfglobalmembershiptype',
    'W':  'properties.icfteamcoachingcredential',
    'X':  'properties.icfteamcoachingcredentialaward',
    'Y':  'properties.icfteamcoachingcredentialexpir',
    'Z':  'properties.hasemail',
    'AA': 'properties.icfimportdate',
}

USER_AGENT = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
              'AppleWebKit/537.36 (KHTML, like Gecko) '
              'Chrome/146.0.0.0 Safari/537.36')

# ── Helpers ────────────────────────────────────────────────────────────────────

def log(msg: str):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f'[{ts}] {msg}')


def die(msg: str):
    print(f'\n✗ {msg}')
    sys.exit(1)


def make_headers(referer: str, phpsessid: str, token_cookie: str = None) -> dict:
    h = {
        'Accept':            'application/json, text/javascript, */*; q=0.01',
        'Accept-Language':   'en-US,en;q=0.9',
        'Origin':            BASE_URL,
        'Referer':           referer,
        'X-Requested-With':  'XMLHttpRequest',
        'User-Agent':        USER_AGENT,
    }
    return h


def make_cookies(phpsessid: str, token_cookie: str = None) -> dict:
    c = {'PHPSESSID': phpsessid}
    if token_cookie:
        c['TOKEN'] = token_cookie
    # AWS load balancer cookies — required by the upload endpoint
    for name in ('AWSALB', 'AWSALBTG', 'AWSALBCORS', 'AWSALBTGCORS'):
        val = os.environ.get(f'GLUEUP_{name}', '').strip()
        if val:
            c[name] = val
    return c


def ajax_post(session: requests.Session, endpoint: str, referer: str,
              action: str, data: dict, phpsessid: str,
              token_cookie: str = None, current_path: str = '') -> dict:
    """POST an ajax action to a GlueUp admin endpoint."""

    # Build form-encoded payload matching exactly what the browser sends
    payload = {
        'action':      action,
        'data':        json.dumps(data, separators=(',', ':')),
        'token':       token_cookie or '',
        'orgID':       ORG_ID,
        'currentPath': current_path,
    }

    resp = session.post(
        endpoint,
        headers={
            **make_headers(referer, phpsessid, token_cookie),
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        },
        cookies=make_cookies(phpsessid, token_cookie),
        data=payload,
        timeout=30
    )

    if resp.status_code not in (200, 204):
        die(f'HTTP {resp.status_code} on {action}: {resp.text[:400]}')

    try:
        return resp.json()
    except Exception:
        return {'raw': resp.text}


def check_redirect(result: dict, step: str):
    """Fail fast if GlueUp returned a login redirect."""
    redirect = result.get('redirect', '')
    if 'login' in redirect:
        print(f'\n✗ Login redirect on {step}.')
        print('  Your PHPSESSID cookie has expired.')
        print('  Get a fresh one from DevTools and re-run.')
        sys.exit(1)


# ── Import Steps ───────────────────────────────────────────────────────────────

def step1_upload_file(session: requests.Session, xls_path: str,
                      phpsessid: str, token_cookie: str) -> tuple[str, str]:
    """
    Upload the XLS file to GlueUp's file store.
    Returns (file_name_uuid, process_id).
    """
    log('Step 1 — Uploading file ...')

    file_name = os.path.basename(xls_path)
    upload_url = f'{BASE_URL}/upload/files'
    referer    = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/'

    with open(xls_path, 'rb') as f:
        resp = session.post(
            upload_url,
            headers=make_headers(referer, phpsessid, token_cookie),
            cookies=make_cookies(phpsessid, token_cookie),
            files={'file': (file_name, f,
                            'application/vnd.openxmlformats-officedocument'
                            '.spreadsheetml.sheet')},
            timeout=60
        )

    if resp.status_code not in (200, 201):
        die(f'File upload failed: HTTP {resp.status_code}\n{resp.text[:400]}')

    try:
        data = resp.json()
    except Exception:
        die(f'File upload response not JSON:\n{resp.text[:400]}')

    check_redirect(data, 'Step 1 (file upload)')

    # Extract fileName (UUID) and processId from response
    # GlueUp typically returns these in data.value or data.data
    value = data.get('data', {}) or data.get('value', {}) or data
    if isinstance(value, dict):
        file_name_uuid = value.get('fileName') or value.get('filename') or value.get('id')
        process_id     = value.get('processId') or value.get('process_id')
    else:
        file_name_uuid = None
        process_id     = None

    # Fallback: scan entire response for these keys
    if not file_name_uuid or not process_id:
        raw = json.dumps(data)
        log(f'  Full upload response: {raw[:600]}')
        die('Could not extract fileName/processId from upload response.\n'
            'Check the response above and update step1_upload_file() accordingly.')

    log(f'  ✓ File UUID  : {file_name_uuid}')
    log(f'  ✓ Process ID : {process_id}')
    return file_name_uuid, process_id


def step2_map_fields(session: requests.Session, process_id: str,
                     file_name_uuid: str, xls_filename: str,
                     phpsessid: str, token_cookie: str):
    """
    Submit column mapping, owner assignment, and field definitions.
    action=memberFieldsSubmit
    """
    log('Step 2 — Mapping fields and assigning owner ...')

    endpoint     = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/mapping/individual/{process_id}/ajax'
    referer      = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/mapping/individual/{process_id}'
    current_path = f'/admin/memberships/import/{MEMBER_TYPE_ID}/mapping/individual/{process_id}'

    data = {
        'currentSheet':       'memberSheet',
        'hasHeader':          'true',
        'type':               'IndividualMembership',
        'processId':          process_id,
        'companyFields':      [],
        'membershipTypeType': 'IndividualMembership',
        'membershipTypeId':   MEMBER_TYPE_ID,
        'documentName':       xls_filename,
        'fileName':           file_name_uuid,
        'assignee':           ASSIGNEE_ID,
        **COLUMN_MAPPING
    }

    result = ajax_post(session, endpoint, referer, 'memberFieldsSubmit',
                       data, phpsessid, token_cookie, current_path)
    check_redirect(result, 'Step 2 (memberFieldsSubmit)')

    errors = result.get('data', {}).get('errors', [])
    if errors:
        die(f'Step 2 errors: {errors}')

    log('  ✓ Fields mapped and owner assigned.')
    return result


def step3_import(session: requests.Session, process_id: str,
                 phpsessid: str, token_cookie: str):
    """
    Trigger the actual import processing.
    action=continueActiveMembersSubmit
    """
    log('Step 3 — Triggering import ...')

    endpoint     = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/mapping/individual/{process_id}/ajax'
    referer      = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/report/individual/{process_id}'
    current_path = f'/admin/memberships/import/{MEMBER_TYPE_ID}/report/individual/{process_id}'

    data = {
        'id':        MEMBER_TYPE_ID,
        'type':      'IndividualMembership',
        'processId': process_id,
    }

    result = ajax_post(session, endpoint, referer, 'continueActiveMembersSubmit',
                       data, phpsessid, token_cookie, current_path)
    check_redirect(result, 'Step 3 (continueActiveMembersSubmit)')

    errors = result.get('data', {}).get('errors', [])
    if errors:
        die(f'Step 3 errors: {errors}')

    log('  ✓ Import triggered.')
    return result


def step4_notification(session: requests.Session, process_id: str,
                   phpsessid: str, token_cookie: str):
    """
    Submit notification settings.
    action=notificationSettingsSubmit
    """
    log('Step 4 -- Submitting notification settings ...')

    endpoint     = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/mapping/individual/{process_id}/ajax'
    referer      = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/active/individual/{process_id}'
    current_path = f'/admin/memberships/import/{MEMBER_TYPE_ID}/active/individual/{process_id}'

    data = {
        'type':       'IndividualMembership',
        'processId':  process_id,
        'id':         MEMBER_TYPE_ID,
        'graceEmail': False,
        'renewEmail': False,
    }

    result = ajax_post(session, endpoint, referer, 'notificationSettingsSubmit',
                       data, phpsessid, token_cookie, current_path)
    check_redirect(result, 'Step 4 (notificationSettingsSubmit)')

    errors = result.get('data', {}).get('errors', [])
    if errors:
        die(f'Step 4 errors: {errors}')

    log('  + Notification settings submitted.')
    return result


# ── Token extraction ───────────────────────────────────────────────────────────

def get_token_cookie(session: requests.Session, phpsessid: str) -> str:
    """
    The Admin UI stores a TOKEN cookie alongside PHPSESSID.
    If the user has provided PHPSESSID from a live browser session,
    the TOKEN cookie will also be present. We can extract it by making
    any request and reading Set-Cookie headers, or the user can provide
    it directly.

    For now, read from environment or prompt. The TOKEN cookie was
    visible in all captured cURL requests alongside PHPSESSID.
    """
    token = os.environ.get('GLUEUP_TOKEN_COOKIE', '').strip()
    if token:
        log(f'TOKEN cookie loaded from environment.')
    for name in ('AWSALB', 'AWSALBTG', 'AWSALBCORS', 'AWSALBTGCORS'):
        if os.environ.get(f'GLUEUP_{name}'):
            log(f'  {name} cookie loaded.')
        else:
            log(f'  WARNING: GLUEUP_{name} not set -- upload may fail.')
        return token

    log('Note: TOKEN cookie not set. Attempting without it.')
    log('If Step 2 fails, also export GLUEUP_TOKEN_COOKIE from DevTools.')
    return ''


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]

    if not args:
        print('Usage: python AdminImportFull.py <path_to_import_xlsx>')
        print()
        print('Environment variables:')
        print('  GLUEUP_PHPSESSID      Required. PHP session cookie.')
        print('  GLUEUP_TOKEN_COOKIE   Optional. TOKEN cookie (also from DevTools).')
        print()
        print('HOW TO GET COOKIES (both from same DevTools session):')
        print('  1. Log in to https://icfwashingtonstate.glueup.com in Chrome')
        print('  2. DevTools → Application → Cookies → icfwashingtonstate.glueup.com')
        print('  3. Copy PHPSESSID value → export GLUEUP_PHPSESSID=<value>')
        print('  4. Copy TOKEN value    → export GLUEUP_TOKEN_COOKIE=<value>')
        sys.exit(1)

    xls_path = args[0]
    if not os.path.exists(xls_path):
        die(f'File not found: {xls_path}')

    # ── Credentials ───────────────────────────────────────────────────────────
    phpsessid = os.environ.get('GLUEUP_PHPSESSID', '').strip()
    if not phpsessid:
        print('Paste your PHPSESSID cookie value.')
        print('(DevTools → Application → Cookies → icfwashingtonstate.glueup.com → PHPSESSID)')
        phpsessid = getpass.getpass('PHPSESSID: ').strip()
    if not phpsessid:
        die('No PHPSESSID provided.')

    file_name    = os.path.basename(xls_path)
    file_size    = os.path.getsize(xls_path)

    print()
    log(f'Import file : {file_name} ({file_size:,} bytes)')
    log(f'Base URL    : {BASE_URL}')
    log(f'Assignee    : {ASSIGNEE_GIVEN} {ASSIGNEE_FAMILY} (id={ASSIGNEE_ID})')
    print()

    # Use a session to persist cookies across requests
    session = requests.Session()
    token_cookie = get_token_cookie(session, phpsessid)

    # ── Run the 4-step wizard ─────────────────────────────────────────────────
    file_name_uuid, process_id = step1_upload_file(
        session, xls_path, phpsessid, token_cookie)

    step2_map_fields(
        session, process_id, file_name_uuid, file_name, phpsessid, token_cookie)

    step3_import(
        session, process_id, phpsessid, token_cookie)

    step4_notification(
        session, process_id, phpsessid, token_cookie)

    step5_confirm_activate(
        session, process_id, phpsessid, token_cookie)

    print()
    print('✓ Import complete!')
    print(f'  Process ID : {process_id}')
    print(f'  Verify in GlueUp Admin → Members that records were created.')
    print()
    print('  If records do not appear within a few minutes, check:')
    print(f'  https://icfwashingtonstate.glueup.com/admin/memberships/import/{MEMBER_TYPE_ID}/active/individual/{process_id}')


if __name__ == '__main__':
    main()


if __name__ == '__main__':
    main()


def step5_confirm_activate(session, process_id, phpsessid, token_cookie):
    """Final confirmation -- activates the imported members. action=ImportActivateConfirm"""
    log('Step 5 -- Confirming activation ...')

    endpoint     = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/mapping/individual/{process_id}/ajax'
    referer      = f'{BASE_URL}/admin/memberships/import/{MEMBER_TYPE_ID}/active/individual/{process_id}'
    current_path = f'/admin/memberships/import/{MEMBER_TYPE_ID}/active/individual/{process_id}'

    data = {
        'id':        None,
        'processId': process_id,
        'type':      'IndividualMembership',
        'membershipImportSettings': {
            'membershipType':           {'id': MEMBER_TYPE_ID},
            'sendUpcomingRenewalEmail': False,
            'sendExpiringEmail':        False,
        }
    }

    result = ajax_post(session, endpoint, referer, 'ImportActivateConfirm',
                       data, phpsessid, token_cookie, current_path)
    check_redirect(result, 'Step 5 (ImportActivateConfirm)')

    errors = result.get('data', {}).get('errors', [])
    if errors:
        die(f'Step 5 errors: {errors}')

    log('  + Activation confirmed.')
    return result