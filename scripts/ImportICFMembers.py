"""
ImportICFMembers.py — Import ICF Global Member Export into GlueUp
ICF Washington State Chapter

Reads the ICF Global chapter member export (XLSX, PC sheet) and syncs
each member into GlueUp using the full GET→PUT→DELETE update sequence
described in the Design Specification v3.1.

PROCESSING LOGIC (from spec Section 7.1):
  Outcome 1 — Known member, no changes     → skip (no GlueUp call)
  Outcome 2 — Known member, email promoted → promote then update
  Outcome 3 — In GlueUp, not in log        → GET→PUT→DELETE, add to log
  Outcome 4 — New member, not in GlueUp    → PUT to create, add to log
  Outcome 5 — In log, not in GlueUp        → flag for manual review
  Outcome 6 — Email conflict               → flag for manual review

SAFETY FEATURES:
  - Dry run by default — shows what would happen without making API calls
  - Requires --confirm to execute actual GlueUp writes
  - Processes only N rows with --limit N (useful for Gate 2 testing)
  - Persists a local run log (JSON) that acts as the Python equivalent
    of the Make Data Store — enables change detection across runs
  - Detailed per-row result log saved after every run
  - Validates option codes against field_codes.json (run GetFieldCodes.py first)

Loads the session token from glueup_token.json (run GetToken.py first).
Loads field option codes from field_codes.json (run GetFieldCodes.py first).

Usage:
    python3 ImportICFMembers.py --in "ICF_Export.xlsx"            # dry run
    python3 ImportICFMembers.py --in "ICF_Export.xlsx" --limit 5  # dry run, 5 rows
    python3 ImportICFMembers.py --in "ICF_Export.xlsx" --confirm  # live run
    python3 ImportICFMembers.py --in "ICF_Export.xlsx" --confirm --limit 5
    python3 ImportICFMembers.py --in "ICF_Export.xlsx" --confirm --reset-log
"""

import argparse
import json
import hmac
import hashlib
import time
import os
import sys
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import HTTPError

try:
    from openpyxl import load_workbook
except ImportError:
    print('✗ openpyxl not installed. Run: pip3 install openpyxl --break-system-packages')
    sys.exit(1)

# ── GlueUp API constants ──────────────────────────────────────────────────────
GLUEUP_BASE_URL    = 'https://api-services.glueup.com/v2'
ORG_ID             = '7912'
PUBLIC_KEY         = 'icfwshts'
PRIVATE_KEY        = ('MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPf'
                      'Xb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfU'
                      'CCAP67vZiAQ55')
MEMBERSHIP_TYPE_ID = 37600   # ICF Chapter Member
TOKEN_FILE         = os.path.join(os.path.dirname(__file__), 'glueup_token.json')
RUN_LOG_FILE       = os.path.join(os.path.dirname(__file__), 'import_run_log.json')
FIELD_CODES_FILE   = os.path.join(os.path.dirname(__file__), 'field_codes.json')

SHADOW_DOMAIN = 'members.icfwashingtonstate.org'

# ICF XLS PC sheet column indices (0-based), columns A-U
COL = {
    'member_id':              0,   # A
    'status':                 1,   # B
    'first_name':             2,   # C
    'last_name':              3,   # D
    'role':                   4,   # E
    'email':                  5,   # F
    'phone':                  6,   # G
    'city':                   7,   # H
    'state':                  8,   # I
    'zip':                    9,   # J
    'country':                10,  # K
    'creation_date':          11,  # L — excluded from fingerprint
    'expiration_date':        12,  # M
    'credential':             13,  # N
    'credential_award_date':  14,  # O
    'credential_expire_date': 15,  # P
    'tc_credential':          16,  # Q
    'tc_award_date':          17,  # R
    'tc_expire_date':         18,  # S
    'reinstate_rejoin':       19,  # T
    'auto_renewal':           20,  # U
}

# GlueUp property keys — confirmed from GetFieldCodes.py output
# NOTE: GlueUp truncates keys at ~30 chars. Use the exact keys returned by the API.
PROP = {
    'icfmemberid':                  'icfmemberid',
    'hasemail':                     'hasemail',
    'icfimportdate':                'icfimportdate',
    'icfglobalmembershiptype':      'icfglobalmembershiptype',
    'icfglobalmembershipstartdate': 'icfglobalmembershipstartdate',
    'icfglobalmembershipenddate':   'icfglobalmembershipenddate',
    'icfglobalautorenewal':         'icfglobalautorenewal',
    'icfcredential':                'icfcredential',
    'icfcredentialawarddate':       'icfcredentialawarddate',
    'icfcredentialexpiredate':      'icfcredentialexpiredate',
    'icfteamcoachingcredential':    'icfteamcoachingcredential',
    # Truncated keys as returned by GlueUp API (confirmed via GetFieldCodes.py)
    'icfteamcoachingcredentialaward':  'icfteamcoachingcredentialaward',
    'icfteamcoachingcredentialexpir':  'icfteamcoachingcredentialexpir',
}

# Credential code map — ICF XLS text → GlueUp option code
# Populated from field_codes.json at startup; fallback hardcoded values below
DEFAULT_CREDENTIAL_CODES = {
    'ACC': 'acc', 'PCC': 'pcc', 'MCC': 'mcc', 'CC': 'cc',
    'ACTC': 'actc', 'PCTC': 'pctc', 'MCTC': 'mctc',
    '': 'none',
}
DEFAULT_AUTORENEWAL_CODES = {
    'Yes': 'yes', 'No': 'no', 'yes': 'yes', 'no': 'no', '': 'unspecified',
}


# ── Field codes loader ────────────────────────────────────────────────────────

def load_field_codes() -> dict:
    """Load field_codes.json produced by GetFieldCodes.py."""
    if not os.path.exists(FIELD_CODES_FILE):
        return {}
    with open(FIELD_CODES_FILE) as f:
        return json.load(f)


def build_code_maps(field_codes: dict) -> tuple[dict, dict]:
    """
    Build credential and auto-renewal code maps from field_codes.json.
    Falls back to DEFAULT_* maps if field_codes.json is not available.
    ICF XLS values are uppercase (ACC, PCC) — map to lowercase GlueUp codes.
    """
    cred_map   = dict(DEFAULT_CREDENTIAL_CODES)
    renew_map  = dict(DEFAULT_AUTORENEWAL_CODES)
    fields     = field_codes.get('fields', {})

    # Build credential map from icfcredential options
    cred_field = fields.get('icfcredential', {})
    for opt in cred_field.get('options', []):
        code  = opt['code']
        title = opt.get('title', '').upper()
        cred_map[title] = code
        cred_map[code]  = code   # also accept lowercase

    # Also apply to TC credential from icfteamcoachingcredential options
    tc_field = fields.get('icfteamcoachingcredential', {})
    for opt in tc_field.get('options', []):
        code  = opt['code']
        title = opt.get('title', '').upper()
        cred_map[title] = code
        cred_map[code]  = code

    # Build auto-renewal map from icfglobalautorenewal options
    renew_field = fields.get('icfglobalautorenewal', {})
    for opt in renew_field.get('options', []):
        code  = opt['code']
        title = opt.get('title', '')
        renew_map[title.capitalize()] = code
        renew_map[title.lower()]      = code
        renew_map[code]               = code

    return cred_map, renew_map


# ── Auth helpers ──────────────────────────────────────────────────────────────

def build_a_header(method: str) -> str:
    ts     = str(int(time.time() * 1000))
    msg    = method + PUBLIC_KEY + '1.0' + ts
    digest = hmac.new(PRIVATE_KEY.encode(), msg.encode(), hashlib.sha256).hexdigest()
    return f'v=1.0;k={PUBLIC_KEY};ts={ts};d={digest}'


def load_token() -> str:
    if not os.path.exists(TOKEN_FILE):
        print('✗ No token file found. Run GetToken.py first.')
        sys.exit(1)
    with open(TOKEN_FILE) as f:
        data = json.load(f)
    return data['token']


# ── GlueUp API calls ──────────────────────────────────────────────────────────

def _headers(method: str, token: str) -> dict:
    return {
        'Content-Type':          'application/json',
        'a':                     build_a_header(method),
        'token':                 token,
        'requestOrganizationId': ORG_ID,
        'User-Agent':            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'
    }


def glueup_post(endpoint: str, token: str, body: dict) -> dict:
    req = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        data=json.dumps(body).encode(),
        headers=_headers('POST', token),
        method='POST'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def glueup_get(endpoint: str, token: str) -> dict:
    req = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        headers=_headers('GET', token),
        method='GET'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def glueup_put(endpoint: str, token: str, body: dict) -> dict:
    req = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        data=json.dumps(body).encode(),
        headers=_headers('PUT', token),
        method='PUT'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def glueup_delete(member_id: int, token: str) -> dict:
    req = Request(
        f'{GLUEUP_BASE_URL}/membershipDirectory/member/{member_id}',
        headers=_headers('DELETE', token),
        method='DELETE'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def search_by_icfmemberid(member_id: str, token: str) -> dict | None:
    """Search GlueUp for a member by icfmemberid. Returns first exact match or None."""
    result  = glueup_post('/membershipDirectory/members', token, {
        'filter': [{'projection': 'properties.icfmemberid',
                    'operator': 'eq', 'values': [member_id]}],
        'offset': 0, 'limit': 10
    })
    for rec in result.get('value', []):
        props = (rec.get('individualMember') or {}).get('properties') or {}
        if str(props.get('icfmemberid', '')) == str(member_id):
            return rec
    return None


# ── XLS parsing ───────────────────────────────────────────────────────────────

def cell_str(val) -> str:
    if val is None:
        return ''
    if isinstance(val, datetime):
        return val.strftime('%m/%d/%Y')
    return str(val).strip()


def read_icf_xls(path: str) -> list[dict]:
    """
    Read the PC sheet from the ICF Global export XLSX.
    Row 1: chapter title label (skip)
    Row 2: column headers (skip)
    Row 3+: data
    """
    wb = load_workbook(path, read_only=True)
    sheet_name = next((s for s in wb.sheetnames if s.strip().upper() == 'PC'), None)
    if not sheet_name:
        print(f'✗ Could not find "PC" sheet. Sheets available: {wb.sheetnames}')
        sys.exit(1)

    ws   = wb[sheet_name]
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2:
            continue
        if not any(row):
            continue
        member_id = cell_str(row[COL['member_id']])
        if not member_id:
            continue
        rows.append({k: cell_str(row[v]) for k, v in COL.items()})
    return rows


# ── Business logic helpers ────────────────────────────────────────────────────

def effective_email(row: dict) -> tuple[str, bool]:
    raw = row['email'].strip().lower()
    if raw:
        return raw, True
    return f'id_{row["member_id"]}@{SHADOW_DOMAIN}', False


def build_fingerprint(row: dict, email: str) -> str:
    """Pipe-delimited fingerprint of all tracked fields (excludes member_id and creation_date)."""
    return '|'.join([
        row['status'], row['first_name'], row['last_name'], row['role'],
        email, row['phone'], row['city'], row['state'], row['zip'], row['country'],
        row['expiration_date'], row['credential'], row['credential_award_date'],
        row['credential_expire_date'], row['tc_credential'],
        row['tc_award_date'], row['tc_expire_date'],
        row['reinstate_rejoin'], row['auto_renewal'],
    ])


def changed_fields(old_fp: str, new_fp: str) -> list[str]:
    names = [
        'status', 'first_name', 'last_name', 'role', 'email',
        'phone', 'city', 'state', 'zip', 'country',
        'expiration_date', 'credential', 'credential_award_date',
        'credential_expire_date', 'tc_credential',
        'tc_award_date', 'tc_expire_date', 'reinstate_rejoin', 'auto_renewal',
    ]
    old_parts = old_fp.split('|')
    new_parts = new_fp.split('|')
    return [
        names[i] for i in range(len(names))
        if (old_parts[i] if i < len(old_parts) else '') !=
           (new_parts[i] if i < len(new_parts) else '')
    ]


def today_str() -> str:
    return datetime.utcnow().strftime('%m/%d/%Y')


def to_single_choice(code_map: dict, value: str) -> dict | None:
    """Map a text value to a GlueUp single-choice {code} object. Returns None if unmappable."""
    if value is None:
        value = ''
    code = code_map.get(value.strip())
    if code is None:
        code = code_map.get(value.strip().upper())
    if code is None:
        code = code_map.get('')   # fallback to blank/none code
    if code:
        return {'code': code}
    return None


def build_glueup_payload(row: dict, email: str, has_email: bool,
                         cred_map: dict, renew_map: dict,
                         existing_member_id: int | None = None) -> dict:
    """Build the full GlueUp PUT payload for a member record."""

    props = {
        PROP['icfmemberid']:                  row['member_id'],
        PROP['hasemail']:                     {'code': 'yes' if has_email else 'no'},
        PROP['icfimportdate']:                today_str(),
        PROP['icfglobalmembershiptype']:      {'code': 'individual'},
        PROP['icfglobalmembershipstartdate']: row['creation_date'],
        PROP['icfglobalmembershipenddate']:   row['expiration_date'],
    }

    # Auto renewal — with unspecified fallback
    auto = to_single_choice(renew_map, row['auto_renewal'])
    if auto:
        props[PROP['icfglobalautorenewal']] = auto

    # ICF Credential (single-choice, 'none' if blank)
    cred = to_single_choice(cred_map, row['credential'])
    if cred:
        props[PROP['icfcredential']] = cred
    if row['credential_award_date']:
        props[PROP['icfcredentialawarddate']]  = row['credential_award_date']
    if row['credential_expire_date']:
        props[PROP['icfcredentialexpiredate']] = row['credential_expire_date']

    # Team Coaching Credential (single-choice, 'none' if blank)
    tc = to_single_choice(cred_map, row['tc_credential'])
    if tc:
        props[PROP['icfteamcoachingcredential']] = tc
    # Use truncated keys confirmed by GetFieldCodes.py
    if row['tc_award_date']:
        props[PROP['icfteamcoachingcredentialaward']] = row['tc_award_date']
    if row['tc_expire_date']:
        props[PROP['icfteamcoachingcredentialexpir']] = row['tc_expire_date']

    member_payload = {
        'membershipType': {'id': MEMBERSHIP_TYPE_ID},
        'emailAddress':   {'value': email},
        'givenName':      row['first_name'],
        'familyName':     row['last_name'],
        'properties':     props,
    }

    if row['phone']:
        member_payload['phone'] = {'value': row['phone']}

    if any([row['city'], row['state'], row['zip'], row['country']]):
        addr = {}
        if row['city']:    addr['cityName'] = row['city']
        if row['state']:   addr['province'] = row['state']
        if row['zip']:     addr['zipCode']  = row['zip']
        if row['country']: addr['country']  = {'code': row['country']}
        member_payload['address'] = addr

    if existing_member_id:
        member_payload['id'] = existing_member_id

    return member_payload


# ── Run log ───────────────────────────────────────────────────────────────────

def load_run_log() -> dict:
    if os.path.exists(RUN_LOG_FILE):
        with open(RUN_LOG_FILE) as f:
            return json.load(f)
    return {}


def save_run_log(log: dict) -> None:
    with open(RUN_LOG_FILE, 'w') as f:
        json.dump(log, f, indent=2)


# ── Core import logic ─────────────────────────────────────────────────────────

def process_row(row: dict, run_log: dict, token: str, confirm: bool,
                cred_map: dict, renew_map: dict) -> dict:
    """Process a single XLS row. Returns a result dict."""
    member_id        = row['member_id']
    email, has_email = effective_email(row)
    fingerprint      = build_fingerprint(row, email)
    log_entry        = run_log.get(member_id)
    now              = datetime.utcnow().isoformat() + 'Z'

    result = {
        'member_id':            member_id,
        'name':                 f'{row["first_name"]} {row["last_name"]}',
        'email':                email,
        'has_email':            has_email,
        'fingerprint':          fingerprint,
        'timestamp':            now,
        'outcome':              '',
        'changed':              [],
        'glueup_member_id':     '',
        'glueup_membership_id': '',
        'notes':                '',
    }

    # ── BRANCH B: Known member (in run log) ───────────────────────────────────
    if log_entry:
        old_fp         = log_entry.get('fingerprint', '')
        email_promoted = (not log_entry.get('has_email')) and has_email

        if old_fp == fingerprint and not email_promoted:
            result['outcome']             = 'skipped'
            result['glueup_member_id']    = log_entry.get('glueup_member_id', '')
            result['glueup_membership_id']= log_entry.get('glueup_membership_id', '')
            return result

        result['changed'] = changed_fields(old_fp, fingerprint)
        if email_promoted:
            result['notes'] = 'email_promoted'

        if not confirm:
            result['outcome'] = 'would_update'
            result['notes']  += f' changed={result["changed"]}'
            return result

        old_member_id = log_entry.get('glueup_member_id')
        if not old_member_id:
            result['outcome'] = 'flagged'
            result['notes']   = 'in_log_no_glueup_member_id'
            return result

        try:
            get_resp          = glueup_get(f'/membershipDirectory/member/{old_member_id}', token)
        except HTTPError as e:
            result['outcome'] = 'flagged'
            result['notes']   = f'GET_failed_{e.code}'
            return result

        payload = build_glueup_payload(row, email, has_email, cred_map, renew_map,
                                       int(old_member_id))
        try:
            put_resp          = glueup_put(
                f'/membershipDirectory/member/{old_member_id}', token, payload)
            new_member_id     = put_resp.get('value', {}).get('individualMember', {}).get('id')
            new_membership_id = put_resp.get('value', {}).get('membership', {}).get('id')
        except HTTPError as e:
            result['outcome'] = 'error'
            result['notes']   = f'PUT_failed_{e.code}: {e.read().decode()}'
            return result

        try:
            glueup_delete(old_member_id, token)
        except HTTPError as e:
            result['outcome']             = 'duplicate_pending'
            result['glueup_member_id']    = str(new_member_id)
            result['glueup_membership_id']= str(new_membership_id)
            result['notes']               = f'PUT_ok_DELETE_failed_{e.code}'
            return result

        result['outcome']             = 'updated'
        result['glueup_member_id']    = str(new_member_id)
        result['glueup_membership_id']= str(new_membership_id)
        return result

    # ── BRANCH A: First encounter (not in run log) ────────────────────────────
    if not confirm:
        result['outcome'] = 'would_create_or_update'
        return result

    existing = search_by_icfmemberid(member_id, token)

    if existing:
        # Outcome 3 — in GlueUp but not in log
        existing_member_id = (existing.get('individualMember') or {}).get('id')

        try:
            glueup_get(f'/membershipDirectory/member/{existing_member_id}', token)
        except HTTPError as e:
            result['outcome'] = 'error'
            result['notes']   = f'GET_failed_{e.code}'
            return result

        payload = build_glueup_payload(row, email, has_email, cred_map, renew_map,
                                       existing_member_id)
        try:
            put_resp          = glueup_put(
                f'/membershipDirectory/member/{existing_member_id}', token, payload)
            new_member_id     = put_resp.get('value', {}).get('individualMember', {}).get('id')
            new_membership_id = put_resp.get('value', {}).get('membership', {}).get('id')
        except HTTPError as e:
            result['outcome'] = 'error'
            result['notes']   = f'PUT_failed_{e.code}: {e.read().decode()}'
            return result

        try:
            glueup_delete(existing_member_id, token)
        except HTTPError as e:
            result['outcome']             = 'duplicate_pending'
            result['glueup_member_id']    = str(new_member_id)
            result['glueup_membership_id']= str(new_membership_id)
            result['notes']               = f'PUT_ok_DELETE_failed_{e.code}'
            return result

        result['outcome']             = 'updated'
        result['glueup_member_id']    = str(new_member_id)
        result['glueup_membership_id']= str(new_membership_id)
        result['notes']               = 'first_encounter_found_in_glueup'
        return result

    else:
        # Outcome 4 — new member, check for email conflict first
        email_search = glueup_post('/membershipDirectory/members', token, {
            'search': {'fields': ['emailAddress'], 'value': email, 'fullText': True},
            'offset': 0, 'limit': 5
        })
        conflict = any(
            (r.get('individualMember') or {}).get('emailAddress', {})
            .get('value', '').lower() == email
            for r in email_search.get('value', [])
        )
        if conflict:
            result['outcome'] = 'flagged'
            result['notes']   = 'email_conflict'
            return result

        payload = build_glueup_payload(row, email, has_email, cred_map, renew_map)
        try:
            put_resp          = glueup_put('/membershipDirectory/member/0', token, payload)
            new_member_id     = put_resp.get('value', {}).get('individualMember', {}).get('id')
            new_membership_id = put_resp.get('value', {}).get('membership', {}).get('id')
        except HTTPError as e:
            result['outcome'] = 'error'
            result['notes']   = f'PUT_failed_{e.code}: {e.read().decode()}'
            return result

        result['outcome']             = 'created'
        result['glueup_member_id']    = str(new_member_id)
        result['glueup_membership_id']= str(new_membership_id)
        return result


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Import ICF Global member export XLSX into GlueUp.'
    )
    parser.add_argument('--in',        dest='infile',    required=True,
                        help='ICF Global export XLSX file path')
    parser.add_argument('--confirm',   action='store_true',
                        help='Execute actual GlueUp writes (default is dry run)')
    parser.add_argument('--limit',     type=int, default=None,
                        help='Process only the first N rows (useful for testing)')
    parser.add_argument('--reset-log', action='store_true',
                        help='Clear the run log before processing (full re-import)')
    args = parser.parse_args()

    if not os.path.exists(args.infile):
        print(f'✗ Input file not found: {args.infile}')
        sys.exit(1)

    # Load field codes
    field_codes = load_field_codes()
    if field_codes:
        print(f'Field codes loaded from field_codes.json '
              f'(generated {field_codes.get("generated", "unknown")})')
    else:
        print('⚠  field_codes.json not found — using hardcoded defaults. '
              'Run GetFieldCodes.py to generate it.')
    cred_map, renew_map = build_code_maps(field_codes)

    # Load XLS
    print(f'Reading ICF Global export: {args.infile}')
    rows = read_icf_xls(args.infile)
    print(f'Loaded {len(rows)} member rows from PC sheet.')

    if args.limit:
        rows = rows[:args.limit]
        print(f'Limited to first {args.limit} rows.')

    if args.reset_log:
        print('⚠  Run log cleared — all members will be treated as new.')
        run_log = {}
    else:
        run_log = load_run_log()
        print(f'Run log loaded: {len(run_log)} previously processed members.\n')

    if not args.confirm:
        print('=' * 72)
        print('DRY RUN — no GlueUp API calls will be made.')
        print('Pass --confirm to execute. Pass --limit N to test a small batch.')
        print('=' * 72 + '\n')

    token = load_token() if args.confirm else None

    counts  = {}
    results = []

    for i, row in enumerate(rows, start=1):
        member_id = row['member_id']
        name      = f'{row["first_name"]} {row["last_name"]}'
        print(f'  [{i:4d}/{len(rows)}] {member_id}  {name:<35}', end=' ', flush=True)

        result  = process_row(row, run_log, token, args.confirm, cred_map, renew_map)
        outcome = result['outcome']
        counts[outcome] = counts.get(outcome, 0) + 1
        results.append(result)

        # Update run log for successful live runs
        if args.confirm and outcome in ('created', 'updated'):
            run_log[member_id] = {
                'member_id':            member_id,
                'fingerprint':          result['fingerprint'],
                'has_email':            result['has_email'],
                'glueup_member_id':     result['glueup_member_id'],
                'glueup_membership_id': result['glueup_membership_id'],
                'last_processed':       result['timestamp'],
                'glueup_result':        outcome,
            }
        elif args.confirm and outcome == 'skipped' and member_id in run_log:
            run_log[member_id]['last_processed'] = result['timestamp']

        notes   = f'  [{result["notes"]}]' if result['notes'] else ''
        changed = f'  changed={result["changed"]}' if result['changed'] else ''
        print(f'{outcome}{notes}{changed}')

    # Save run log
    if args.confirm:
        save_run_log(run_log)
        print(f'\nRun log saved to: {RUN_LOG_FILE}')

    # Save detailed results
    date_str     = datetime.utcnow().strftime('%Y%m%d_%H%M')
    results_file = os.path.join(os.path.dirname(os.path.abspath(args.infile)),
                                f'import_results_{date_str}.json')
    with open(results_file, 'w') as f:
        json.dump(results, f, indent=2)

    # Summary
    print('\n' + '=' * 72)
    print(f'{"LIVE RUN" if args.confirm else "DRY RUN"} COMPLETE — {len(rows)} rows processed')
    print('=' * 72)
    for outcome, count in sorted(counts.items()):
        if count > 0:
            print(f'  {outcome:<28} {count}')
    print(f'\nDetailed results: {results_file}')

    if counts.get('flagged', 0):
        print(f'\n⚠  {counts["flagged"]} record(s) flagged — check results file '
              f'(email_conflict, missing IDs, etc.)')
    if counts.get('duplicate_pending', 0):
        print(f'\n⚠  {counts["duplicate_pending"]} duplicate_pending — '
              f'PUT succeeded but DELETE failed. Run DeleteMultiMembers.py to clean up.')
    if counts.get('error', 0):
        print(f'\n✗  {counts["error"]} error(s) — check results file for HTTP details.')


if __name__ == '__main__':
    main()