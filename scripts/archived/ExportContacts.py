"""
ExportContacts.py — Export GlueUp Contacts Without ICF Chapter Membership
ICF Washington State Chapter

Fetches all members across ALL membership types from the GlueUp directory,
identifies contacts who do NOT have an ICF Chapter Member membership
(membershipType.id = 37600), and exports them to a CSV file.

A contact is excluded from the export if ANY of their membership records
is of type 37600. If they have other membership types but not 37600,
they are included.

Loads the session token from glueup_token.json (run GetToken.py first).

Output:
    contacts_without_icf_membership.csv

Usage:
    python3 ExportContacts.py
    python3 ExportContacts.py --out my_filename.csv   # custom output filename
"""

import argparse
import csv
import json
import hmac
import hashlib
import time
import os
import sys
from collections import defaultdict
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import HTTPError

# ── GlueUp API constants ──────────────────────────────────────────────────────
GLUEUP_BASE_URL    = 'https://api-services.glueup.com/v2'
ORG_ID             = '7912'
PUBLIC_KEY         = 'icfwshts'
PRIVATE_KEY        = ('MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPf'
                      'Xb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfU'
                      'CCAP67vZiAQ55')
ICF_MEMBERSHIP_TYPE = 37600   # ICF Chapter Member
TOKEN_FILE          = os.path.join(os.path.dirname(__file__), 'glueup_token.json')
DEFAULT_OUT         = os.path.join(os.path.dirname(__file__),
                                   'contacts_without_icf_membership.csv')

# CSV columns — in output order
CSV_FIELDS = [
    'email',
    'givenName',
    'familyName',
    'companyName',
    'positionTitle',
    'phone',
    'streetAddress',
    'cityName',
    'zipCode',
    'membershipTypeId',
    'membershipTypeName',
    'membershipId',
    'individualMemberId',
    'membershipStartDate',
    'membershipEndDate',
    'icfmemberid',
    'icfimportdate',
    'icfcredential',
    'icfcredentialexpiredate',
    'icfcredentialawarddate',
    'icfglobalmembershipstartdate',
    'icfglobalmembershipenddate',
    'icfglobalautorenewal',
    'icfteamcoachingcredential',
    'hasemail',
    'volunteerrole',
]


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


def glueup_post(endpoint: str, token: str, body: dict) -> dict:
    a_header = build_a_header('POST')
    encoded  = json.dumps(body).encode()
    req = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        data=encoded,
        headers={
            'Content-Type':          'application/json',
            'a':                     a_header,
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'
        },
        method='POST'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def fetch_all_members(token: str) -> list:
    """
    Page through /membershipDirectory/members with no membership type filter,
    returning all records across all membership types.
    """
    all_records = []
    page_size   = 50
    offset      = 0

    print('Fetching all members across all membership types...', end='', flush=True)

    while True:
        body = {
            'projection': [],
            'offset':     offset,
            'limit':      page_size,
            'order':      {'familyName': 'asc'}
        }
        result  = glueup_post('/membershipDirectory/members', token, body)
        records = result.get('value', [])

        if not records:
            break

        all_records.extend(records)
        print('.', end='', flush=True)

        if len(records) < page_size:
            break

        offset += page_size

    print(f' done. ({len(all_records)} total records fetched)\n')
    return all_records


def ms_to_date(ms) -> str:
    """Convert GlueUp millisecond timestamp to YYYY-MM-DD string."""
    if not ms:
        return ''
    try:
        return datetime.utcfromtimestamp(int(ms) / 1000).strftime('%Y-%m-%d')
    except Exception:
        return str(ms)


def extract_prop(props: dict, key: str) -> str:
    """
    Safely extract a property value from the GlueUp properties dict.
    Handles text, date strings, and single-choice objects {code, title}.
    Multi-choice (list) values are joined as pipe-separated codes.
    """
    val = props.get(key)
    if val is None:
        return ''
    if isinstance(val, str):
        return val
    if isinstance(val, dict):
        # Single-choice field: {code, title: {en: ...}}
        return val.get('code', '')
    if isinstance(val, list):
        # Multi-choice field: [{code, title}, ...]
        return '|'.join(item.get('code', '') for item in val if isinstance(item, dict))
    return str(val)


def build_contact_rows(all_records: list) -> tuple[list, set]:
    """
    Group all records by email. Identify which emails have at least one
    ICF Chapter Member (37600) membership. Build a row dict for each
    non-37600 record, deduplicated by email (keep one row per email,
    preferring the record with the most recent membershipStartDate).

    Returns:
        rows          — list of CSV row dicts for non-ICF contacts
        icf_emails    — set of emails that DO have a 37600 membership
    """
    # Group all records by email
    by_email = defaultdict(list)
    for rec in all_records:
        member     = rec.get('individualMember', {})
        membership = rec.get('membership', {})
        email = (member.get('emailAddress') or {}).get('value', '').lower().strip()
        if not email:
            email = f'[no-email]-member-{member.get("id", "unknown")}'
        by_email[email].append((member, membership))

    # Find all emails that have at least one 37600 membership
    icf_emails = set()
    for email, pairs in by_email.items():
        for member, membership in pairs:
            mt_id = (membership.get('membershipType') or {}).get('id')
            if mt_id == ICF_MEMBERSHIP_TYPE:
                icf_emails.add(email)
                break

    # Build rows for non-ICF contacts
    rows = []
    for email, pairs in by_email.items():
        if email in icf_emails:
            continue

        # If this contact has multiple non-37600 memberships, pick the one
        # with the most recent startDate for the primary row
        pairs.sort(key=lambda p: p[1].get('startDate') or 0, reverse=True)
        member, membership = pairs[0]
        props = member.get('properties') or {}

        mt      = membership.get('membershipType') or {}
        address = member.get('address') or {}

        # Volunteer role — pipe-separated list of codes
        vol_roles = props.get('volunteerrole', [])
        if isinstance(vol_roles, list):
            vol_str = '|'.join(r.get('code', '') for r in vol_roles
                               if isinstance(r, dict))
        else:
            vol_str = ''

        # ICF Global auto-renewal — extract code from single-choice object
        autorenewal = extract_prop(props, 'icfglobalautorenewal')
        hasemail    = extract_prop(props, 'hasemail')
        credential  = extract_prop(props, 'icfcredential')
        teamcred    = extract_prop(props, 'icfteamcoachingcredential')

        rows.append({
            'email':                      email,
            'givenName':                  member.get('givenName', ''),
            'familyName':                 member.get('familyName', ''),
            'companyName':                member.get('companyName', ''),
            'positionTitle':              member.get('positionTitle', ''),
            'phone':                      (member.get('phone') or {}).get('value', ''),
            'streetAddress':              address.get('streetAddress', ''),
            'cityName':                   address.get('cityName', ''),
            'zipCode':                    address.get('zipCode', ''),
            'membershipTypeId':           mt.get('id', ''),
            'membershipTypeName':         mt.get('title', ''),
            'membershipId':               membership.get('id', ''),
            'individualMemberId':         member.get('id', ''),
            'membershipStartDate':        ms_to_date(membership.get('startDate')),
            'membershipEndDate':          ms_to_date(membership.get('endDate')),
            'icfmemberid':                props.get('icfmemberid', ''),
            'icfimportdate':              props.get('icfimportdate', ''),
            'icfcredential':              credential,
            'icfcredentialexpiredate':    props.get('icfcredentialexpiredate', ''),
            'icfcredentialawarddate':     props.get('icfcredentialawarddate', ''),
            'icfglobalmembershipstartdate': props.get('icfglobalmembershipstartdate', ''),
            'icfglobalmembershipenddate': props.get('icfglobalmembershipenddate', ''),
            'icfglobalautorenewal':       autorenewal,
            'icfteamcoachingcredential':  teamcred,
            'hasemail':                   hasemail,
            'volunteerrole':              vol_str,
        })

    # Sort output by familyName then givenName
    rows.sort(key=lambda r: (r['familyName'].lower(), r['givenName'].lower()))
    return rows, icf_emails


def write_csv(rows: list, out_path: str) -> None:
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(
        description='Export GlueUp contacts without an ICF Chapter Member membership.'
    )
    parser.add_argument('--out', default=DEFAULT_OUT,
                        help=f'Output CSV path (default: {DEFAULT_OUT})')
    args = parser.parse_args()

    token       = load_token()
    all_records = fetch_all_members(token)
    rows, icf_emails = build_contact_rows(all_records)

    total_contacts  = len(set(
        (rec.get('individualMember') or {}).get('emailAddress', {}).get('value', '').lower()
        for rec in all_records
    ))

    print(f'Total unique contacts in directory: {total_contacts}')
    print(f'Contacts WITH ICF Chapter Member (37600) membership: {len(icf_emails)}')
    print(f'Contacts WITHOUT ICF Chapter Member membership: {len(rows)}')

    if not rows:
        print('\n✓ No contacts found without an ICF Chapter Member membership.')
        return

    write_csv(rows, args.out)
    print(f'\nCSV exported to: {args.out}')


if __name__ == '__main__':
    main()