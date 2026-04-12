"""
GetMultiMembers.py — Find GlueUp Contacts With Duplicate Membership Records
ICF Washington State Chapter

Retrieves all members from the GlueUp membership directory and identifies
any contact (matched by email address) who has more than one membership
record. Prints a summary and optionally saves a detailed report.

"Newest" is determined by icfimportdate (most recent import = keep).
Records without icfimportdate were created via the GlueUp membership
workflow (not an ICF Global import) and are flagged separately.

Loads the session token from glueup_token.json (run GetToken.py first).

Usage:
    python3 GetMultiMembers.py
    python3 GetMultiMembers.py --save    # also save multi_members_report.json
    python3 GetMultiMembers.py --csv     # also save a CSV summary
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
GLUEUP_BASE_URL = 'https://api-services.glueup.com/v2'
ORG_ID          = '7912'
PUBLIC_KEY      = 'icfwshts'
PRIVATE_KEY     = ('MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPf'
                   'Xb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfU'
                   'CCAP67vZiAQ55')
TOKEN_FILE  = os.path.join(os.path.dirname(__file__), 'glueup_token.json')
REPORT_FILE = os.path.join(os.path.dirname(__file__), 'multi_members_report.json')
CSV_FILE    = os.path.join(os.path.dirname(__file__), 'multi_members_report.csv')


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
    """Page through /membershipDirectory/members and return every record."""
    all_records = []
    page_size   = 50
    offset      = 0

    print('Fetching all members from GlueUp...', end='', flush=True)

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


def parse_glueup_date(date_str: str) -> datetime | None:
    """
    Parse a GlueUp date string to a datetime for comparison.
    GlueUp returns Date fields as MM/DD/YYYY text strings.
    Returns None if blank or unparseable.
    """
    if not date_str:
        return None
    for fmt in ('%m/%d/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def ms_to_date(ms) -> str:
    """Convert a GlueUp millisecond timestamp to YYYY-MM-DD."""
    if not ms:
        return ''
    try:
        return datetime.utcfromtimestamp(int(ms) / 1000).strftime('%Y-%m-%d')
    except Exception:
        return str(ms)


def sort_key(record: dict) -> tuple:
    """
    Sort key — newest icfimportdate first.
    Records WITH icfimportdate sort before those without.
    Tiebreaker: membershipId descending.

    Returns a tuple that sorts correctly with sorted(...) default (ascending),
    so we negate numeric values and invert the has_import_date boolean.
    """
    import_dt     = parse_glueup_date(record.get('icfimportdate', ''))
    has_import    = import_dt is not None
    import_ts     = import_dt.timestamp() if import_dt else 0
    membership_id = record.get('membershipId') or 0

    # has_import=True should sort first → negate (False=0 < True=1, so negate flips)
    return (not has_import, -import_ts, -membership_id)


def find_duplicates(all_records: list) -> dict:
    """
    Group records by email. Return only contacts with more than one membership,
    sorted so the record to KEEP is always first.
    """
    by_email = defaultdict(list)

    for rec in all_records:
        member     = rec.get('individualMember', {})
        membership = rec.get('membership', {})
        props      = member.get('properties') or {}

        email = (member.get('emailAddress') or {}).get('value', '').lower().strip()
        if not email:
            email = f'[no-email]-member-{member.get("id", "unknown")}'

        by_email[email].append({
            'email':              email,
            'givenName':          member.get('givenName', ''),
            'familyName':         member.get('familyName', ''),
            'individualMemberId': member.get('id'),
            'membershipId':       membership.get('id'),
            'startDate':          ms_to_date(membership.get('startDate')),
            'endDate':            ms_to_date(membership.get('endDate')),
            'icfmemberid':        props.get('icfmemberid', ''),
            'icfimportdate':      props.get('icfimportdate', ''),
        })

    duplicates = {}
    for email, records in by_email.items():
        if len(records) > 1:
            records.sort(key=sort_key)
            duplicates[email] = records

    return duplicates


def action_label(records: list, i: int) -> str:
    """
    Determine the action label for display, using the three-case logic:
      Case 1: Both have icfimportdate → keep most recent import date
      Case 2: Only one has icfimportdate → keep that one
      Case 3: Neither has icfimportdate → flag for manual review
    """
    has_import = [bool(r.get('icfimportdate')) for r in records]

    if not any(has_import):
        # Case 3 — no icfimportdate on any record
        return 'REVIEW (no icfimportdate on any record)'

    if i == 0:
        if has_import[0]:
            return 'KEEP (most recent icfimportdate)' if sum(has_import) > 1 else 'KEEP (only record with icfimportdate)'
        else:
            return 'KEEP (newest membershipId fallback)'
    else:
        return 'DELETE'


def print_summary(duplicates: dict, total_fetched: int) -> None:
    if not duplicates:
        print('✓ No contacts with duplicate membership records found.')
        return

    total_extra = sum(len(v) - 1 for v in duplicates.values())
    print(f'Found {len(duplicates)} contact(s) with duplicate membership records.')
    print(f'Total extra (to-be-deleted) records: {total_extra}')
    print(f'Total records fetched: {total_fetched}\n')
    print('-' * 90)

    for email, records in duplicates.items():
        name = f'{records[0]["givenName"]} {records[0]["familyName"]}'.strip()
        icf  = records[0]['icfmemberid'] or '(none)'
        print(f'\n  {name} <{email}>  |  ICF ID: {icf}')
        print(f'  {"membershipId":>14}  {"startDate":>12}  {"endDate":>12}  '
              f'{"icfimportdate":>14}  Action')

        for i, r in enumerate(records):
            imp  = r['icfimportdate'] or '(none)'
            act  = action_label(records, i)
            print(f'  {r["membershipId"]:>14}  {r["startDate"]:>12}  '
                  f'{r["endDate"]:>12}  {imp:>14}  {act}')

    print('\n' + '-' * 90)
    print('\nRun DeleteMultiMembers.py to remove the duplicate records.')


def save_json(duplicates: dict) -> None:
    with open(REPORT_FILE, 'w') as f:
        json.dump(duplicates, f, indent=2)
    print(f'\nJSON report saved to: {REPORT_FILE}')


def save_csv(duplicates: dict) -> None:
    fieldnames = ['email', 'givenName', 'familyName', 'icfmemberid',
                  'membershipId', 'individualMemberId', 'startDate', 'endDate',
                  'icfimportdate', 'action']
    rows = []
    for email, records in duplicates.items():
        for i, r in enumerate(records):
            rows.append({
                'email':              email,
                'givenName':          r['givenName'],
                'familyName':         r['familyName'],
                'icfmemberid':        r['icfmemberid'],
                'membershipId':       r['membershipId'],
                'individualMemberId': r['individualMemberId'],
                'startDate':          r['startDate'],
                'endDate':            r['endDate'],
                'icfimportdate':      r['icfimportdate'],
                'action':             action_label(records, i)
            })

    with open(CSV_FILE, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f'CSV report saved to: {CSV_FILE}')


def main():
    parser = argparse.ArgumentParser(
        description='Find GlueUp contacts with more than one membership record.'
    )
    parser.add_argument('--save', action='store_true',
                        help='Save full report to multi_members_report.json')
    parser.add_argument('--csv', action='store_true',
                        help='Save JSON and CSV reports (implies --save)')
    args = parser.parse_args()

    token       = load_token()
    all_records = fetch_all_members(token)
    duplicates  = find_duplicates(all_records)

    print_summary(duplicates, len(all_records))

    if args.save or args.csv:
        save_json(duplicates)
    if args.csv:
        save_csv(duplicates)


if __name__ == '__main__':
    main()