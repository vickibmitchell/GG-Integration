"""
DeleteMultiMembers.py — Remove Duplicate GlueUp Membership Records
ICF Washington State Chapter

Finds all contacts with more than one membership record and deletes the
older ones using icfimportdate as the primary decision field:

  Case 1: Both records have icfimportdate → keep the most recent import date
  Case 2: Only one record has icfimportdate → keep that one (it's the ICF
          Global authoritative record; the other came from the workflow)
  Case 3: Neither record has icfimportdate → both came from the GlueUp
          membership workflow. Flagged for MANUAL REVIEW — not auto-deleted.

SAFETY FEATURES:
  - Dry run by default — prints what would be deleted without doing it
  - Requires --confirm to execute actual deletes
  - Asks for interactive confirmation before each delete (override with --yes)
  - Logs every action to delete_log.json
  - Never auto-deletes Case 3 records (no icfimportdate on either record)

Loads the session token from glueup_token.json (run GetToken.py first).

Usage:
    python3 DeleteMultiMembers.py                  # dry run — safe to run anytime
    python3 DeleteMultiMembers.py --confirm        # execute, confirm each record
    python3 DeleteMultiMembers.py --confirm --yes  # execute all without prompting
"""

import argparse
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
TOKEN_FILE = os.path.join(os.path.dirname(__file__), 'glueup_token.json')
LOG_FILE   = os.path.join(os.path.dirname(__file__), 'delete_log.json')


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


def glueup_delete(member_id: int, token: str) -> dict:
    """DELETE /membershipDirectory/member/{memberId}"""
    a_header = build_a_header('DELETE')
    req = Request(
        f'{GLUEUP_BASE_URL}/membershipDirectory/member/{member_id}',
        headers={
            'Content-Type':          'application/json',
            'a':                     a_header,
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'
        },
        method='DELETE'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def parse_glueup_date(date_str: str) -> datetime | None:
    """Parse MM/DD/YYYY or YYYY-MM-DD date strings returned by GlueUp."""
    if not date_str:
        return None
    for fmt in ('%m/%d/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def ms_to_date(ms) -> str:
    if not ms:
        return ''
    try:
        return datetime.utcfromtimestamp(int(ms) / 1000).strftime('%Y-%m-%d')
    except Exception:
        return str(ms)


def sort_key(record: dict) -> tuple:
    """
    Sort so the record to KEEP is always first (index 0).
    Priority: has icfimportdate → most recent icfimportdate → highest membershipId.
    """
    import_dt     = parse_glueup_date(record.get('icfimportdate', ''))
    has_import    = import_dt is not None
    import_ts     = import_dt.timestamp() if import_dt else 0
    membership_id = record.get('membershipId') or 0
    return (not has_import, -import_ts, -membership_id)


def classify(records: list) -> str:
    """
    Return the case classification for a group of duplicate records.
      'both'   — all records have icfimportdate (Case 1)
      'one'    — only some records have icfimportdate (Case 2)
      'neither'— no records have icfimportdate (Case 3 — manual review)
    """
    has_import = [bool(r.get('icfimportdate')) for r in records]
    if all(has_import):
        return 'both'
    if any(has_import):
        return 'one'
    return 'neither'


def fetch_all_members(token: str) -> list:
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


def find_duplicates(all_records: list) -> dict:
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


def save_log(log_entries: list) -> None:
    with open(LOG_FILE, 'w') as f:
        json.dump(log_entries, f, indent=2)
    print(f'\nDelete log saved to: {LOG_FILE}')


def main():
    parser = argparse.ArgumentParser(
        description='Delete duplicate GlueUp membership records using icfimportdate logic.'
    )
    parser.add_argument('--confirm', action='store_true',
                        help='Actually execute deletes (default is dry run)')
    parser.add_argument('--yes', action='store_true',
                        help='Skip per-record confirmation prompts (requires --confirm)')
    args = parser.parse_args()

    if args.yes and not args.confirm:
        print('✗ --yes requires --confirm.')
        sys.exit(1)

    token       = load_token()
    all_records = fetch_all_members(token)
    duplicates  = find_duplicates(all_records)

    if not duplicates:
        print('✓ No contacts with duplicate membership records found. Nothing to do.')
        sys.exit(0)

    # Separate into actionable and manual-review groups
    actionable = {e: r for e, r in duplicates.items() if classify(r) != 'neither'}
    review     = {e: r for e, r in duplicates.items() if classify(r) == 'neither'}
    total_extra = sum(len(v) - 1 for v in actionable.items())

    if not args.confirm:
        print('=' * 80)
        print('DRY RUN — no records will be deleted. Pass --confirm to execute.')
        print('=' * 80)
    else:
        print('=' * 80)
        print(f'LIVE RUN — up to {sum(len(v)-1 for v in actionable.values())} '
              f'record(s) will be deleted.')
        print('=' * 80)

    log_entries = []
    deleted     = 0
    failed      = 0

    # ── Process actionable duplicates ─────────────────────────────────────────
    for email, records in actionable.items():
        keep   = records[0]
        to_del = records[1:]
        case   = classify(records)
        name   = f'{keep["givenName"]} {keep["familyName"]}'.strip()
        icf    = keep['icfmemberid'] or '(none)'

        print(f'\n  {name} <{email}>  |  ICF ID: {icf}  |  Case: {case}')
        print(f'  KEEP   → membershipId {keep["membershipId"]}  '
              f'icfimportdate {keep["icfimportdate"] or "(none)"}  '
              f'startDate {keep["startDate"]}')

        for rec in to_del:
            print(f'  DELETE → membershipId {rec["membershipId"]}  '
                  f'icfimportdate {rec["icfimportdate"] or "(none)"}  '
                  f'startDate {rec["startDate"]}', end='')

            log_entry = {
                'timestamp':              datetime.utcnow().isoformat() + 'Z',
                'case':                   case,
                'email':                  email,
                'name':                   name,
                'icfmemberid':            icf,
                'kept_membershipId':      keep['membershipId'],
                'kept_icfimportdate':     keep['icfimportdate'],
                'del_membershipId':       rec['membershipId'],
                'del_individualMemberId': rec['individualMemberId'],
                'del_icfimportdate':      rec['icfimportdate'],
                'del_startDate':          rec['startDate'],
                'action':                 'dry_run'
            }

            if not args.confirm:
                print('  [DRY RUN]')
                log_entry['action'] = 'dry_run'
            else:
                if not args.yes:
                    answer = input('  — Delete this record? [y/N] ').strip().lower()
                    if answer != 'y':
                        print('  Skipped.')
                        log_entry['action'] = 'skipped'
                        log_entries.append(log_entry)
                        continue
                else:
                    print()

                try:
                    glueup_delete(rec['individualMemberId'], token)
                    print(f'  ✓ Deleted individualMemberId {rec["individualMemberId"]}')
                    log_entry['action'] = 'deleted'
                    deleted += 1
                except HTTPError as e:
                    error_body = e.read().decode()
                    print(f'  ✗ HTTP {e.code} — {e.reason}: {error_body}')
                    log_entry['action']     = 'failed'
                    log_entry['error_code'] = e.code
                    log_entry['error_body'] = error_body
                    failed += 1

            log_entries.append(log_entry)

    # ── Report manual review cases ────────────────────────────────────────────
    if review:
        print(f'\n{"=" * 80}')
        print(f'MANUAL REVIEW REQUIRED — {len(review)} contact(s) have duplicate records')
        print('where neither record has an icfimportdate (both from GlueUp workflow).')
        print(f'{"=" * 80}')
        for email, records in review.items():
            name = f'{records[0]["givenName"]} {records[0]["familyName"]}'.strip()
            print(f'\n  {name} <{email}>')
            for r in records:
                print(f'    membershipId {r["membershipId"]}  '
                      f'startDate {r["startDate"]}  endDate {r["endDate"]}')
            log_entries.append({
                'timestamp': datetime.utcnow().isoformat() + 'Z',
                'case':      'neither',
                'email':     email,
                'name':      name,
                'action':    'manual_review_required',
                'records':   records
            })

    # ── Summary ───────────────────────────────────────────────────────────────
    print('\n' + '=' * 80)
    if not args.confirm:
        print(f'Dry run complete. {sum(len(v)-1 for v in actionable.values())} '
              f'record(s) would be deleted, {len(review)} flagged for manual review.')
        print('Run with --confirm to execute, or --confirm --yes to skip prompts.')
    else:
        print(f'Done. Deleted: {deleted}  Failed: {failed}  '
              f'Manual review: {len(review)}')

    save_log(log_entries)


if __name__ == '__main__':
    main()