"""
TestUpdateMember.py — Verify GET→PUT→DELETE sequence on an existing GlueUp member
ICF Washington State Chapter

Tests the update sequence on Scott Willard's record (created via Admin UI).
Performs a GET to retrieve the current record, PUT to write it back with
icfmemberid and icfimportdate populated, and DELETE to remove the old record.

This is a NON-DESTRUCTIVE test in the sense that the member record will
still exist after the sequence — just as a new record with a new ID.

Usage:
    python3 TestUpdateMember.py --dryrun    # GET only, print record, no PUT/DELETE
    python3 TestUpdateMember.py             # full GET→PUT→DELETE
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

GLUEUP_BASE_URL = 'https://api-services.glueup.com/v2'
ORG_ID          = '7912'
PUBLIC_KEY      = 'icfwshts'
PRIVATE_KEY     = ('MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPf'
                   'Xb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfU'
                   'CCAP67vZiAQ55')
TOKEN_FILE      = os.path.join(os.path.dirname(__file__), 'glueup_token.json')

# Scott Willard — created via Admin UI, ICF member ID 9710829
TEST_EMAIL     = 'scowilla@gmail.com'
TEST_MEMBER_ID = '9710829'


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
        return json.load(f)['token']


def glueup_post(endpoint: str, token: str, body: dict) -> dict:
    req = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        data=json.dumps(body).encode(),
        headers={
            'Content-Type':          'application/json',
            'a':                     build_a_header('POST'),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0'
        },
        method='POST'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def glueup_get(endpoint: str, token: str) -> dict:
    req = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        headers={
            'Content-Type':          'application/json',
            'a':                     build_a_header('GET'),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0'
        },
        method='GET'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def glueup_put(endpoint: str, token: str, body: dict) -> dict:
    req = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        data=json.dumps(body).encode(),
        headers={
            'Content-Type':          'application/json',
            'a':                     build_a_header('PUT'),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0'
        },
        method='PUT'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def glueup_delete(member_id: int, token: str) -> dict:
    req = Request(
        f'{GLUEUP_BASE_URL}/membershipDirectory/member/{member_id}',
        headers={
            'Content-Type':          'application/json',
            'a':                     build_a_header('DELETE'),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0'
        },
        method='DELETE'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def main():
    parser = argparse.ArgumentParser(
        description='Test GET→PUT→DELETE sequence on Scott Willard\'s record.'
    )
    parser.add_argument('--dryrun', action='store_true',
                        help='GET and print only — no PUT or DELETE')
    args = parser.parse_args()

    token = load_token()

    # ── Step 1: Find Scott by email in the directory ──────────────────────────
    print(f'Step 1: Searching for {TEST_EMAIL}...')
    try:
        search_resp = glueup_post('/membershipDirectory/members', token, {
            'search': {
                'fields':   ['emailAddress'],
                'value':    TEST_EMAIL,
                'fullText': True
            },
            'offset': 0,
            'limit':  10
        })
        records = search_resp.get('value', [])
        # Client-side exact match
        match = next((
            r for r in records
            if (r.get('individualMember') or {})
               .get('emailAddress', {}).get('value', '').lower() == TEST_EMAIL
        ), None)

        if not match:
            print(f'✗ No record found for {TEST_EMAIL}.')
            print(f'  Records returned: {len(records)}')
            sys.exit(1)

        individual_member_id = match['individualMember']['id']
        membership_id        = match['membership']['id']
        print(f'✓ Found:  individualMember.id={individual_member_id}  '
              f'membership.id={membership_id}')

    except HTTPError as e:
        print(f'✗ Search failed: HTTP {e.code} — {e.read().decode()}')
        sys.exit(1)

    # ── Step 2: GET full record ───────────────────────────────────────────────
    print(f'\nStep 2: GET /membershipDirectory/member/{individual_member_id}...')
    try:
        get_resp   = glueup_get(
            f'/membershipDirectory/member/{individual_member_id}', token)
        full_record = get_resp.get('value', {})
        print('✓ GET succeeded. Full record:')
        print(json.dumps(full_record, indent=2))
    except HTTPError as e:
        print(f'✗ GET failed: HTTP {e.code} — {e.read().decode()}')
        sys.exit(1)

    if args.dryrun:
        print('\n-- Dry run: stopping before PUT. --')
        return

    # ── Step 3: PUT updated record ────────────────────────────────────────────
    # Build payload from the existing record, adding our ICF fields
    existing_member = full_record.get('individualMember', {})
    existing_props  = existing_member.get('properties') or {}

    payload = {
        'id':             individual_member_id,
        'membershipType': {'id': 37600},
        'emailAddress':   {'value': TEST_EMAIL},
        'givenName':      existing_member.get('givenName', 'SCOTT'),
        'familyName':     existing_member.get('familyName', 'A WILLARD'),
        'properties': {
            **existing_props,
            'icfmemberid':  TEST_MEMBER_ID,
            'hasemail':     {'code': 'yes'},
            'icfimportdate': datetime.utcnow().strftime('%m/%d/%Y'),
            'icfglobalmembershiptype': {'code': 'individual'},
        }
    }

    # Preserve address and phone if present
    if existing_member.get('address'):
        payload['address'] = existing_member['address']
    if existing_member.get('phone'):
        payload['phone'] = existing_member['phone']

    print(f'\nStep 3: PUT /membershipDirectory/member/{individual_member_id}...')
    print(f'Payload: {json.dumps(payload, indent=2)}')
    try:
        put_resp          = glueup_put(
            f'/membershipDirectory/member/{individual_member_id}', token, payload)
        new_member_id     = (put_resp.get('value') or {}).get(
            'individualMember', {}).get('id')
        new_membership_id = (put_resp.get('value') or {}).get(
            'membership', {}).get('id')
        print(f'✓ PUT succeeded.')
        print(f'  New individualMember.id={new_member_id}  '
              f'New membership.id={new_membership_id}')
        print(f'  Full response: {json.dumps(put_resp, indent=2)}')
    except HTTPError as e:
        print(f'✗ PUT failed: HTTP {e.code} — {e.read().decode()}')
        sys.exit(1)

    # ── Step 4: DELETE old record ─────────────────────────────────────────────
    print(f'\nStep 4: DELETE old individualMember.id={individual_member_id}...')
    try:
        del_resp = glueup_delete(individual_member_id, token)
        print(f'✓ DELETE succeeded.')
        print(f'  Response: {json.dumps(del_resp, indent=2)}')
    except HTTPError as e:
        print(f'✗ DELETE failed: HTTP {e.code} — {e.read().decode()}')
        print(f'  ⚠  PUT succeeded but DELETE failed — duplicate record now exists.')
        print(f'     New individualMember.id={new_member_id} (keep this one)')
        print(f'     Old individualMember.id={individual_member_id} (delete manually)')
        sys.exit(1)

    print(f'\n{"=" * 60}')
    print(f'✓ GET→PUT→DELETE sequence completed successfully.')
    print(f'  Scott Willard now has:')
    print(f'    individualMember.id = {new_member_id}')
    print(f'    membership.id       = {new_membership_id}')
    print(f'    icfmemberid         = {TEST_MEMBER_ID}')
    print(f'    icfimportdate       = {datetime.utcnow().strftime("%m/%d/%Y")}')


if __name__ == '__main__':
    main()