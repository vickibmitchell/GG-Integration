"""
TestMethods.py — Test PATCH and DELETE on an existing GlueUp member record
ICF Washington State Chapter

Tests PATCH and DELETE on Scott Willard's record to determine which
write methods the v2 API actually supports.

Usage:
    python3 TestMethods.py --patch-only    # test PATCH only, no DELETE
    python3 TestMethods.py --delete-only   # test DELETE only
    python3 TestMethods.py                 # test PATCH only (safe default)
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

TEST_EMAIL         = 'scowilla@gmail.com'
TEST_MEMBER_ID     = '9710829'


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


def make_request(method: str, endpoint: str, token: str,
                 body: dict | None = None) -> dict:
    data = json.dumps(body).encode() if body else None
    req  = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        data=data,
        headers={
            'Content-Type':          'application/json',
            'a':                     build_a_header(method),
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0'
        },
        method=method
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


def find_scott(token: str) -> tuple[int, int]:
    """Find Scott Willard and return (individualMemberId, membershipId)."""
    resp    = make_request('POST', '/membershipDirectory/members', token, {
        'search': {'fields': ['emailAddress'], 'value': TEST_EMAIL, 'fullText': True},
        'offset': 0, 'limit': 10
    })
    records = resp.get('value', [])
    match   = next((
        r for r in records
        if (r.get('individualMember') or {})
           .get('emailAddress', {}).get('value', '').lower() == TEST_EMAIL
    ), None)
    if not match:
        print(f'✗ Could not find {TEST_EMAIL} in GlueUp.')
        sys.exit(1)
    ind_id  = match['individualMember']['id']
    mem_id  = match['membership']['id']
    print(f'✓ Found Scott: individualMember.id={ind_id}  membership.id={mem_id}')
    return ind_id, mem_id


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--patch-only',  action='store_true',
                        help='Test PATCH only (default safe mode)')
    parser.add_argument('--delete-only', action='store_true',
                        help='Test DELETE only — will remove Scott\'s record')
    args = parser.parse_args()

    token          = load_token()
    ind_id, mem_id = find_scott(token)

    # ── Test PATCH ────────────────────────────────────────────────────────────
    if not args.delete_only:
        # Minimal PATCH payload — just try to set icfmemberid
        patch_payload = {
            'properties': {
                'icfmemberid':   TEST_MEMBER_ID,
                'icfimportdate': datetime.utcnow().strftime('%m/%d/%Y'),
            }
        }
        print(f'\nTesting PATCH /membershipDirectory/member/{ind_id}...')
        try:
            resp = make_request('PATCH',
                                f'/membershipDirectory/member/{ind_id}',
                                token, patch_payload)
            print(f'✓ PATCH succeeded!')
            print(json.dumps(resp, indent=2))
        except HTTPError as e:
            print(f'✗ PATCH failed: HTTP {e.code} — {e.read().decode()}')

    # ── Test DELETE ───────────────────────────────────────────────────────────
    if args.delete_only:
        print(f'\nTesting DELETE /membershipDirectory/member/{ind_id}...')
        print(f'⚠  This will remove Scott Willard\'s record from GlueUp.')
        confirm = input('Type YES to proceed: ').strip()
        if confirm != 'YES':
            print('Cancelled.')
            return
        try:
            resp = make_request('DELETE',
                                f'/membershipDirectory/member/{ind_id}',
                                token)
            print(f'✓ DELETE succeeded!')
            print(json.dumps(resp, indent=2))
        except HTTPError as e:
            print(f'✗ DELETE failed: HTTP {e.code} — {e.read().decode()}')


if __name__ == '__main__':
    main()