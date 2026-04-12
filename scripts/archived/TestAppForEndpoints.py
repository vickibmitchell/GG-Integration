"""
TestAppFormEndpoints.py — Test membership application form endpoints for writes
ICF Washington State Chapter

Tests whether the individual membership application form endpoints
accept POST/PUT to create a new membership record for Tina Abbott.

Usage:
    python3 TestAppFormEndpoints.py
"""

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

# Tina Abbott — not yet in GlueUp, has ACC credential
TEST_PAYLOAD = {
    'membershipType': {'id': 37600},
    'emailAddress':   {'value': 'tina.abbott@outlook.com'},
    'givenName':      'Tina',
    'familyName':     'Abbott',
    'properties': {
        'icfmemberid':                  '9334468',
        'hasemail':                     {'code': 'yes'},
        'icfimportdate':                datetime.utcnow().strftime('%m/%d/%Y'),
        'icfglobalmembershiptype':      {'code': 'individual'},
        'icfglobalmembershipstartdate': '04/25/2022',
        'icfglobalmembershipenddate':   '03/31/2026',
        'icfcredential':                {'code': 'acc'},
        'icfcredentialawarddate':       '04/25/2022',
        'icfcredentialexpiredate':      '04/30/2028',
        'icfglobalautorenewal':         {'code': 'yes'},
    }
}


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


def try_endpoint(method: str, endpoint: str, token: str,
                 body: dict | None = None) -> None:
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
    try:
        with urlopen(req) as resp:
            body_resp = json.loads(resp.read())
            print(f'\n✓ {method} {endpoint} → {resp.status}')
            print(json.dumps(body_resp, indent=2)[:800])
    except HTTPError as e:
        print(f'\n✗ {method} {endpoint} → {e.code}: {e.read().decode()[:300]}')


def main():
    token = load_token()
    print('Testing application form and membership endpoints for write access...\n')
    print(f'Test payload: Tina Abbott <tina.abbott@outlook.com> (ICF ID 9334468)')

    endpoints = [
        ('POST', '/public/membership/individualApplicationForm'),
        ('PUT',  '/public/membership/individualApplicationForm'),
        ('POST', '/membership/individualApplicationForm'),
        ('PUT',  '/membership/individualApplicationForm'),
        ('POST', '/membership/application'),
        ('POST', '/membership/apply'),
        ('POST', '/membershipDirectory/individualMemberships'),
        ('POST', '/membership/member'),
        ('PUT',  '/membership/member'),
    ]

    for method, endpoint in endpoints:
        try_endpoint(method, endpoint, token, TEST_PAYLOAD)

    print('\n' + '=' * 60)
    print('Done. Any ✓ above indicates a working write endpoint.')


if __name__ == '__main__':
    main()