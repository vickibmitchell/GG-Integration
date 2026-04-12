"""
TestApplicationForm.py — Test GET and PUT on the individual application form endpoint
ICF Washington State Chapter

First GETs the form definition to confirm the endpoint works and to see
the exact field structure expected. Then tries PUT/POST with a payload
that mirrors the form submission structure rather than the directory structure.

Usage:
    python3 TestApplicationForm.py --get     # GET form definition only (safe)
    python3 TestApplicationForm.py --submit  # also try PUT/POST to submit
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


def try_request(method: str, endpoint: str, token: str,
                body: dict | None = None, label: str = '') -> dict | None:
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
    desc = label or f'{method} {endpoint}'
    try:
        with urlopen(req) as resp:
            body_resp = json.loads(resp.read())
            print(f'\n✓ {desc} → {resp.status}')
            print(json.dumps(body_resp, indent=2)[:1500])
            return body_resp
    except HTTPError as e:
        print(f'\n✗ {desc} → {e.code}: {e.read().decode()[:400]}')
        return None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--get',    action='store_true',
                        help='GET form definition only (safe, default)')
    parser.add_argument('--submit', action='store_true',
                        help='Also try PUT/POST submit variations')
    args = parser.parse_args()

    # Default to --get if nothing specified
    if not args.get and not args.submit:
        args.get = True

    token = load_token()

    if args.get:
        print('=' * 60)
        print('Testing GET on application form endpoints...')
        print('=' * 60)

        # GET the default individual application form
        try_request('GET', '/public/membership/individualApplicationForm',
                    token, label='GET /public/membership/individualApplicationForm')

        # GET with membership type ID — may return type-specific form
        try_request('GET',
                    f'/public/membership/individualApplicationForm?membershipTypeId=37600',
                    token,
                    label='GET individualApplicationForm?membershipTypeId=37600')

        # GET all individual application forms
        try_request('POST', '/public/membership/individualApplicationForms',
                    token, body={},
                    label='POST /public/membership/individualApplicationForms (list all forms)')

    if args.submit:
        print('\n' + '=' * 60)
        print('Testing write endpoints with form-style payload...')
        print('=' * 60)

        # Form-style payload — mirrors how a member would submit the application form
        # Uses field keys from the form definition rather than the directory structure
        form_payload = {
            'membershipTypeId': 37600,
            'givenName':        'Tina',
            'familyName':       'Abbott',
            'emailAddress':     {'value': 'tina.abbott@outlook.com'},
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

        # Form ID from the GET response
        form_id = '690b951ee4b0a64517a08ffc'

        # Form-style payload using field IDs from the GET response
        form_id_payload = {
            'formId':           form_id,
            'membershipTypeId': 37600,
            'answers': [
                {'fieldId': '68deb80b2f3d90755c20cec6', 'value': 'Tina'},       # givenName
                {'fieldId': '68deb80b2f3d90755c20cec7', 'value': 'Abbott'},     # familyName
                {'fieldId': '68deb80b2f3d90755c20cec8',
                 'value': 'tina.abbott@outlook.com'},                            # emailAddress
            ]
        }

        # Try PUT and POST on several form-related paths including form ID variants
        for method, endpoint in [
            # Form ID in path variants
            ('POST', f'/public/membership/individualApplicationForm/{form_id}'),
            ('POST', f'/public/membership/individualApplicationForm/{form_id}/submit'),
            ('POST', f'/membership/individualApplicationForm/{form_id}'),
            ('POST', f'/membership/individualApplicationForm/{form_id}/submit'),
            ('POST', f'/public/membership/applicationForm/{form_id}/submit'),
            ('POST', f'/public/membership/applicationForm/{form_id}'),
            # Generic form submission endpoints
            ('PUT',  '/public/membership/individualApplicationForm'),
            ('POST', '/public/membership/individualApplicationForm'),
            ('PUT',  '/membership/individualApplicationForm'),
            ('POST', '/membership/individualApplicationForm'),
            ('POST', '/membership/apply/individual'),
            ('POST', '/membership/application/individual'),
        ]:
            # Use form ID payload for form ID path variants, regular for others
            payload = form_id_payload if form_id in endpoint else form_payload
            try_request(method, endpoint, token, body=payload)

    print('\n' + '=' * 60)
    print('Done.')


if __name__ == '__main__':
    main()