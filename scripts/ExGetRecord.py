"""
ExGetRecord.py — GlueUp Member Record Search
ICF Washington State Chapter

Searches for GlueUp member records by one of:
  --icfmemberid   ICF Global Member ID (custom field)
  --email         Email address
  --firstname     First name  (can combine with --lastname)
  --lastname      Last name   (can combine with --firstname)

Loads the session token from glueup_token.json (run GetToken.py first).
Prints matching records as formatted JSON.

Usage:
    python3 ExGetRecord.py --icfmemberid 9710829
    python3 ExGetRecord.py --email jane.doe@example.com
    python3 ExGetRecord.py --firstname Jane --lastname Doe
    python3 ExGetRecord.py --lastname Doe
"""

import argparse
import json
import hmac
import hashlib
import time
import os
import sys
from urllib.request import urlopen, Request
from urllib.error import HTTPError

# ── GlueUp API constants ──────────────────────────────────────────────────────
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
        data = json.load(f)
    return data['token']


def glueup_post(endpoint: str, token: str, body: dict) -> dict:
    """POST to a GlueUp endpoint and return parsed JSON."""
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


def build_request(args) -> dict:
    """
    Build the GlueUp request body using the correct field names from the API blueprint.

    - icfmemberid uses a filter on properties.icfmemberid (custom field, exact match)
    - email/name use the 'search' block which supports full-text matching across
      the fields GlueUp actually indexes (givenName, familyName, emailAddress, companyName)
    - If both icfmemberid AND name/email are provided, icfmemberid is used as a
      filter and name/email as the search term (ANDed together by GlueUp)
    """
    body = {
        'projection': [],
        'offset': 0,
        'limit':  50,
        'order':  {'familyName': 'asc'}
    }

    # Custom field filter — exact match
    if args.icfmemberid:
        body['filter'] = [
            {
                'projection': 'properties.icfmemberid',
                'operator':   'eq',
                'values':     [args.icfmemberid]
            }
        ]

    # Build search term from email / name args
    # GlueUp's search block does full-text matching across indexed fields
    search_terms = []
    if args.email:
        search_terms.append(args.email)
    if args.firstname:
        search_terms.append(args.firstname)
    if args.lastname:
        search_terms.append(args.lastname)

    if search_terms:
        body['search'] = {
            'fields':   ['givenName', 'familyName', 'emailAddress', 'companyName'],
            'value':    ' '.join(search_terms),
            'fullText': True
        }

    return body


def client_filter(records: list, args) -> list:
    """
    Apply exact client-side filtering after GlueUp returns results.
    GlueUp's search is full-text and may return partial matches, so
    we do a precise check here to return only truly matching records.
    """
    results = []
    for rec in records:
        member = rec.get('individualMember', rec)  # handle both response shapes

        email  = (member.get('emailAddress') or {}).get('value', '').lower()
        given  = member.get('givenName', '').lower()
        family = member.get('familyName', '').lower()
        props  = member.get('properties', {})
        icf_id = str(props.get('icfmemberid', ''))

        match = True

        if args.icfmemberid and icf_id != args.icfmemberid:
            match = False
        if args.email and email != args.email.lower():
            match = False
        if args.firstname and given != args.firstname.lower():
            match = False
        if args.lastname and family != args.lastname.lower():
            match = False

        if match:
            results.append(rec)

    return results


def search_members(token: str, body: dict) -> list:
    """
    Page through /membershipDirectory/members and collect all results.
    Uses offset-based pagination as documented in the blueprint.
    """
    all_records = []
    page_size   = body.get('limit', 50)

    while True:
        result  = glueup_post('/membershipDirectory/members', token, body)
        records = result.get('value', [])

        if not records:
            break

        all_records.extend(records)

        # Stop if we received fewer records than the page size (last page)
        if len(records) < page_size:
            break

        body['offset'] += page_size

    return all_records


def main():
    parser = argparse.ArgumentParser(
        description='Search GlueUp member records by icfmemberid, email, or name.'
    )
    parser.add_argument('--icfmemberid', help='ICF Global Member ID (custom field, exact match)')
    parser.add_argument('--email',       help='Member email address (exact match)')
    parser.add_argument('--firstname',   help='First name (exact match)')
    parser.add_argument('--lastname',    help='Last name (exact match)')

    args = parser.parse_args()

    if not any([args.icfmemberid, args.email, args.firstname, args.lastname]):
        parser.print_help()
        print('\n✗ Please provide at least one search parameter.')
        sys.exit(1)

    token = load_token()
    body  = build_request(args)

    print('Searching GlueUp...\n')

    try:
        raw_records = search_members(token, body)
        matches     = client_filter(raw_records, args)

        if not matches:
            print(f'No matching records found. (GlueUp returned {len(raw_records)} candidates.)')
        else:
            print(f'Found {len(matches)} matching record(s):\n')
            print(json.dumps(matches, indent=2))

    except HTTPError as e:
        print(f'✗ HTTP {e.code}: {e.reason}')
        print(e.read().decode())


if __name__ == '__main__':
    main()