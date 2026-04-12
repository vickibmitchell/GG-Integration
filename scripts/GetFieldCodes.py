"""
GetFieldCodes.py — Fetch GlueUp Membership Form Field Definitions
ICF Washington State Chapter

Calls the GlueUp API to retrieve all individual membership application
form fields and their valid option codes for single-choice, multi-choice,
and radio button fields.

Saves results to field_codes.json in the same directory. This file is
consumed by ImportICFMembers.py and other scripts to validate and map
field values without hardcoding option codes.

Re-run this script any time a new custom field is added in GlueUp Admin
to keep field_codes.json current.

Usage:
    python3 GetFieldCodes.py           # fetch and save
    python3 GetFieldCodes.py --print   # also print full reference to screen
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

# ── GlueUp API constants ──────────────────────────────────────────────────────
GLUEUP_BASE_URL = 'https://api-services.glueup.com/v2'
ORG_ID          = '7912'
PUBLIC_KEY      = 'icfwshts'
PRIVATE_KEY     = ('MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPf'
                   'Xb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfU'
                   'CCAP67vZiAQ55')
TOKEN_FILE      = os.path.join(os.path.dirname(__file__), 'glueup_token.json')
OUTPUT_FILE     = os.path.join(os.path.dirname(__file__), 'field_codes.json')

# Field types that have option codes
CHOICE_TYPES = {'single_choice', 'multiple_choice', 'checkbox',
                'radio', 'select', 'cascading_selection'}


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


def glueup_get(endpoint: str, token: str) -> dict:
    a_header = build_a_header('GET')
    req = Request(
        f'{GLUEUP_BASE_URL}{endpoint}',
        headers={
            'Content-Type':          'application/json',
            'a':                     a_header,
            'token':                 token,
            'requestOrganizationId': ORG_ID,
            'User-Agent':            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'
        },
        method='GET'
    )
    with urlopen(req) as resp:
        return json.loads(resp.read())


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


def extract_options(options: list) -> list[dict]:
    """
    Recursively extract all option codes and titles from an options array.
    Handles nested cascading_selection options.
    """
    result = []
    for opt in options:
        if not isinstance(opt, dict):
            continue
        code  = opt.get('code', '')
        title = (opt.get('title') or {}).get('en', '') or str(opt.get('title', ''))
        if code:
            result.append({'code': code, 'title': title})
        # Recurse into nested options (cascading_selection)
        if opt.get('options'):
            result.extend(extract_options(opt['options']))
    return result


def merge_fields(all_fields: list[dict]) -> dict:
    """
    Merge field definitions from multiple API responses, keyed by field key.
    Later entries with more data (options) win over earlier sparse entries.
    Returns a dict: field_key → field definition.
    """
    merged = {}
    for field in all_fields:
        key = field.get('key', '')
        if not key:
            continue
        existing = merged.get(key)
        # Prefer the entry that has options defined
        if existing is None:
            merged[key] = field
        elif field.get('options') and not existing.get('options'):
            merged[key] = field
    return merged


def build_field_reference(merged: dict) -> dict:
    """
    Build the field_codes.json structure:
    {
      "generated": "ISO timestamp",
      "fields": {
        "field_key": {
          "key": "field_key",
          "type": "single_choice",
          "isDefault": false,
          "isChoiceType": true,
          "options": [
            {"code": "acc", "title": "ACC"},
            ...
          ]
        },
        ...
      }
    }
    """
    reference = {
        'generated': datetime.utcnow().isoformat() + 'Z',
        'fields':    {}
    }

    for key, field in sorted(merged.items()):
        field_type   = field.get('type', '')
        is_choice    = field_type in CHOICE_TYPES or bool(field.get('options'))
        options      = extract_options(field.get('options', []))

        # Also check 'subtype' — GlueUp uses subtype='select'/'checkbox' on
        # some fields that have type='multiple_choice' or 'url'
        if field.get('subtype') in ('select', 'checkbox', 'radio'):
            is_choice = True

        reference['fields'][key] = {
            'key':        key,
            'type':       field_type,
            'subtype':    field.get('subtype', ''),
            'isDefault':  field.get('isDefault', False),
            'isMandatory': field.get('isMandatory', False),
            'isChoiceType': is_choice,
            'displayName': (field.get('name') or {}).get('en', '') or key,
            'options':    options,
        }

    return reference


def fetch_all_field_definitions(token: str) -> list[dict]:
    """
    Fetch field definitions from two endpoints and combine:
    1. /public/membership/individualApplicationFormFields
       — all fields across all forms (may lack option details)
    2. /public/membership/individualApplicationForms
       — full form definitions including option codes
    """
    all_fields = []

    # Endpoint 1 — all individual application form fields (flat list)
    print('Fetching individual application form fields...', end='', flush=True)
    try:
        resp = glueup_get('/public/membership/individualApplicationFormFields', token)
        fields_flat = resp.get('value', [])
        all_fields.extend(fields_flat)
        print(f' {len(fields_flat)} fields.')
    except HTTPError as e:
        print(f' HTTP {e.code} — skipping.')

    # Endpoint 2 — full form definitions with options (POST, returns all forms)
    print('Fetching full form definitions with option codes...', end='', flush=True)
    try:
        resp = glueup_post('/public/membership/individualApplicationForms', token, {})
        forms = resp.get('value', [])
        for form in forms:
            for field in form.get('fields', []):
                all_fields.append(field)
        print(f' {len(forms)} form(s), {sum(len(f.get("fields",[])) for f in forms)} field instances.')
    except HTTPError as e:
        print(f' HTTP {e.code} — skipping.')

    return all_fields


def print_reference(reference: dict) -> None:
    """Print a human-readable summary of all fields and their option codes."""
    fields = reference['fields']
    total  = len(fields)
    choice = sum(1 for f in fields.values() if f['isChoiceType'])

    print(f'\n{"=" * 70}')
    print(f'GLUEUP FIELD REFERENCE  (generated {reference["generated"]})')
    print(f'{total} total fields  |  {choice} with option codes')
    print('=' * 70)

    # Default fields first, then custom
    default_fields = {k: v for k, v in fields.items() if v['isDefault']}
    custom_fields  = {k: v for k, v in fields.items() if not v['isDefault']}

    print(f'\n── Standard Fields ({len(default_fields)}) ──────────────────────────────')
    for key, f in sorted(default_fields.items()):
        print(f'  {key:<35} type={f["type"]:<20}')

    print(f'\n── Custom Fields ({len(custom_fields)}) ───────────────────────────────────')
    for key, f in sorted(custom_fields.items()):
        display = f['displayName'] or key
        type_str = f['type']
        if f['subtype']:
            type_str += f'/{f["subtype"]}'
        mandatory = ' [mandatory]' if f['isMandatory'] else ''
        print(f'\n  {key}  ({display}){mandatory}')
        print(f'  type: {type_str}')
        if f['isChoiceType'] and f['options']:
            print(f'  options:')
            for opt in f['options']:
                print(f'    {opt["code"]:<30} {opt["title"]}')
        elif f['isChoiceType']:
            print(f'  options: (none returned by API)')

    print('\n' + '=' * 70)


def load_field_codes() -> dict:
    """
    Load field_codes.json for use by other scripts.

    Example usage in another script:
        from GetFieldCodes import load_field_codes
        field_codes = load_field_codes()
        cred_options = field_codes['fields']['icfcredential']['options']
        valid_codes  = [o['code'] for o in cred_options]
    """
    if not os.path.exists(OUTPUT_FILE):
        return {}
    with open(OUTPUT_FILE) as f:
        return json.load(f)


def get_valid_codes(field_key: str) -> list[str]:
    """
    Convenience function: return list of valid codes for a field.
    Returns empty list if field not found or has no options.

    Example:
        codes = get_valid_codes('icfcredential')
        # → ['acc', 'pcc', 'mcc', ...]
    """
    ref = load_field_codes()
    field = (ref.get('fields') or {}).get(field_key, {})
    return [o['code'] for o in field.get('options', [])]


def main():
    parser = argparse.ArgumentParser(
        description='Fetch GlueUp membership form field definitions and save to field_codes.json.'
    )
    parser.add_argument('--print', dest='print_ref', action='store_true',
                        help='Print full field reference to screen after saving')
    args = parser.parse_args()

    token      = load_token()
    all_fields = fetch_all_field_definitions(token)

    if not all_fields:
        print('✗ No field definitions returned. Check token and org ID.')
        sys.exit(1)

    merged    = merge_fields(all_fields)
    reference = build_field_reference(merged)

    with open(OUTPUT_FILE, 'w') as f:
        json.dump(reference, f, indent=2)

    total  = len(reference['fields'])
    choice = sum(1 for f in reference['fields'].values() if f['isChoiceType'])
    print(f'\n✓ {total} fields saved to: {OUTPUT_FILE}')
    print(f'  {choice} fields have option codes (single/multi-choice).')
    print(f'  Generated: {reference["generated"]}')

    if args.print_ref:
        print_reference(reference)
    else:
        print('\nRun with --print to see the full field reference.')
        print('Run again any time a new custom field is added in GlueUp Admin.')


if __name__ == '__main__':
    main()