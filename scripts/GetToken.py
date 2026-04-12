"""
GetToken.py — GlueUp Authentication Helper
ICF Washington State Chapter

Prompts for email and password, hashes the password with MD5,
computes the HMAC-SHA256 'a' header, and authenticates to GlueUp.

Saves the session token to glueup_token.json in the same directory
so other scripts can load it without re-authenticating.

Usage:
    python3 GetToken.py

Output:
    glueup_token.json  — { "token": "...", "email": "...", "expires": "..." }
"""

import hmac
import hashlib
import time
import json
import getpass
import os
from urllib.request import urlopen, Request
from urllib.error import HTTPError

# ── GlueUp API constants ──────────────────────────────────────────────────────
GLUEUP_BASE_URL  = 'https://api-services.glueup.com/v2'
PUBLIC_KEY       = 'icfwshts'
PRIVATE_KEY      = ('MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPf'
                    'Xb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfU'
                    'CCAP67vZiAQ55')
TOKEN_FILE       = os.path.join(os.path.dirname(__file__), 'glueup_token.json')


def md5_hash(password: str) -> str:
    """Return the MD5 hex digest of a plain-text password."""
    return hashlib.md5(password.encode()).hexdigest()


def build_a_header(method: str) -> str:
    """Compute the HMAC-SHA256 signed 'a' header for a GlueUp request."""
    ts  = str(int(time.time() * 1000))
    msg = method + PUBLIC_KEY + '1.0' + ts
    digest = hmac.new(
        PRIVATE_KEY.encode(),
        msg.encode(),
        hashlib.sha256
    ).hexdigest()
    return f'v=1.0;k={PUBLIC_KEY};ts={ts};d={digest}'


def get_token(email: str, md5_password: str) -> dict:
    """
    POST to /v2/user/session and return the full parsed response.
    Raises HTTPError on failure.
    """
    a_header = build_a_header('POST')

    body = json.dumps({
        'email':      {'value': email},
        'passphrase': {'value': md5_password}
    }).encode()

    req = Request(
        f'{GLUEUP_BASE_URL}/user/session',
        data=body,
        headers={
            'Content-Type': 'application/json',
            'a':            a_header,
            'User-Agent':   'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)'
        },
        method='POST'
    )

    with urlopen(req) as resp:
        return json.loads(resp.read())


def save_token(token: str, email: str, expires=None) -> None:
    """Persist token to glueup_token.json for use by other scripts."""
    payload = {
        'token':   token,
        'email':   email,
        'expires': expires
    }
    with open(TOKEN_FILE, 'w') as f:
        json.dump(payload, f, indent=2)
    print(f'\nToken saved to: {TOKEN_FILE}')


def load_token() -> dict | None:
    """
    Load a previously saved token from glueup_token.json.
    Returns the dict, or None if the file does not exist.
    Call this from other scripts instead of re-authenticating.

    Example usage in another script:
        from GetToken import load_token
        token_data = load_token()
        token = token_data['token']
    """
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE) as f:
            return json.load(f)
    return None


def main():
    print('=== GlueUp Authentication ===\n')
    email    = input('Email: ').strip()
    password = getpass.getpass('Password (input hidden): ')

    md5_password = md5_hash(password)
    print(f'\nMD5 hash: {md5_password}')
    print('Authenticating to GlueUp...')

    try:
        result  = get_token(email, md5_password)
        token   = result['value']['token']
        expires = result['value'].get('expiry') or result['value'].get('expires')

        print('\n✓ Authentication successful.')
        print(f'Token: {token[:40]}...  (truncated for display)')
        if expires:
            print(f'Expires: {expires}')

        save_token(token, email, expires)

    except HTTPError as e:
        print(f'\n✗ HTTP {e.code}: {e.reason}')
        print(e.read().decode())
    except KeyError:
        print('\n✗ Unexpected response structure:')
        print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()