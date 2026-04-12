import hmac, hashlib, time, json
from urllib.request import urlopen, Request
from urllib.error import HTTPError

PUBLIC_KEY  = 'icfwshts'
PRIVATE_KEY = ('MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPf'
               'Xb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfU'
               'CCAP67vZiAQ55')

with open('glueup_token.json') as f:
    token = json.load(f)['token']

payload = {
    'membershipType': {'id': 37600},
    'emailAddress':   {'value': 'tina.abbott@outlook.com'},
    'givenName':      'Tina',
    'familyName':     'Abbott',
    'properties':     {'icfmemberid': '9334468'}
}

for method, endpoint in [
    ('PUT',  '/membershipDirectory/member'),
    ('PUT',  '/membershipDirectory/members'),
    ('PUT',  '/membershipDirectory/individualMembership'),
    ('POST', '/membershipDirectory/member'),
    ('POST', '/membershipDirectory/individualMembership'),
]:
    ts  = str(int(time.time() * 1000))
    msg = method + PUBLIC_KEY + '1.0' + ts
    d   = hmac.new(PRIVATE_KEY.encode(), msg.encode(), hashlib.sha256).hexdigest()
    a   = f'v=1.0;k={PUBLIC_KEY};ts={ts};d={d}'

    try:
        req = Request(
            f'https://api-services.glueup.com/v2{endpoint}',
            data=json.dumps(payload).encode(),
            headers={
                'Content-Type':          'application/json',
                'a':                     a,
                'token':                 token,
                'requestOrganizationId': '7912',
                'User-Agent':            'Mozilla/5.0'
            },
            method=method
        )
        with urlopen(req) as resp:
            body = json.loads(resp.read())
            print(f'\n✓ {method} {endpoint} → {resp.status}')
            print(json.dumps(body, indent=2)[:500])
    except HTTPError as e:
        body = e.read().decode()[:200]
        print(f'\n✗ {method} {endpoint} → {e.code}: {body}')