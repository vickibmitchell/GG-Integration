"""
ICFGlueUpSync.py  —  v1.7.0  (2026-05-15)
ICF Washington State Chapter — GlueUp Member Sync Comparison Tool

CHANGELOG
---------
v1.7.0  2026-05-15
  - Cloud Run migration (Open Item 15): replaced OAuth 2.0 browser flow in
    _get_drive_service() with Application Default Credentials (ADC). When
    running on Cloud Run the service account assigned to the job supplies
    credentials automatically via the GCP metadata server — no key file or
    browser interaction required. Local execution continues to work via
    'gcloud auth application-default login'.
  - Added _load_secrets_from_secret_manager() helper. When the environment
    variable GLUEUP_USE_SECRET_MANAGER=1 is set (as it will be in Cloud Run),
    GLUEUP_SK and GLUEUP_MD5_PW are pulled from GCP Secret Manager at startup,
    overriding the hardcoded fallback values. The GCP project ID is read from
    GOOGLE_CLOUD_PROJECT (set automatically by Cloud Run).
  - Hardcoded credential fallbacks retained for local execution compatibility.
  - DRIVE_SCOPES broadened from drive.file to drive so the service account
    can access files in Shared Drives it has been granted Content Manager on.
  - Removed drive_token.json and credentials.json references (no longer used
    in Cloud Run; local ADC flow does not require them).

v1.6.1  2026-04-18
  - Removed 'creation date' → 'Chapter_Start_Date' alias from HEADER_ALIASES.
    The Make CSV has both a real 'Chapter_Start_Date' column (col L, populated)
    and a legacy 'Creation Date' column appended at position 22 (blank). The
    alias caused the blank legacy column to overwrite the real value, making
    Chapter_Start_Date appear blank and triggering the endDate-1yr fallback for
    all members. This produced incorrect Membership Start Dates in the import
    file, causing 121 spurious duplicate records in GlueUp today. The real
    column ('chapter_start_date') still maps correctly without the alias.

v1.6.0  2026-04-18
  - Expanded COUNTRY_CODES from 5 entries to 80+ covering all commonly seen
    countries in ICF member data. The SOAP API returns full country names
    (e.g. "JAPAN") but GlueUp requires 2-letter ISO codes (e.g. "JP") —
    previously unmapped countries passed through as full names and caused
    GlueUp import failures. ESTONIA added so Elena Raudsepp's shifted row
    continues to be caught as non-standard.
  - Added _map_country() helper used by write_import_file for Country/Region.
    Prints a named warning to stdout for any country not in COUNTRY_CODES,
    making unmapped values visible rather than silently wrong.
  - KNOWN_COUNTRIES in read_icf_file now derives automatically from
    COUNTRY_CODES keys and values — no longer a separate duplicate list.

v1.5.9  2026-04-18
  - Non-standard record detection: some ICF Global members (e.g. Estonian
    addresses) have an extra address sub-field in the SOAP response that shifts
    all subsequent columns one position to the right, producing garbled data.
    read_icf_file() now checks whether the Country field contains a recognised
    value; if not, the row is flagged as non-standard, excluded from both
    comparison and import, and reported on a new "Non-Standard Records" sheet
    in the comparison report with the reason and a note to import manually.
    KNOWN_COUNTRIES covers all countries currently seen in the member data
    plus common additions; add entries as new countries appear.

v1.5.8  2026-04-18
  - Fixed "20 Cells found for column startDate with missing mandatory information"
    GlueUp error. Membership Start Date was always blank because the SOAP API
    does not reliably return Chapter_Start_Date. Added _derive_start_date()
    helper that subtracts one year from the expiration date as a fallback.
    Both Membership Start Date and ICF Global Membership Start Date now use:
    Chapter_Start_Date if available, otherwise expiry_date minus one year.

v1.5.7  2026-04-18
  - Skip ICF Credential and TC Credential comparison when ICF Global returns
    blank but GlueUp has a value. The SOAP API drops credentials that have
    expired, so a blank is not a signal to overwrite GlueUp's historical data.
    Previously only the credential date fields were skipped; now the credential
    code fields themselves (ICF Credential, TC Credential) are skipped too.

v1.5.6  2026-04-18
  - Removed dead code: US_STATE_ABBREVS table and the State/Province
    normalization in extract_icf_fields are no longer needed since State is
    not compared. Import file passes through whatever the SOAP API provides
    (abbreviation e.g. 'WA'), which is the desired behaviour — members in
    other countries/states will get whatever ICF Global returns for them.

v1.5.5  2026-04-18
  - Removed State/Province from COMPARISON_FIELDS entirely.
    GlueUp stores inconsistent values ('WA' vs 'Washington') depending on
    import history, and the field is not in the GlueUp import template so it
    cannot drive a meaningful CHANGED classification. The per-comparison
    State/Province normalization block in compare_records is also removed.
    Field is still captured in the comparison report for reference.

v1.5.4  2026-04-18
  - Hotfix: State normalization was applied to the ICF side only, but GlueUp
    is the side returning the full name (e.g. "Washington"). The comparison
    loop now normalizes BOTH sides to 2-letter abbreviation via US_STATE_ABBREVS
    before comparing, so the direction of the mismatch doesn't matter.

v1.5.3  2026-04-18
  - State normalization updated: the ICF Global SOAP API returns full state
    names (e.g. "Washington") while the manual XLS export returned abbreviations
    (e.g. "WA"). extract_icf_fields() now looks up the full name in a
    US_STATE_ABBREVS table and converts it to a 2-letter code before comparison,
    matching how GlueUp stores the value. If the value is already an abbreviation
    (2 chars or not found in the table), it is uppercased and kept as-is.
    Added US_STATE_ABBREVS constant covering all 50 states plus DC and territories.

v1.5.2  2026-04-18
  - Removed Chapter_Start_Date (Creation Date) from COMPARISON_FIELDS.
    The ICF Global SOAP API does not reliably return this field (it is blank
    in the Make CSV output), causing every record to appear CHANGED when
    compared against the value stored in GlueUp. Per the design spec, Creation
    Date was always excluded from the fingerprint — it does not change once set.
    The field is still written to the import file via icf('Chapter_Start_Date').
  - Removed the DATE_TOLERANCE entry for 'Creation Date' (no longer compared).

v1.5.1  2026-04-18
  - Hotfix: the actual CSV produced by the Make scenario uses a mix of old
    manual-export header names (e.g. "Member ID", "First Name", "Expiration Date")
    and new API-derived names (e.g. "Chapter_Start_Date", "ACTC_Credential").
    read_icf_file() now applies a HEADER_ALIASES map that normalises any known
    alias to the canonical name the rest of the script expects, so both old and
    new header formats are accepted transparently. Unrecognised headers pass
    through unchanged and are logged to stdout.

v1.5  2026-04-18
  - Input file changed from .xlsx (PC sheet) to .csv produced by the Make
    "Get Active Members" scenario (activemembers_YYYYMMDD.csv in Data-Transfer
    > Inbound shared Drive folder).
  - Updated find_latest_icf_file() to search for *activemembers*.csv instead
    of *activemembers*.xlsx.
  - Updated read_icf_file() to use csv.DictReader instead of openpyxl; file
    now has a single header row (row 1) with no skip rows.
  - Column names updated to match the Make scenario output headers:
      Member ID           → Member_ID
      First Name          → First_Name
      Last Name           → Last_Name
      Creation Date       → Chapter_Start_Date / Membership_Join_Date
      Expiration Date     → Membership_Expiration_Date
      Credential          → Flagship_Credential
      Team Coaching Credential → ACTC_Credential
      TC Award/Expire Date → ACTC_Credential_Award/Expire_Date
      Reinstate/Rejoin    → Rejoin
      Auto Renewal        → Auto_Renewal
  - COMPARISON_FIELDS icf_column_name entries updated to new header names.
  - Date normalisation updated for MM/DD/YYYY format written by Make.
  - Member ID and Zip float-guard updated for CSV string values.
  - Command-line help text updated to reference .csv.

v1.4  2026-04-01
  - Fixed duplicate classification: use native membership.endDate timestamp
    instead of icfglobalmembershipenddate custom field to determine whether
    expiry dates differ. The renewal import updates the custom field on both
    records to the new date, making them look identical; the native end date
    correctly reflects the two different expiry dates, so WAIT duplicates
    (renewals) are now classified correctly instead of being flagged DELETE OLD.

v1.3  2026-04-01
  - Phone comparison: skip if either side is blank (previously only skipped
    when GlueUp had no phone; now also skips when ICF Global has no phone,
    preventing spurious CHANGED when GlueUp has a number ICF doesn't).
  - State/Province comparison: skip for non-US members (international
    addresses don't use US-style state codes; avoids spurious CHANGED
    for members with UK, CA, etc. addresses).

v1.2  2026-04-01
  - Fixed State/Province and Country/Region field paths in COMPARISON_FIELDS:
    province (not stateName) and address.country {code} (not countryName).
  - extract_icf_fields now normalizes Country to ISO code (via COUNTRY_CODES)
    and State to uppercase before comparison, to match GlueUp storage format.

v1.1  2026-04-01
  - Phone normalization: strip non-digits and compare last 10 digits on both
    sides so format differences (e.g. 512.773.5096 vs +1 5127735096) are
    treated as SAME.
  - Added State/Province and Country/Region to COMPARISON_FIELDS so changes
    to those fields are now detected (they are stored in GlueUp as of the
    April 2026 import template update).
  - Note in comment: State and Country comparison is intent-only for now;
    the _normalise_phone helper is the active fix.
  - Version constant added: SCRIPT_VERSION.

v1.0  2026-03-31
  - Initial release.

PURPOSE
-------
1. Reads the ICF Global member export CSV produced by the Make scenario
   "Get Active Members" (activemembers_YYYYMMDD.csv in Data-Transfer > Inbound).
2. Pulls the full active member list from GlueUp via API.
3. Compares records using ICF Global Member ID as the anchor key.
4. Classifies every ICF Global record as: NEW, CHANGED, or SAME.
5. Produces three output files:
     a) comparison_report_<date>.xlsx   — full comparison with status column
     b) glueup_import_<date>.xlsx       — NEW + CHANGED records, formatted for
                                          GlueUp Admin UI bulk import
     c) duplicate_members_<date>.xlsx   — GlueUp members where the same
                                          icfmemberid appears on more than one
                                          active membership record

USAGE
-----
    python ICFGlueUpSync.py [path/to/activemembers_YYYYMMDD.csv]

    If no file path is given, the script looks for the most recently
    modified .csv in the current directory whose name contains
    "activemembers" (case-insensitive).

CREDENTIALS
-----------
Set the four constants below before running, or export them as
environment variables:
    GLUEUP_PK          — GlueUp API public key
    GLUEUP_SK          — GlueUp API private key
    GLUEUP_MD5_PW      — MD5 hash of GlueUp login password
    GLUEUP_EMAIL       — GlueUp login email

REQUIREMENTS
------------
    pip install openpyxl requests google-api-python-client google-auth
    google-auth-httplib2 google-cloud-secret-manager
"""

import csv
import glob
import hmac
import hashlib
import json
import os
import sys
import time
from datetime import datetime, date
from pathlib import Path

import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────────────────────
# CREDENTIALS — edit here or set as environment variables
# ──────────────────────────────────────────────────────────────────────────────
# When running on Cloud Run, GLUEUP_USE_SECRET_MANAGER=1 is set as an
# environment variable and secrets are fetched from GCP Secret Manager at
# startup (see _load_secrets_from_secret_manager() below), overriding the
# hardcoded fallback values here. For local execution, the fallback values
# are used directly — no Secret Manager access required.
GLUEUP_PK       = os.environ.get('GLUEUP_PK',      'icfwshts')
GLUEUP_SK       = os.environ.get('GLUEUP_SK',      'MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPfXb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfUCCAP67vZiAQ55')
GLUEUP_MD5_PW   = os.environ.get('GLUEUP_MD5_PW',  '70bdb05ade647079069cfebed391758a')
GLUEUP_EMAIL    = os.environ.get('GLUEUP_EMAIL',   'technology@icfwashingtonstate.org')

SCRIPT_VERSION  = '1.6.1'

# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────────────────────
GLUEUP_ORG_ID       = '7912'
GLUEUP_BASE_URL     = 'https://api-services.glueup.com/v2'
GLUEUP_MEMBERSHIP_TYPE_ID   = 37600
GLUEUP_MEMBERSHIP_TYPE_NAME = 'ICF Chapter Member'
SHADOW_EMAIL_DOMAIN = 'members.icfwashingtonstate.org'

# Country name → ISO 3166-1 alpha-2 code mapping.
# Keys are uppercase full names; values are 2-letter ISO codes as required
# by GlueUp's import template. Add entries here when new countries appear
# in the ICF Global export. The SOAP API returns full country names;
# the import file requires 2-letter codes.
COUNTRY_CODES = {
    # Already-seen in ICF WA member data
    'UNITED STATES': 'US',
    'CANADA':        'CA',
    'CHINA':         'CN',
    'SWEDEN':        'SE',
    'UNITED KINGDOM':'GB',
    'JAPAN':         'JP',
    'ESTONIA':       'EE',
    'AUSTRALIA':     'AU',
    'GERMANY':       'DE',
    'FRANCE':        'FR',
    'NETHERLANDS':   'NL',
    'BRAZIL':        'BR',
    'INDIA':         'IN',
    'SINGAPORE':     'SG',
    'NEW ZEALAND':   'NZ',
    'IRELAND':       'IE',
    'DENMARK':       'DK',
    'NORWAY':        'NO',
    'FINLAND':       'FI',
    'BELGIUM':       'BE',
    'AUSTRIA':       'AT',
    'SWITZERLAND':   'CH',
    'SPAIN':         'ES',
    'ITALY':         'IT',
    'PORTUGAL':      'PT',
    'MEXICO':        'MX',
    'COLOMBIA':      'CO',
    'CHILE':         'CL',
    'ARGENTINA':     'AR',
    'SOUTH AFRICA':  'ZA',
    'KENYA':         'KE',
    'NIGERIA':       'NG',
    'ISRAEL':        'IL',
    'UAE':           'AE',
    'UNITED ARAB EMIRATES': 'AE',
    'SAUDI ARABIA':  'SA',
    'HONG KONG':     'HK',
    'TAIWAN':        'TW',
    'SOUTH KOREA':   'KR',
    'KOREA':         'KR',
    'THAILAND':      'TH',
    'PHILIPPINES':   'PH',
    'INDONESIA':     'ID',
    'MALAYSIA':      'MY',
    'VIETNAM':       'VN',
    'PAKISTAN':      'PK',
    'BANGLADESH':    'BD',
    'SRI LANKA':     'LK',
    'NEPAL':         'NP',
    'GHANA':         'GH',
    'ETHIOPIA':      'ET',
    'EGYPT':         'EG',
    'MOROCCO':       'MA',
    'TURKEY':        'TR',
    'GREECE':        'GR',
    'POLAND':        'PL',
    'CZECH REPUBLIC':'CZ',
    'SLOVAKIA':      'SK',
    'HUNGARY':       'HU',
    'ROMANIA':       'RO',
    'BULGARIA':      'BG',
    'CROATIA':       'HR',
    'SERBIA':        'RS',
    'UKRAINE':       'UA',
    'RUSSIA':        'RU',
    'LATVIA':        'LV',
    'LITHUANIA':     'LT',
    'ICELAND':       'IS',
    'LUXEMBOURG':    'LU',
    'MALTA':         'MT',
    'CYPRUS':        'CY',
    'NEW CALEDONIA': 'NC',
    'PERU':          'PE',
    'ECUADOR':       'EC',
    'VENEZUELA':     'VE',
    'COSTA RICA':    'CR',
    'PANAMA':        'PA',
    'GUATEMALA':     'GT',
    'DOMINICAN REPUBLIC': 'DO',
    'JAMAICA':       'JM',
    'TRINIDAD AND TOBAGO': 'TT',
}


# Google Drive — Sync Output folder
DRIVE_SYNC_FOLDER_ID = '1BQ53mlYzkl3N5wz6AiTZm1kDpibempJy'
# drive scope (not drive.file) is required so the service account can access
# files in Shared Drives it has been granted Content Manager on.
DRIVE_SCOPES         = ['https://www.googleapis.com/auth/drive']

# Fields compared between ICF Global and GlueUp.
# Each entry: (icf_column_name, glueup_field_path, label_for_report)
# glueup_field_path uses dot notation for nested fields; 'properties.X' for
# custom fields.
# Fields compared between ICF Global and GlueUp.
# Each entry: (icf_column_name, glueup_field_path, label_for_report)
#
# GlueUp response structure:
#   { "membership": { "id":..., "startDate":..., "endDate":... },
#     "individualMember": {
#       "givenName":..., "familyName":...,
#       "emailAddress": {"value":...},
#       "address": {"cityName":..., "zipCode":...},
#       "properties": { "icfmemberid":..., "icfcredential": {"code":...}, ... }
#     }
#   }
# Paths are relative to the top-level record dict passed to _get_nested,
# which first navigates into individualMember or membership as appropriate.
#
# State/Province and Country/Region are stored in GlueUp as of the April 2026
# import template update and are included in comparison.
# GlueUp field paths for address sub-fields (stateName, countryName) are
# approximate — verify against live API response if comparison results look wrong.
COMPARISON_FIELDS = [
    ('First_Name',                    'individualMember.givenName',                         'First Name'),
    ('Last_Name',                     'individualMember.familyName',                        'Last Name'),
    ('Email',                         'individualMember.emailAddress.value',                'Email'),
    ('Phone',                         'individualMember.phone.value',                       'Phone'),
    ('City',                          'individualMember.address.cityName',                  'City'),
    ('Zip',                           'individualMember.address.zipCode',                   'Zip'),
    # State/Province is intentionally excluded from comparison. GlueUp stores
    # inconsistent values (e.g. 'WA' vs 'Washington') depending on import history,
    # and the field is not in the GlueUp import template — it cannot drive a
    # meaningful CHANGED classification. Captured in comparison report for reference.
    # ('State', 'individualMember.address.province', 'State/Province'),
    ('Country',                       'individualMember.address.country',                'Country/Region'),
    ('Membership_Expiration_Date',    'individualMember.properties.icfglobalmembershipenddate',   'Expiration Date'),
    # Creation Date (Chapter_Start_Date) is intentionally excluded from comparison —
    # the SOAP API does not reliably return this field, and it never changes once set.
    # It is still written to the import file via icf('Chapter_Start_Date') directly.
    ('Flagship_Credential',           'individualMember.properties.icfcredential',          'ICF Credential'),
    ('Credential_Award_Date',         'individualMember.properties.icfcredentialawarddate',      'ICF Credential Award Date'),
    ('Credential_Expire_Date',        'individualMember.properties.icfcredentialexpiredate',     'ICF Credential Expire Date'),
    ('ACTC_Credential',               'individualMember.properties.icfteamcoachingcredential',   'TC Credential'),
    ('ACTC_Credential_Award_Date',    'individualMember.properties.icfteamcoachingcredentialaward',  'TC Award Date'),
    ('ACTC_Credential_Expire_Date',   'individualMember.properties.icfteamcoachingcredentialexpir',  'TC Expire Date'),
    ('Auto_Renewal',                  'individualMember.properties.icfglobalautorenewal',   'Auto Renewal'),
]

# Columns in the GlueUp bulk import template (in order).
# These must exactly match the template column headers.
IMPORT_TEMPLATE_COLUMNS = [
    'Membership Start Date',
    'Membership End Date',
    'Currency',
    'First Name',
    'Last Name',
    'Email',
    'Phone',
    'Postal Code/Zip Code',
    'Address',
    'City',
    'State/Province',
    'Country/Region',
    'Volunteer Role',
    'Coach Industry',
    'Coach Specialty',
    'Findable',
    'Directory Listing Text',
    'ICF Credential',
    'ICF Credential Award Date',
    'ICF Credential Expire Date',
    'ICF Global Auto Renewal',
    'ICF Global Member ID',
    'ICF Global Membership End Date',
    'ICF Global Membership Start Date',
    'ICF Global Membership Type',
    'ICF Team Coaching Credential',
    'ICF Team Coaching Credential Award Date',
    'ICF Team Coaching Credential Expire Date',
    'Has Email',
    'ICF Global Import Date',
]

# Styling
COLOR_NEW       = 'C6EFCE'   # green
COLOR_CHANGED   = 'FFEB9C'   # yellow
COLOR_SAME      = 'FFFFFF'   # white
COLOR_HEADER    = '4472C4'   # blue
COLOR_DUPLICATE = 'FFC7CE'   # red


# ──────────────────────────────────────────────────────────────────────────────
# SECRET MANAGER (Cloud Run only)
# ──────────────────────────────────────────────────────────────────────────────

def _load_secrets_from_secret_manager() -> None:
    """
    Fetch GLUEUP_SK and GLUEUP_MD5_PW from GCP Secret Manager and override
    the module-level globals. Only runs when GLUEUP_USE_SECRET_MANAGER=1 is
    set in the environment (i.e. when running on Cloud Run).

    The GCP project ID is read from GOOGLE_CLOUD_PROJECT, which Cloud Run
    sets automatically. Secret names must match what was created in Secret
    Manager: 'glueup-private-key' and 'glueup-md5-password'.
    """
    global GLUEUP_SK, GLUEUP_MD5_PW

    if os.environ.get('GLUEUP_USE_SECRET_MANAGER') != '1':
        return

    try:
        from google.cloud import secretmanager
    except ImportError:
        print('WARNING: google-cloud-secret-manager not installed. '
              'Using hardcoded credential fallbacks.')
        return

    project_id = os.environ.get('GOOGLE_CLOUD_PROJECT')
    if not project_id:
        print('WARNING: GOOGLE_CLOUD_PROJECT not set. '
              'Using hardcoded credential fallbacks.')
        return

    client = secretmanager.SecretManagerServiceClient()

    def _get_secret(secret_id: str) -> str:
        name = f'projects/{project_id}/secrets/{secret_id}/versions/latest'
        response = client.access_secret_version(request={'name': name})
        return response.payload.data.decode('utf-8').strip()

    print('Loading secrets from GCP Secret Manager...')
    GLUEUP_SK     = _get_secret('glueup-private-key')
    GLUEUP_MD5_PW = _get_secret('glueup-md5-password')
    print('  ✓ Secrets loaded.')


# Call at import time so globals are set before any function uses them.
_load_secrets_from_secret_manager()


# ──────────────────────────────────────────────────────────────────────────────
# AUTH HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def _make_a_header(method: str = 'POST') -> str:
    """Compute the GlueUp HMAC-SHA256 'a' header."""
    ts = str(int(time.time() * 1000))
    msg = method + GLUEUP_PK + '1.0' + ts
    digest = hmac.new(GLUEUP_SK.encode(), msg.encode(), hashlib.sha256).hexdigest()
    return f'v=1.0;k={GLUEUP_PK};ts={ts};d={digest}'


def _user_agent() -> str:
    return (
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    )


def get_token() -> str:
    """Authenticate to GlueUp and return a session token."""
    print('Authenticating to GlueUp...')
    url = f'{GLUEUP_BASE_URL}/user/session'
    payload = {
        'email':      {'value': GLUEUP_EMAIL},
        'passphrase': {'value': GLUEUP_MD5_PW},
    }
    headers = {
        'Content-Type': 'application/json',
        'a':            _make_a_header('POST'),
        'User-Agent':   _user_agent(),
    }
    resp = requests.post(url, json=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    token = data.get('value', {}).get('token')
    if not token:
        raise RuntimeError(f'No token in login response: {json.dumps(data, indent=2)}')
    print('  ✓ Token obtained.')
    return token


# ──────────────────────────────────────────────────────────────────────────────
# GLUEUP DATA PULL
# ──────────────────────────────────────────────────────────────────────────────

def _glueup_headers(token: str) -> dict:
    return {
        'Content-Type':         'application/json',
        'a':                    _make_a_header('POST'),
        'token':                token,
        'requestOrganizationId': GLUEUP_ORG_ID,
        'User-Agent':           _user_agent(),
    }


def fetch_all_glueup_members(token: str) -> list[dict]:
    """
    Pull all active individual members from GlueUp, paginating as needed.
    Returns a list of raw member dicts from the API.
    """
    url = f'{GLUEUP_BASE_URL}/membershipDirectory/members'
    page_size = 100
    offset = 0
    all_members = []

    print('Fetching members from GlueUp...')
    while True:
        payload = {
            'filter': [
                {
                    'projection': 'membershipType',
                    'operator':   'eq',
                    'values':     [GLUEUP_MEMBERSHIP_TYPE_ID],
                }
            ],
            'offset': offset,
            'limit':  page_size,
        }
        resp = requests.post(
            url,
            json=payload,
            headers=_glueup_headers(token),
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()

        batch = data.get('value', [])
        all_members.extend(batch)

        total = data.get('metadata', {}).get('pagination', {}).get('total', 0)
        print(f'  Fetched {len(all_members)} / {total}')

        if len(all_members) >= total or not batch:
            break
        offset += page_size

    print(f'  ✓ Total GlueUp records fetched: {len(all_members)}')
    return all_members


def index_glueup_by_member_id(glueup_members: list[dict]) -> dict[str, list[dict]]:
    """
    Build a dict: icfmemberid -> list of GlueUp member records.
    A list is used because duplicates (same ID, multiple records) are possible
    and must be detected.

    The API response wraps each record as:
        { "membership": {...}, "individualMember": { "properties": {...}, ... } }
    so we navigate into individualMember.properties to find icfmemberid.
    """
    index: dict[str, list[dict]] = {}
    for m in glueup_members:
        individual = m.get('individualMember', {})
        props = individual.get('properties', {})
        icf_id = str(props.get('icfmemberid', '') or '').strip()
        if icf_id:
            index.setdefault(icf_id, []).append(m)
    return index


# ──────────────────────────────────────────────────────────────────────────────
# ICF GLOBAL FILE READER
# ──────────────────────────────────────────────────────────────────────────────

def find_latest_icf_file() -> str:
    """Find the most recently modified activemembers CSV in the CWD.
    Skips lock files (prefixed with ~$).
    """
    candidates = [
        f for f in
        glob.glob('*activemembers*.csv') + glob.glob('*ActiveMembers*.csv')
        if not os.path.basename(f).startswith('~$')
    ]
    if not candidates:
        raise FileNotFoundError(
            'No ICF Global member export CSV found in current directory.\n'
            'Pass the file path as an argument: python ICFGlueUpSync.py <file.csv>\n'
            'Expected filename pattern: activemembers_YYYYMMDD.csv'
        )
    return max(candidates, key=os.path.getmtime)


def fmt_date(val) -> str:
    """Normalise a date value to YYYY-MM-DD string, or empty string.
    Handles:
      - datetime / date objects
      - YYYY-MM-DD strings (pass-through)
      - MM/DD/YYYY strings written by the Make scenario
    """
    if val is None or val == '':
        return ''
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    if not s:
        return ''
    # MM/DD/YYYY → YYYY-MM-DD  (format written by Make Google Sheets module)
    if len(s) == 10 and s[2] == '/' and s[5] == '/':
        try:
            return datetime.strptime(s, '%m/%d/%Y').strftime('%Y-%m-%d')
        except ValueError:
            pass
    return s




def _map_country(raw: str, member_id: str = '') -> str:
    """
    Convert a country name or code to the 2-letter ISO code GlueUp requires.
    If the value is already a 2-letter code, return it uppercased.
    If it maps via COUNTRY_CODES, return the mapped code.
    Otherwise, print a warning and return the raw value so the problem
    is visible in the import file rather than silently wrong.
    """
    val = (raw or '').strip()
    if not val:
        return ''
    upper = val.upper()
    # Already a 2-letter code
    if len(upper) == 2:
        return upper
    # Look up full name
    mapped = COUNTRY_CODES.get(upper)
    if mapped:
        return mapped
    # Unknown — warn and pass through so the admin can see the issue
    print(f'  ⚠  UNMAPPED COUNTRY: Member {member_id} has country "{val}" — '
          f'not in COUNTRY_CODES. Add it to the table. Passing through as-is.')
    return val


def _derive_start_date(expiry_date_str: str) -> str:
    """
    Derive a membership start date from an expiration date by subtracting one
    year. Used as a fallback when Chapter_Start_Date is blank (the SOAP API
    does not reliably return it).
    Returns a YYYY-MM-DD string, or '' if the expiry date cannot be parsed.
    """
    s = (expiry_date_str or '').strip()
    if not s:
        return ''
    try:
        dt = datetime.strptime(s, '%Y-%m-%d')
        # Subtract one year, handling Feb-29 edge case
        try:
            start = dt.replace(year=dt.year - 1)
        except ValueError:
            # Feb 29 on a leap year — use Feb 28
            start = dt.replace(year=dt.year - 1, day=28)
        return start.strftime('%Y-%m-%d')
    except ValueError:
        return ''


def read_icf_file(path: str) -> list[dict]:
    """
    Read the CSV produced by the Make "Get Active Members" scenario.
    Row 1 = column headers (no skip rows).
    Returns a list of dicts keyed by the canonical column names used
    throughout this script (see HEADER_ALIASES below).

    The Make scenario CSV may use a mix of old manual-export header names
    and new API-derived names. HEADER_ALIASES normalises either form to the
    single canonical name the script expects, so both formats are accepted
    without code changes.

    Canonical names this function guarantees in every returned dict:
        Member_ID, Status, First_Name, Last_Name, Role, Email, Phone,
        City, State, Zip, Country, Chapter_Start_Date, Membership_Join_Date,
        Membership_Expiration_Date, Rejoin, Auto_Renewal, Flagship_Credential,
        Credential_Award_Date, Credential_Expire_Date, ACTC_Credential,
        ACTC_Credential_Award_Date, ACTC_Credential_Expire_Date
    """

    # Maps any known alias → canonical name.
    # Keys are stripped/lowercased for case-insensitive matching.
    HEADER_ALIASES = {
        # Member ID
        'member id':                        'Member_ID',
        'member_id':                        'Member_ID',
        # Names
        'first name':                       'First_Name',
        'first_name':                       'First_Name',
        'last name':                        'Last_Name',
        'last_name':                        'Last_Name',
        # Dates
        'expiration date':                  'Membership_Expiration_Date',
        'membership_expiration_date':       'Membership_Expiration_Date',
        # NOTE: 'creation date' intentionally NOT aliased here.
        # The Make CSV has both a real 'Chapter_Start_Date' column (col L)
        # AND a legacy 'Creation Date' column appended at the end (col V) which
        # is blank. Aliasing 'creation date' → 'Chapter_Start_Date' caused the
        # blank legacy column to overwrite the real value. The real column maps
        # correctly via 'chapter_start_date' below.
        'chapter_start_date':               'Chapter_Start_Date',
        'membership_join_date':             'Membership_Join_Date',
        'credential award date':            'Credential_Award_Date',
        'credential_award_date':            'Credential_Award_Date',
        'credential expire date':           'Credential_Expire_Date',
        'credential_expire_date':           'Credential_Expire_Date',
        'actc_credential_award_date':       'ACTC_Credential_Award_Date',
        'actc_credential_expire_date':      'ACTC_Credential_Expire_Date',
        # Credentials
        'credential':                       'Flagship_Credential',
        'flagship_credential':              'Flagship_Credential',
        'actc_credential':                  'ACTC_Credential',
        # Other
        'auto renewal':                     'Auto_Renewal',
        'auto_renewal':                     'Auto_Renewal',
        'rejoin':                           'Rejoin',
        # Pass-through columns (already canonical, listed for documentation)
        'status':   'Status',
        'role':     'Role',
        'email':    'Email',
        'phone':    'Phone',
        'city':     'City',
        'state':    'State',
        'zip':      'Zip',
        'country':  'Country',
    }

    # Date columns (canonical names) written by Make in MM/DD/YYYY format
    date_cols = {
        'Chapter_Start_Date', 'Membership_Join_Date', 'Membership_Expiration_Date',
        'Credential_Award_Date', 'Credential_Expire_Date',
        'ACTC_Credential_Award_Date', 'ACTC_Credential_Expire_Date',
    }

    # Valid country values: all keys (full names) and values (2-letter codes)
    # from COUNTRY_CODES. Anything else in the Country field suggests a
    # column-shifted row (e.g. Estonian district pushing columns right).
    KNOWN_COUNTRIES = set(COUNTRY_CODES.keys()) | set(COUNTRY_CODES.values())

    print(f'Reading ICF Global file: {path}')

    records = []
    nonstandard = []  # rows with column-shift or other structural problems
    with open(path, newline='', encoding='utf-8-sig') as fh:
        reader = csv.DictReader(fh)

        # Report which headers were found and how they were mapped
        raw_headers = reader.fieldnames or []
        mapped, unmapped = [], []
        for h in raw_headers:
            canonical = HEADER_ALIASES.get(h.strip().lower())
            if canonical:
                mapped.append(f'{h!r} → {canonical!r}')
            else:
                unmapped.append(repr(h))
        print(f'  CSV headers recognised: {len(mapped)}, unrecognised (kept as-is): {len(unmapped)}')
        if unmapped:
            print(f'  Unrecognised headers (passed through): {", ".join(unmapped)}')

        for row in reader:
            # Skip completely blank rows
            if not any(v.strip() for v in row.values()):
                continue

            # Remap headers to canonical names; unknown headers pass through
            rec = {}
            for k, v in row.items():
                canonical = HEADER_ALIASES.get((k or '').strip().lower(), k)
                rec[canonical] = (v.strip() if v is not None else '')

            # Normalise dates to YYYY-MM-DD
            for col in date_cols:
                rec[col] = fmt_date(rec.get(col, ''))

            # Normalise Member_ID: strip trailing .0 from numeric strings
            # (Google Sheets may export integers as "12345.0")
            mid = rec.get('Member_ID', '').strip()
            if mid.endswith('.0'):
                mid = mid[:-2]
            rec['Member_ID'] = mid

            # Zip: strip trailing .0 for the same reason
            zip_val = rec.get('Zip', '').strip()
            if zip_val.endswith('.0'):
                zip_val = zip_val[:-2]
            rec['Zip'] = zip_val

            # Detect column-shifted rows: if Country doesn't look like a known
            # country value, the address likely had an extra field (e.g. Estonian
            # district) that pushed all subsequent columns one position to the right.
            country_val = rec.get('Country', '').strip().upper()
            if country_val and country_val not in KNOWN_COUNTRIES:
                nonstandard.append({
                    'Member_ID':  rec.get('Member_ID', ''),
                    'First_Name': rec.get('First_Name', ''),
                    'Last_Name':  rec.get('Last_Name', ''),
                    'Country':    rec.get('Country', ''),
                    'Reason':     (
                        f'Unrecognised country value "{rec.get("Country", "")}" — '
                        f'address may have an extra field causing column misalignment. '
                        f'Review raw ICF Global record and import manually if needed.'
                    ),
                })
                print(f'  ⚠  NON-STANDARD: Member {rec.get("Member_ID", "?")} '
                      f'({rec.get("First_Name", "")} {rec.get("Last_Name", "")}) — '
                      f'unrecognised country "{rec.get("Country", "")}", skipping.')
                continue

            records.append(rec)

    if nonstandard:
        print(f'  ⚠  {len(nonstandard)} non-standard record(s) skipped — '
              f'see comparison report Non-Standard sheet.')
    print(f'  ✓ {len(records)} ICF Global records loaded.')
    # Attach nonstandard list to the return value via a module-level variable
    # so write_comparison_report can access it without changing all call sites.
    read_icf_file._nonstandard = nonstandard
    return records


# ──────────────────────────────────────────────────────────────────────────────
# FIELD EXTRACTION FROM GLUEUP RECORD
# ──────────────────────────────────────────────────────────────────────────────

def _get_nested(obj: dict, path: str) -> str:
    """
    Walk a dot-separated path into a nested dict and return the value as a
    normalised string. Returns '' for missing / None.

    Handles three GlueUp value shapes:
      - Plain string / number
      - Timestamp in milliseconds  → YYYY-MM-DD
      - Coded field {code, title}  → the code string (lowercased)
      - Date string MM/DD/YYYY     → YYYY-MM-DD
    """
    parts = path.split('.')
    cur = obj
    for p in parts:
        if not isinstance(cur, dict):
            return ''
        cur = cur.get(p)
    if cur is None:
        return ''
    # GlueUp coded fields: {"code": "pcc", "title": {...}}
    if isinstance(cur, dict) and 'code' in cur:
        code = str(cur['code']).strip().lower()
        return '' if code == 'none' else code
    # GlueUp timestamps in milliseconds
    if isinstance(cur, (int, float)) and cur > 1_000_000_000_000:
        dt = datetime.fromtimestamp(cur / 1000)
        return dt.strftime('%Y-%m-%d')
    s = str(cur).strip()
    # GlueUp date strings in MM/DD/YYYY format → normalise to YYYY-MM-DD
    if len(s) == 10 and s[2] == '/' and s[5] == '/':
        try:
            return datetime.strptime(s, '%m/%d/%Y').strftime('%Y-%m-%d')
        except ValueError:
            pass
    return s


def extract_glueup_fields(member: dict) -> dict[str, str]:
    """
    Flatten a GlueUp member record into a simple {field_label: value} dict
    using the COMPARISON_FIELDS mapping.
    """
    result = {}
    for icf_col, glueup_path, label in COMPARISON_FIELDS:
        result[label] = _get_nested(member, glueup_path)
    # Also pull the raw icfmemberid for reference
    individual = member.get('individualMember', {})
    result['ICF Global Member ID'] = str(
        individual.get('properties', {}).get('icfmemberid', '') or ''
    ).strip()
    result['GlueUp Membership ID'] = str(
        member.get('membership', {}).get('id', '') or ''
    ).strip()
    result['GlueUp Member Record ID'] = str(individual.get('id', '') or '').strip()
    result['ICF Global Import Date'] = str(
        individual.get('properties', {}).get('icfimportdate', '') or ''
    ).strip()
    return result


def extract_icf_fields(rec: dict) -> dict[str, str]:
    """
    Normalise an ICF Global record row into the same {label: value} shape
    as extract_glueup_fields, for easy side-by-side comparison.
    All string values are stripped of leading/trailing whitespace.

    Country is converted from full name (e.g. 'UNITED STATES') to ISO code
    (e.g. 'US') to match how GlueUp stores it.
    State is normalized to a 2-letter abbreviation: if the SOAP API returns a
    full name (e.g. 'Washington'), it is looked up in US_STATE_ABBREVS. If it
    is already an abbreviation (e.g. 'WA'), it is uppercased and kept as-is.
    """
    result = {}
    for icf_col, _, label in COMPARISON_FIELDS:
        val = rec.get(icf_col)
        str_val = '' if val is None else str(val).strip()
        # Normalize Country to ISO code to match GlueUp's coded field storage
        if label == 'Country/Region':
            str_val = COUNTRY_CODES.get(str_val.upper(), str_val)
        # State/Province is not compared (excluded from COMPARISON_FIELDS) —
        # no normalization needed here.
        result[label] = str_val
    result['ICF Global Member ID'] = rec.get('Member_ID', '')
    return result


def _normalise(val: str) -> str:
    """Normalise a field value for comparison: strip whitespace and lowercase."""
    return (val or '').strip().lower()


def _normalise_phone(val: str) -> str:
    """
    Normalise a phone number for comparison by stripping all non-digit characters
    and returning the last 10 digits (the local number, ignoring country code).

    This ensures format differences are not treated as changes:
        '512.773.5096'   → '5127735096'
        '+1 5127735096'  → '5127735096'
        '(512) 773-5096' → '5127735096'
    Returns '' if the value has fewer than 10 digits (i.e. not a real number).
    """
    import re as _re
    if not val:
        return ''
    digits = _re.sub(r'\D', '', str(val))
    return digits[-10:] if len(digits) >= 10 else ''


def _date_diff_days(val_a: str, val_b: str) -> int | None:
    """
    Return the absolute difference in days between two YYYY-MM-DD date strings.
    Returns None if either value is blank or cannot be parsed.
    """
    if not val_a or not val_b:
        return None
    try:
        a = datetime.strptime(val_a.strip(), '%Y-%m-%d')
        b = datetime.strptime(val_b.strip(), '%Y-%m-%d')
        return abs((a - b).days)
    except ValueError:
        return None


# ──────────────────────────────────────────────────────────────────────────────
# COMPARISON ENGINE
# ──────────────────────────────────────────────────────────────────────────────

def compare_records(icf_records: list[dict],
                    glueup_index: dict[str, list[dict]]) -> list[dict]:
    """
    For each ICF Global record, determine status and which fields differ.

    Returns a list of comparison result dicts:
        member_id, status, icf_fields, glueup_fields, changed_fields
    """
    results = []
    skipped_shadow = []
    for rec in icf_records:
        mid = rec.get('Member_ID', '').strip()
        if not mid:
            print(f'  WARNING: skipping row with blank Member ID (row data: {rec})')
            continue

        # Skip shadow-email members with no expiration date — they have no
        # useful data to import. Named warning printed for each.
        raw_email   = (rec.get('Email') or '').strip()
        expiry_date = (rec.get('Membership_Expiration_Date') or '').strip()
        if not raw_email and not expiry_date:
            fname = (rec.get('First_Name') or '').strip()
            lname = (rec.get('Last_Name') or '').strip()
            skipped_shadow.append((mid, fname, lname))
            continue

        icf_fields = extract_icf_fields(rec)

        glueup_matches = glueup_index.get(mid, [])

        if not glueup_matches:
            results.append({
                'member_id':     mid,
                'status':        'NEW',
                'icf_fields':    icf_fields,
                'glueup_fields': {},
                'changed_fields': [],
                'glueup_record': None,
            })
            continue

        # When multiple GlueUp records exist for the same icfmemberid (i.e.
        # duplicates from a previous import), always compare against the NEWEST
        # one. Primary sort: expiration date. Tiebreaker: import date (most
        # recently imported record wins). Both are YYYY-MM-DD strings so
        # lexicographic sort is correct.
        def _newest_record_key(m):
            ind = m.get('individualMember', {})
            props = ind.get('properties', {})
            exp_date    = str(props.get('icfglobalmembershipenddate', '') or '')
            import_date = str(props.get('icfimportdate', '') or '')
            # Final tiebreaker: highest individualMember.id = most recently
            # created record in GlueUp. Padded to 10 digits for correct
            # lexicographic comparison.
            member_id   = str(ind.get('id', 0) or 0).zfill(10)
            return (exp_date, import_date, member_id)
        glueup_rec = max(glueup_matches, key=_newest_record_key)
        glueup_fields = extract_glueup_fields(glueup_rec)

        # Date fields where a small difference is tolerated (timezone drift etc.)
        # Key = comparison label, value = max tolerated difference in days.
        DATE_TOLERANCE: dict[str, int] = {}

        # If ICF Global has no credential, skip the credential field and its
        # associated date fields. The SOAP API drops credentials that have
        # expired, so a blank from ICF Global does not mean the credential
        # should be removed from GlueUp — we preserve whatever GlueUp has
        # for historical reporting and admin awareness.
        icf_credential_blank     = _normalise(icf_fields.get('ICF Credential', '')) == ''
        icf_tc_credential_blank  = _normalise(icf_fields.get('TC Credential', '')) == ''
        SKIP_IF_CREDENTIAL_BLANK = {
            'ICF Credential':             icf_credential_blank,
            'ICF Credential Award Date':  icf_credential_blank,
            'ICF Credential Expire Date': icf_credential_blank,
            'TC Credential':              icf_tc_credential_blank,
            'TC Award Date':              icf_tc_credential_blank,
            'TC Expire Date':             icf_tc_credential_blank,
        }

        changed = []
        for _, _, label in COMPARISON_FIELDS:
            icf_val  = icf_fields.get(label, '')
            glu_val  = glueup_fields.get(label, '')

            # Phone: skip comparison if either side has no phone.
            # We cannot update a phone via import if ICF Global has none,
            # and we don't want to blank out a good GlueUp phone number.
            # Also skip if GlueUp has no phone — cannot update via import.
            if label == 'Phone':
                icf_norm = _normalise_phone(icf_val)
                glu_norm = _normalise_phone(glu_val)
                if icf_norm == '' or glu_norm == '':
                    continue
                if icf_norm == glu_norm:
                    continue
                changed.append(label)
                continue

            icf_norm = _normalise(icf_val)
            glu_norm = _normalise(glu_val)
            # Treat blank and None as equivalent
            if icf_norm == glu_norm or (icf_norm == '' and glu_norm == ''):
                continue
            # Skip credential date fields when the parent credential is blank
            # in ICF Global — we preserve GlueUp dates for historical reporting.
            if SKIP_IF_CREDENTIAL_BLANK.get(label, False):
                continue
            # For date fields with a tolerance, ignore small differences
            if label in DATE_TOLERANCE:
                diff = _date_diff_days(icf_fields.get(label, ''), glueup_fields.get(label, ''))
                if diff is not None and diff <= DATE_TOLERANCE[label]:
                    continue
            changed.append(label)

        status = 'CHANGED' if changed else 'SAME'
        results.append({
            'member_id':      mid,
            'status':         status,
            'icf_fields':     icf_fields,
            'glueup_fields':  glueup_fields,
            'changed_fields': changed,
            'glueup_record':  glueup_rec,
        })

    if skipped_shadow:
        print(f'\n  ⚠  {len(skipped_shadow)} shadow-email member(s) with no expiration date '
              f'skipped (no useful import data):')
        for mid, fn, ln in skipped_shadow:
            print(f'       icfmemberid {mid} — {fn} {ln}')
        print()

    return results


def _classify_duplicate_group(members: list[dict]) -> str:
    """
    Given a list of GlueUp records that share the same icfmemberid, determine
    whether human action is required.

    Rules:
      - If native membership end dates are identical across all records → 'DELETE OLD'
        (GlueUp cannot auto-expire when dates match; manual deletion required)
      - If native end dates differ AND the only differing custom field is
        icfglobalmembershipenddate / Import Date → 'WAIT — will auto-expire'
      - If any other field differs → 'DELETE OLD — manual action required'

    ICF Global Import Date is always excluded from the comparison because it
    will differ by definition between an old and a newly-imported record.

    NOTE: We use the native membership.endDate timestamp (not the custom
    icfglobalmembershipenddate field) to determine whether dates differ,
    because a renewal import updates the custom field on both records but
    GlueUp's native end date correctly reflects the two different expiry dates.
    """
    EXCLUDE_FROM_CLASSIFY = {'Expiration Date', 'ICF Global Import Date'}

    field_sets = [extract_glueup_fields(m) for m in members]

    # Use native membership.endDate (ms timestamp → YYYY-MM-DD) for expiry
    # comparison — this is what GlueUp actually uses to auto-expire records.
    def _native_end_date(m):
        end_ms = m.get('membership', {}).get('endDate')
        if end_ms and isinstance(end_ms, (int, float)) and end_ms > 1_000_000_000_000:
            return datetime.fromtimestamp(end_ms / 1000).strftime('%Y-%m-%d')
        return ''

    native_end_dates = [_native_end_date(m) for m in members]
    all_same_expiry = len(set(native_end_dates)) == 1

    # If native end dates are all the same, WAIT is meaningless — must delete
    if all_same_expiry:
        return 'DELETE OLD — manual action required'

    # Expiration dates differ — check if that's the ONLY difference
    for i in range(len(field_sets)):
        for j in range(i + 1, len(field_sets)):
            a, b = field_sets[i], field_sets[j]
            for _, _, label in COMPARISON_FIELDS:
                if label in EXCLUDE_FROM_CLASSIFY:
                    continue
                if _normalise(a.get(label, '')) != _normalise(b.get(label, '')):
                    return 'DELETE OLD — manual action required'

    return 'WAIT — will auto-expire'


def find_duplicates(glueup_index: dict[str, list[dict]]) -> list[dict]:
    """
    Return a flat list of GlueUp member records where a given icfmemberid
    appears on more than one membership record, with an Action column
    indicating whether manual deletion is required or GlueUp will self-clean.

    Records within each duplicate group are sorted oldest expiration date first,
    so the record to delete (when action is required) is always at the top of
    each group.
    """
    dupes = []
    for mid, members in glueup_index.items():
        if len(members) <= 1:
            continue

        action = _classify_duplicate_group(members)

        # Sort by expiration date then import date ascending — oldest/earliest first.
        # This ensures the record to delete is listed first in each group.
        def _dup_sort_key(m):
            ind = m.get('individualMember', {})
            props = ind.get('properties', {})
            exp_date    = str(props.get('icfglobalmembershipenddate', '') or '')
            import_date = str(props.get('icfimportdate', '') or '')
            return (exp_date, import_date)

        sorted_members = sorted(members, key=_dup_sort_key)

        for idx, m in enumerate(sorted_members):
            gf = extract_glueup_fields(m)
            gf['ICF Global Member ID'] = mid
            gf['Duplicate Count']      = str(len(members))
            gf['Action']               = action
            # For DELETE OLD groups: oldest = delete, newest = keep.
            # For WAIT groups: oldest = will auto-expire, newest = keep.
            if action.startswith('DELETE'):
                gf['Record Order'] = 'DELETE THIS' if idx == 0 else 'KEEP'
            else:
                gf['Record Order'] = 'WILL AUTO-EXPIRE' if idx == 0 else 'KEEP'
            dupes.append(gf)

    # Sort output: DELETE OLD groups first, then WAIT groups; within each
    # group keep the icfmemberid together
    dupes.sort(key=lambda d: (
        0 if d['Action'].startswith('DELETE') else 1,
        d['ICF Global Member ID'],
        d['Record Order'],
    ))
    return dupes


def find_dropped_members(glueup_index: dict[str, list[dict]],
                          icf_ids: set[str]) -> list[dict]:
    """
    Find GlueUp members who have an icfmemberid but are no longer present
    in the current ICF Global export. These members may have cancelled,
    transferred chapters, or been removed from the source of truth.

    Their GlueUp membership should be reviewed and cancelled manually.
    They should be kept as contacts but their membership record removed.

    Returns a list of dicts, one per GlueUp record, sorted by last name.
    """
    dropped = []
    for mid, members in glueup_index.items():
        if mid in icf_ids:
            continue
        # All records for this member are dropped — report the newest one
        def _newest_key(m):
            ind = m.get('individualMember', {})
            props = ind.get('properties', {})
            exp_date    = str(props.get('icfglobalmembershipenddate', '') or '')
            import_date = str(props.get('icfimportdate', '') or '')
            member_id   = str(ind.get('id', 0) or 0).zfill(10)
            return (exp_date, import_date, member_id)
        rec = max(members, key=_newest_key)
        ind = rec.get('individualMember', {})
        props = ind.get('properties', {})
        dropped.append({
            'ICF Global Member ID':    mid,
            'GlueUp Membership ID':    str(rec.get('membership', {}).get('id', '') or ''),
            'GlueUp Member Record ID': str(ind.get('id', '') or ''),
            'First Name':              str(ind.get('givenName', '') or ''),
            'Last Name':               str(ind.get('familyName', '') or ''),
            'Email':                   str((ind.get('emailAddress') or {}).get('value', '') or ''),
            'Expiration Date':         str(props.get('icfglobalmembershipenddate', '') or ''),
            'ICF Credential':          _get_nested(rec, 'individualMember.properties.icfcredential'),
            'Import Date':             str(props.get('icfimportdate', '') or ''),
            'Duplicate Records':       str(len(members)),
        })
    dropped.sort(key=lambda d: (d['Last Name'].lower(), d['First Name'].lower()))
    return dropped


def write_dropped_file(dropped: list[dict], output_path: str):
    """
    Write the dropped members report.
    These are GlueUp members with an icfmemberid that no longer appears
    in the current ICF Global export.
    Action: keep as GlueUp contact but cancel/delete their membership.
    """
    COLOR_DROPPED = 'FCE4D6'  # light orange

    print(f'Writing dropped members report: {output_path}')
    wb = Workbook()
    ws = wb.active
    ws.title = 'Dropped Members'

    if not dropped:
        ws.append(['No dropped members found — all GlueUp members are in the ICF Global export.'])
        wb.save(output_path)
        print('  ✓ No dropped members found.')
        return

    headers = [
        'ICF Global Member ID',
        'GlueUp Membership ID',
        'GlueUp Member Record ID',
        'First Name',
        'Last Name',
        'Email',
        'Expiration Date',
        'ICF Credential',
        'Import Date',
        'Duplicate Records',
    ]
    ws.append(headers)
    _header_row_style(ws, 1, len(headers))

    drop_fill = PatternFill(fill_type='solid', fgColor=COLOR_DROPPED)
    for d in dropped:
        ws.append([d.get(h, '') for h in headers])
        row_num = ws.max_row
        for col in range(1, len(headers) + 1):
            ws.cell(row=row_num, column=col).fill = drop_fill

    ws.append([])
    ws.append(['INSTRUCTIONS'])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    ws.append([
        'These members are in GlueUp but no longer in the ICF Global export. '
        'They may have cancelled, transferred chapters, or been removed. '
        'Keep them as GlueUp contacts but cancel/delete their membership record. '
        'Use the GlueUp Membership ID to find and cancel the membership in Admin UI.'
    ])

    _auto_width(ws)
    ws.freeze_panes = 'A2'
    wb.save(output_path)
    print(f'  ✓ Dropped members report saved: {len(dropped)} member(s) '
          f'no longer in ICF Global export.')


# ──────────────────────────────────────────────────────────────────────────────
# SHADOW EMAIL LOGIC
# ──────────────────────────────────────────────────────────────────────────────

def effective_email(icf_rec: dict) -> tuple[str, bool]:
    """
    Return (email_to_use, has_real_email).
    If the ICF Global record has no email, generate a shadow email per spec:
        id_[GlobalMemberID]@members.icfwashingtonstate.org
    """
    raw_email = (icf_rec.get('Email') or '').strip()
    if raw_email:
        return raw_email, True
    mid = icf_rec.get('Member_ID', '').strip()
    shadow = f'id_{mid}@{SHADOW_EMAIL_DOMAIN}'
    print(f'  SHADOW EMAIL: Member {mid} has no email — using {shadow}')
    return shadow, False


# ──────────────────────────────────────────────────────────────────────────────
# OUTPUT FILE BUILDERS
# ──────────────────────────────────────────────────────────────────────────────

def _header_row_style(ws, row_num: int, num_cols: int, fill_color: str = COLOR_HEADER):
    fill = PatternFill(fill_type='solid', fgColor=fill_color)
    font = Font(bold=True, color='FFFFFF' if fill_color == COLOR_HEADER else '000000')
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=False)


def _auto_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)


def write_comparison_report(results: list[dict],
                             icf_records: list[dict],
                             output_path: str,
                             import_mode: str = 'all',
                             icf_records_by_id: dict | None = None):
    """
    Write the full comparison report XLS.
    Columns: Status | Member ID | Field Name | ICF Value | GlueUp Value | Match?
    One row per comparison field per member — makes it easy to filter.
    Also writes a Summary sheet.
    """
    print(f'Writing comparison report: {output_path}')
    if icf_records_by_id is None:
        icf_records_by_id = {r.get('Member_ID', ''): r for r in icf_records}
    wb = Workbook()

    # ── Summary sheet ──────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = 'Summary'

    new_count     = sum(1 for r in results if r['status'] == 'NEW')
    changed_count = sum(1 for r in results if r['status'] == 'CHANGED')
    same_count    = sum(1 for r in results if r['status'] == 'SAME')

    import_count = {
        'all':     new_count + changed_count,
        'new':     new_count,
        'changed': changed_count,
    }.get(import_mode, new_count + changed_count)
    import_mode_label = {
        'all':     'NEW + CHANGED',
        'new':     'NEW only',
        'changed': 'CHANGED only',
    }.get(import_mode, import_mode)

    summary_data = [
        ('Run Date',        datetime.now().strftime('%Y-%m-%d %H:%M')),
        ('ICF Global Records', len(results)),
        ('NEW (not in GlueUp)',       new_count),
        ('CHANGED (fields differ)',   changed_count),
        ('SAME (no changes)',         same_count),
        ('',                          ''),
        (f'Import file will contain ({import_mode_label})', import_count),
    ]
    for label, value in summary_data:
        ws_sum.append([label, value])

    _header_row_style(ws_sum, 1, 2)
    ws_sum.column_dimensions['A'].width = 30
    ws_sum.column_dimensions['B'].width = 20

    # ── Detail sheet ───────────────────────────────────────────────
    ws = wb.create_sheet('Comparison Detail')

    detail_headers = [
        'Status', 'Member ID', 'First Name', 'Last Name',
        'Phone', 'State', 'Country',
        'Field', 'ICF Global Value', 'GlueUp Value', 'Match?',
        'Normalised ICF', 'Normalised GlueUp',
    ]
    ws.append(detail_headers)
    _header_row_style(ws, 1, len(detail_headers))

    row_num = 2
    for r in results:
        mid    = r['member_id']
        status = r['status']
        fname  = r['icf_fields'].get('First Name', '')
        lname  = r['icf_fields'].get('Last Name', '')
        # Pull Phone/State/Country directly from the raw ICF record
        icf_rec = icf_records_by_id.get(mid, {})
        phone   = str(icf_rec.get('Phone') or '').strip()
        state   = str(icf_rec.get('State') or '').strip()
        country = str(icf_rec.get('Country') or '').strip()

        if status == 'NEW':
            # One summary row for new members — no GlueUp data to compare
            row_fill = PatternFill(fill_type='solid', fgColor=COLOR_NEW)
            ws.append([status, mid, fname, lname, phone, state, country, '(not in GlueUp)', '', '', ''])
            for col in range(1, len(detail_headers) + 1):
                ws.cell(row=row_num, column=col).fill = row_fill
            row_num += 1
        else:
            changed_set = set(r['changed_fields'])
            fill_color  = COLOR_CHANGED if status == 'CHANGED' else COLOR_SAME
            for _, _, label in COMPARISON_FIELDS:
                icf_val  = r['icf_fields'].get(label, '')
                glu_val  = r['glueup_fields'].get(label, '')
                match    = 'YES' if label not in changed_set else 'NO'
                ws.append([status, mid, fname, lname, phone, state, country,
                           label, icf_val, glu_val, match,
                           _normalise(icf_val), _normalise(glu_val)])
                row_fill = PatternFill(fill_type='solid', fgColor=fill_color)
                for col in range(1, len(detail_headers) + 1):
                    ws.cell(row=row_num, column=col).fill = row_fill
                row_num += 1

    _auto_width(ws)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    # ── Changed fields summary sheet ───────────────────────────────
    ws_ch = wb.create_sheet('Changed Records')
    ch_headers = ['Member ID', 'First Name', 'Last Name', 'Changed Fields']
    ws_ch.append(ch_headers)
    _header_row_style(ws_ch, 1, len(ch_headers))

    for r in results:
        if r['status'] == 'CHANGED':
            ws_ch.append([
                r['member_id'],
                r['icf_fields'].get('First Name', ''),
                r['icf_fields'].get('Last Name', ''),
                ', '.join(r['changed_fields']),
            ])

    _auto_width(ws_ch)

    # ── Non-Standard Records sheet ────────────────────────────────
    nonstandard = getattr(read_icf_file, '_nonstandard', [])
    if nonstandard:
        ws_ns = wb.create_sheet('Non-Standard Records')
        ns_headers = ['Member ID', 'First Name', 'Last Name', 'Country Value', 'Reason']
        ws_ns.append(ns_headers)
        _header_row_style(ws_ns, 1, len(ns_headers), fill_color='FF0000')
        warn_fill = PatternFill(fill_type='solid', fgColor='FCE4D6')
        for ns in nonstandard:
            ws_ns.append([
                ns.get('Member_ID', ''),
                ns.get('First_Name', ''),
                ns.get('Last_Name', ''),
                ns.get('Country', ''),
                ns.get('Reason', ''),
            ])
            row_num_ns = ws_ns.max_row
            for col in range(1, len(ns_headers) + 1):
                ws_ns.cell(row=row_num_ns, column=col).fill = warn_fill
        ws_ns.append([])
        ws_ns.append(['These records were EXCLUDED from the import file. Review the raw '
                      'ICF Global data and import manually if needed.'])
        _auto_width(ws_ns)

    wb.save(output_path)
    print(f'  ✓ Comparison report saved.')


def write_import_file(results: list[dict],
                      icf_records_by_id: dict[str, dict],
                      output_path: str,
                      import_mode: str = 'all'):
    """
    Write the GlueUp bulk import XLS.
    Column order matches the GlueUp import template exactly.

    import_mode:
        'all'     — include NEW and CHANGED records (default)
        'new'     — include NEW records only
        'changed' — include CHANGED records only
    """
    mode_label = {'all': 'NEW + CHANGED', 'new': 'NEW only', 'changed': 'CHANGED only'}
    print(f'Writing import file ({mode_label.get(import_mode, import_mode)}): {output_path}')
    today = datetime.now().strftime('%Y-%m-%d')
    wb = Workbook()
    ws = wb.active
    ws.title = 'Individual Members'

    ws.append(IMPORT_TEMPLATE_COLUMNS)
    _header_row_style(ws, 1, len(IMPORT_TEMPLATE_COLUMNS))

    include_statuses = {
        'all':     {'NEW', 'CHANGED'},
        'new':     {'NEW'},
        'changed': {'CHANGED'},
    }.get(import_mode, {'NEW', 'CHANGED'})

    import_count = 0
    for r in results:
        if r['status'] not in include_statuses:
            continue

        mid     = r['member_id']
        icf_rec = icf_records_by_id.get(mid, {})
        email, has_real_email = effective_email(icf_rec)

        def icf(col):
            return icf_rec.get(col) or ''

        row = {
            'Membership Start Date':               (
                # Use Chapter_Start_Date if available; fall back to one year
                # before the expiration date when the SOAP API omits it.
                fmt_date(icf('Chapter_Start_Date')) or _derive_start_date(icf('Membership_Expiration_Date'))
            ),
            'Membership End Date':                 fmt_date(icf('Membership_Expiration_Date')),
            'Currency':                            '',
            'First Name':                          (icf('First_Name') or '').strip(),
            'Last Name':                           (icf('Last_Name') or '').strip(),
            'Email':                               email,
            'Phone':                               (icf('Phone') or '').strip(),
            'Postal Code/Zip Code':                str(int(float(icf('Zip')))).strip() if icf('Zip') is not None and str(icf('Zip')).replace('.','',1).isdigit() else str(icf('Zip') or '').strip(),
            'Address':                             '',
            'City':                                (icf('City') or '').strip(),
            'State/Province':                      (icf('State') or '').strip(),
            'Country/Region':                      _map_country(icf('Country'), mid),
            'Volunteer Role':                      '',
            'Coach Industry':                      '',
            'Coach Specialty':                     '',
            'Findable':                            '',
            'Directory Listing Text':              '',
            'ICF Credential':                      (icf('Flagship_Credential') or '').strip().lower() or 'none',
            # Leave credential dates blank if ICF Global has no credential —
            # GlueUp will preserve its existing dates for historical reporting.
            'ICF Credential Award Date':           fmt_date(icf('Credential_Award_Date')) if (icf('Flagship_Credential') or '').strip() else '',
            'ICF Credential Expire Date':          fmt_date(icf('Credential_Expire_Date')) if (icf('Flagship_Credential') or '').strip() else '',
            'ICF Global Auto Renewal':             (icf('Auto_Renewal') or '').strip().lower(),
            'ICF Global Member ID':                mid,
            'ICF Global Membership End Date':      fmt_date(icf('Membership_Expiration_Date')),
            'ICF Global Membership Start Date':    (
                fmt_date(icf('Chapter_Start_Date')) or _derive_start_date(icf('Membership_Expiration_Date'))
            ),
            'ICF Global Membership Type':          'individual',
            'ICF Team Coaching Credential':        (icf('ACTC_Credential') or '').strip().lower() or 'none',
            'ICF Team Coaching Credential Award Date': fmt_date(icf('ACTC_Credential_Award_Date')) if (icf('ACTC_Credential') or '').strip() else '',
            'ICF Team Coaching Credential Expire Date': fmt_date(icf('ACTC_Credential_Expire_Date')) if (icf('ACTC_Credential') or '').strip() else '',
            'Has Email':                           'yes' if has_real_email else 'no',
            'ICF Global Import Date':              today,
        }

        row_values = [row[col] for col in IMPORT_TEMPLATE_COLUMNS]

        fill_color = COLOR_NEW if r['status'] == 'NEW' else COLOR_CHANGED
        ws.append(row_values)
        row_fill = PatternFill(fill_type='solid', fgColor=fill_color)
        row_num = ws.max_row
        for col in range(1, len(IMPORT_TEMPLATE_COLUMNS) + 1):
            ws.cell(row=row_num, column=col).fill = row_fill

        import_count += 1

    _auto_width(ws)
    ws.freeze_panes = 'A2'

    # Legend
    ws_leg = wb.create_sheet('Legend')
    ws_leg.append(['Color', 'Meaning'])
    _header_row_style(ws_leg, 1, 2)
    ws_leg.append(['Green', 'NEW — member not currently in GlueUp'])
    ws_leg.cell(row=2, column=1).fill = PatternFill(fill_type='solid', fgColor=COLOR_NEW)
    ws_leg.append(['Yellow', 'CHANGED — member exists but one or more fields differ'])
    ws_leg.cell(row=3, column=1).fill = PatternFill(fill_type='solid', fgColor=COLOR_CHANGED)
    _auto_width(ws_leg)

    wb.save(output_path)
    print(f'  ✓ Import file saved ({import_count} records: new + changed).')


def write_duplicate_file(duplicates: list[dict], output_path: str):
    """
    Write the duplicate membership report.

    Two row colors:
      Red    — DELETE OLD: fields other than Expiration Date differ;
               manual deletion required. OLDEST record is listed first —
               that is the one to delete.
      Orange — WAIT: only Expiration Date differs; GlueUp will auto-expire
               the older record. No action needed.

    Within each color group, records are sorted by ICF Global Member ID so
    duplicates for the same member are always adjacent.
    """
    COLOR_DELETE = 'FFC7CE'   # red   — action required
    COLOR_WAIT   = 'FFE0B2'   # orange — no action needed

    print(f'Writing duplicate report: {output_path}')
    wb = Workbook()
    ws = wb.active
    ws.title = 'Duplicate Memberships'

    if not duplicates:
        ws.append(['No duplicate memberships found.'])
        wb.save(output_path)
        print('  ✓ No duplicates found.')
        return

    headers = [
        'Action',
        'Record Order',
        'ICF Global Member ID',
        'Duplicate Count',
        'GlueUp Membership ID',
        'GlueUp Member Record ID',
        'First Name',
        'Last Name',
        'Email',
        'Expiration Date',
        'ICF Credential',
        'Import Date',
    ]
    ws.append(headers)
    _header_row_style(ws, 1, len(headers))

    delete_count = 0
    wait_count   = 0

    for d in duplicates:
        action = d.get('Action', '')
        is_delete = action.startswith('DELETE')
        fill_color = COLOR_DELETE if is_delete else COLOR_WAIT

        row = [
            action,
            d.get('Record Order', ''),
            d.get('ICF Global Member ID', ''),
            d.get('Duplicate Count', ''),
            d.get('GlueUp Membership ID', ''),
            d.get('GlueUp Member Record ID', ''),
            d.get('First Name', ''),
            d.get('Last Name', ''),
            d.get('Email', ''),
            d.get('Expiration Date', ''),
            d.get('ICF Credential', ''),
            d.get('ICF Global Import Date', ''),
        ]
        ws.append(row)
        row_fill = PatternFill(fill_type='solid', fgColor=fill_color)
        row_num = ws.max_row
        for col in range(1, len(headers) + 1):
            ws.cell(row=row_num, column=col).fill = row_fill

        if is_delete:
            delete_count += 1
        else:
            wait_count += 1

    # Instructions block at the bottom
    ws.append([])
    ws.append(['INSTRUCTIONS'])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    ws.append([
        'RED rows (DELETE OLD): Go to GlueUp Admin → Members, search by '
        'ICF Global Member ID, and delete the OLDEST record (listed first in '
        'each group). Keep the record with the most recent Expiration Date.'
    ])
    ws.append([
        'ORANGE rows (WAIT): Only the Expiration Date differs. GlueUp will '
        'automatically mark the older record as expired once its date passes. '
        'No manual action required.'
    ])

    _auto_width(ws)
    ws.freeze_panes = 'A2'
    wb.save(output_path)
    unique_ids    = len(set(d['ICF Global Member ID'] for d in duplicates))
    delete_groups = len(set(d['ICF Global Member ID'] for d in duplicates
                            if d['Action'].startswith('DELETE')))
    wait_groups   = unique_ids - delete_groups
    print(f'  ✓ Duplicate report saved: {unique_ids} member ID(s) with duplicates '
          f'({delete_groups} require manual deletion, {wait_groups} will auto-expire).')


# ──────────────────────────────────────────────────────────────────────────────
# GOOGLE DRIVE UPLOAD
# ──────────────────────────────────────────────────────────────────────────────

def _get_drive_service():
    """
    Return an authenticated Google Drive service object.

    Authentication strategy (in priority order):
    1. Cloud Run: the service account assigned to the Cloud Run job supplies
       credentials automatically via the GCP metadata server — no key file
       or environment variable needed. Application Default Credentials (ADC)
       picks this up transparently.
    2. Local development: run 'gcloud auth application-default login' once
       to cache credentials in ~/.config/gcloud/. ADC finds them automatically.

    The service account must be granted Content Manager (or Editor) on the
    two Shared Drive folders used by this script.
    """
    try:
        from google.auth import default as google_auth_default
        from googleapiclient.discovery import build
    except ImportError:
        print('\nERROR: Google API libraries not installed.')
        print('Run: pip install google-api-python-client google-auth '
              'google-auth-httplib2')
        sys.exit(1)

    creds, _ = google_auth_default(scopes=DRIVE_SCOPES)
    return build('drive', 'v3', credentials=creds)


def upload_to_drive(file_paths: list[str], timestamp: str) -> str:
    """
    Create a dated subfolder in DRIVE_SYNC_FOLDER_ID and upload all files.
    Returns the subfolder web view URL.
    """
    from googleapiclient.http import MediaFileUpload

    print('\nUploading files to Google Drive...')
    service = _get_drive_service()

    folder_name = f'Sync_{timestamp}'
    folder_meta = {
        'name':     folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents':  [DRIVE_SYNC_FOLDER_ID],
    }
    folder = service.files().create(
        body=folder_meta, fields='id, webViewLink'
    ).execute()
    folder_id  = folder['id']
    folder_url = folder.get('webViewLink',
                            f'https://drive.google.com/drive/folders/{folder_id}')
    print(f'  Created subfolder: {folder_name}')

    for fp in file_paths:
        fname = os.path.basename(fp)
        media = MediaFileUpload(fp, resumable=False)
        service.files().create(
            body={'name': fname, 'parents': [folder_id]},
            media_body=media,
            fields='id',
        ).execute()
        print(f'  ✓ Uploaded: {fname}')

    print(f'  ✓ All files available at: {folder_url}')
    return folder_url


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    import argparse
    parser = argparse.ArgumentParser(
        description='Compare ICF Global member export to GlueUp and generate import files.'
    )
    parser.add_argument(
        'icf_file',
        nargs='?',
        help='Path to the ICF Global member export .csv file (activemembers_YYYYMMDD.csv). '
             'If omitted, the most recent *activemembers*.csv in the current directory is used.',
    )
    parser.add_argument(
        '--import-mode',
        choices=['all', 'new', 'changed'],
        default='all',
        dest='import_mode',
        help=(
            'Which records to include in the import file. '
            '"all" = NEW + CHANGED (default), '
            '"new" = NEW only, '
            '"changed" = CHANGED only.'
        ),
    )
    parser.add_argument(
        '--no-drive',
        action='store_true',
        dest='no_drive',
        help='Skip Google Drive upload and save output files locally only.',
    )
    args = parser.parse_args()

    # Determine ICF file path
    if args.icf_file:
        icf_path = args.icf_file
    else:
        icf_path = find_latest_icf_file()

    if not Path(icf_path).exists():
        print(f'ERROR: File not found: {icf_path}')
        sys.exit(1)

    date_stamp = datetime.now().strftime('%Y%m%d_%H%M')
    report_path    = f'comparison_report_{date_stamp}.xlsx'
    import_path    = f'glueup_import_{date_stamp}.xlsx'
    duplicate_path = f'duplicate_members_{date_stamp}.xlsx'
    dropped_path   = f'dropped_members_{date_stamp}.xlsx'

    print('=' * 60)
    print('ICF Global → GlueUp Member Sync Comparison')
    print(f'  Version:     {SCRIPT_VERSION}')
    print(f'  Import mode: {args.import_mode.upper()}')
    print('=' * 60)

    # Step 1: Read ICF Global file
    icf_records = read_icf_file(icf_path)
    icf_by_id = {r['Member_ID']: r for r in icf_records if r.get('Member_ID')}

    # Step 2: Pull GlueUp data
    token = get_token()
    glueup_members = fetch_all_glueup_members(token)
    glueup_index   = index_glueup_by_member_id(glueup_members)

    # Step 3: Compare
    print('Comparing records...')
    results = compare_records(icf_records, glueup_index)

    new_count     = sum(1 for r in results if r['status'] == 'NEW')
    changed_count = sum(1 for r in results if r['status'] == 'CHANGED')
    same_count    = sum(1 for r in results if r['status'] == 'SAME')

    # Of the CHANGED records, how many differ ONLY on Expiration Date?
    # These will generate WAIT duplicates after import — harmless but noisy.
    expiry_only_count = sum(
        1 for r in results
        if r['status'] == 'CHANGED'
        and set(r['changed_fields']) == {'Expiration Date'}
    )
    changed_other_count = changed_count - expiry_only_count

    print(f'  NEW:                          {new_count}')
    print(f'  CHANGED (total):              {changed_count}')
    print(f'    of which expiry date only:  {expiry_only_count}')
    print(f'    of which other fields:      {changed_other_count}')
    print(f'  SAME:                         {same_count}')
    if expiry_only_count:
        print(f'  NOTE: {expiry_only_count} records will generate WAIT duplicates after import'
              f' — GlueUp will auto-expire the old records once their date passes.')

    # Step 4: Find existing duplicates in GlueUp
    print('Checking for duplicate memberships in GlueUp...')
    duplicates = find_duplicates(glueup_index)

    # Step 5: Find dropped members (in GlueUp but not in ICF Global)
    print('Checking for dropped members...')
    icf_ids = set(icf_by_id.keys())
    dropped = find_dropped_members(glueup_index, icf_ids)

    # Step 6: Write output files
    write_comparison_report(results, icf_records, report_path,
                            import_mode=args.import_mode,
                            icf_records_by_id=icf_by_id)
    write_import_file(results, icf_by_id, import_path, import_mode=args.import_mode)
    write_duplicate_file(duplicates, duplicate_path)
    write_dropped_file(dropped, dropped_path)

    output_files = [report_path, import_path, duplicate_path, dropped_path]

    print()
    print('=' * 60)
    print('DONE')
    print(f'  Comparison report : {report_path}')
    print(f'  Import file       : {import_path}')
    print(f'  Duplicate report  : {duplicate_path}')
    print(f'  Dropped members   : {dropped_path}')
    print('=' * 60)

    # Step 7: Google Drive upload
    if args.no_drive:
        print('\n(Drive upload skipped — --no-drive flag set)')
    else:
        try:
            folder_url = upload_to_drive(output_files, date_stamp)
            print(f'\n✓ Sync complete. Drive folder: {folder_url}')
        except Exception as e:
            print(f'\n⚠  Drive upload failed: {e}')
            print('Output files saved locally. Re-run with --no-drive to skip upload.')


if __name__ == '__main__':
    main()