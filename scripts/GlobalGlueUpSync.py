#!/usr/bin/env python3
"""
GlobalGlueUpSync.py
ICF Washington State Chapter — GlueUp Member Sync
Version: 1.19

Reads the ICF Global active-member CSV (produced by the Get Active Members
Make scenario), looks up each member in GlueUp by email then by ICF Member ID,
and produces five output files:
  - contact_import_YYYYMMDD_HHMM.xlsx
  - membership_import_YYYYMMDD_HHMM.xlsx
  - comparison_report_YYYYMMDD_HHMM.xlsx
  - duplicate_report_YYYYMMDD_HHMM.xlsx
  - run_log_YYYYMMDD_HHMM.txt

Usage:
  python3 GlobalGlueUpSync.py               # auto-discovers most recent activemembers*.csv
  python3 GlobalGlueUpSync.py <file.csv>    # use a specific file
  python3 GlobalGlueUpSync.py --no-drive
  python3 GlobalGlueUpSync.py --dry-run

See GlueUp_Sync_Design_Spec for full field and logic documentation.
"""

import argparse
import csv
import glob
import hmac
import hashlib
import json
import os
import sys
import time
import datetime
from collections import defaultdict
from urllib.request import urlopen, Request
from urllib.error import HTTPError, URLError

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip3 install openpyxl --break-system-packages")
    sys.exit(1)

# Google Drive upload (optional — only imported when Drive upload is attempted)
# Install: pip3 install google-api-python-client google-auth-httplib2 google-auth-oauthlib --break-system-packages
_DRIVE_LIBS_AVAILABLE = False
try:
    from googleapiclient.discovery import build as _gdrive_build
    from googleapiclient.http import MediaFileUpload as _MediaFileUpload
    from google.oauth2.credentials import Credentials as _GCredentials
    from google_auth_oauthlib.flow import InstalledAppFlow as _InstalledAppFlow
    from google.auth.transport.requests import Request as _GRequest
    _DRIVE_LIBS_AVAILABLE = True
except ImportError:
    pass  # reported at upload time if Drive upload is attempted

# ─── Configuration ────────────────────────────────────────────────────────────

GLUEUP_BASE_URL           = "https://api-services.glueup.com"
GLUEUP_ORG_ID             = "7912"
GLUEUP_PK                 = "icfwshts"
GLUEUP_SK                 = "MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPfXb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfUCCAP67vZiAQ55"
GLUEUP_MEMBERSHIP_TYPE_ID = 37600

SHADOW_EMAIL_DOMAIN = "members.icfwashingtonstate.org"
TOKEN_FILE          = "glueup_token.json"
DRIVE_FOLDER_ID     = "1BQ53mlYzkl3N5wz6AiTZm1kDpibempJy"   # Shared Drive: Data-Transfer > Sync
DRIVE_TOKEN_FILE    = "drive_token.json"
DRIVE_CREDENTIALS   = "credentials.json"
DRIVE_SCOPES        = ["https://www.googleapis.com/auth/drive.file"]

RUN_TS            = datetime.datetime.now()
RUN_DATE_MMDDYYYY = RUN_TS.strftime("%m/%d/%Y")
RUN_TS_STR        = RUN_TS.strftime("%Y%m%d_%H%M")

# ─── Input CSV column names (as output by the Make scenario) ──────────────────
# These are the exact header strings Make writes to the CSV.
COL_MEMBER_ID     = "Member ID"
COL_STATUS        = "Status"
COL_FIRST_NAME    = "First Name"
COL_LAST_NAME     = "Last Name"
COL_ROLE          = "Role"
COL_EMAIL         = "Email"
COL_PHONE         = "Phone"
COL_CITY          = "City"
COL_STATE         = "State"
COL_ZIP           = "Zip"
COL_COUNTRY       = "Country"
COL_CHAPTER_START = "Chapter_Start_Date"
COL_JOIN_DATE     = "Membership_Join_Date"
COL_EXPIRY        = "Expiration Date"
COL_REJOIN        = "Rejoin"
COL_AUTO_RENEWAL  = "Auto Renewal"
COL_CREDENTIAL    = "Credential"
COL_CRED_AWARD    = "Credential Award Date"
COL_CRED_EXPIRE   = "Credential Expire Date"
COL_TC_CRED       = "ACTC_Credential"
COL_TC_AWARD      = "ACTC_Credential_Award_Date"
COL_TC_EXPIRE     = "ACTC_Credential_Expire_Date"
COL_MEMBER_TYPE   = "Member_Type"   # optional — added to Make scenario later

# ─── State / Province lookup (2-letter → full name) ───────────────────────────

STATE_LOOKUP = {
    # US states
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho",
    "IL": "Illinois", "IN": "Indiana", "IA": "Iowa", "KS": "Kansas",
    "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma",
    "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah",
    "VT": "Vermont", "VA": "Virginia", "WA": "Washington", "WV": "West Virginia",
    "WI": "Wisconsin", "WY": "Wyoming", "DC": "District of Columbia",
    # US territories
    "PR": "Puerto Rico", "GU": "Guam", "VI": "U.S. Virgin Islands",
    "AS": "American Samoa", "MP": "Northern Mariana Islands",
    # Canadian provinces / territories
    "AB": "Alberta", "BC": "British Columbia", "MB": "Manitoba",
    "NB": "New Brunswick", "NL": "Newfoundland and Labrador",
    "NS": "Nova Scotia", "NT": "Northwest Territories", "NU": "Nunavut",
    "ON": "Ontario", "PE": "Prince Edward Island", "QC": "Quebec",
    "SK": "Saskatchewan", "YT": "Yukon",
}

# ─── Local Region lookup (zip → region) ───────────────────────────────────────
# Loaded from ZipCodes_Areas_CCC.xlsx at runtime (see load_zip_region_lookup).
# Key: (city_lower, zip_str) → region name
# Falls back to zip-only match if city+zip not found.
ZIP_REGION_LOOKUP = {}   # populated by load_zip_region_lookup()
ZIP_ONLY_LOOKUP   = {}   # populated by load_zip_region_lookup()

# ─── Comparison fields (ICF Global col → GlueUp property key) ─────────────────
# Used to build the comparison report diff.
COMPARISON_FIELDS = [
    # (label, icf_csv_col, glueup_property, transform)
    # All alphabetic fields use "lower" to avoid false CHANGED on capitalisation differences.
    ("First Name",                COL_FIRST_NAME,  "givenName",                      "lower"),
    ("Last Name",                 COL_LAST_NAME,   "familyName",                     "lower"),
    ("Email",                     "_effective_email", "emailAddress",                "lower"),
    ("City",                      COL_CITY,        "city",                           "lower"),
    ("Zip",                       COL_ZIP,         "zipCode",                        "none"),
    ("Membership Expiration Date",COL_EXPIRY,      "icfglobalmembershipenddate",     "date"),
    ("ICF Credential",            COL_CREDENTIAL,  "icfcredential",                  "lower"),
    ("Credential Award Date",     COL_CRED_AWARD,  "icfcredentialawarddate",         "date"),
    ("Credential Expire Date",    COL_CRED_EXPIRE, "icfcredentialexpiredate",        "date"),
    ("TC Credential",             COL_TC_CRED,     "icfteamcoachingcredential",      "lower"),
    ("TC Credential Award Date",  COL_TC_AWARD,    "icfteamcoachingcredentialaward", "date"),
    ("TC Credential Expire Date", COL_TC_EXPIRE,   "icfteamcoachingcredentialexpir", "date"),
    ("Auto Renewal",              COL_AUTO_RENEWAL,"icfglobalautorenewal",           "autorenewal"),
]

# ─── Logger ───────────────────────────────────────────────────────────────────

_log_lines = []

def log(msg):
    print(msg)
    _log_lines.append(msg)

def save_log(path):
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(_log_lines))

# ─── Input file discovery ─────────────────────────────────────────────────────

def find_input_file():
    """
    Find the most recently modified file with 'activemembers' in its name
    in the current directory. Returns the path, or exits with an error.
    """
    candidates = glob.glob("*activemembers*")
    if not candidates:
        log("ERROR: No file with 'activemembers' in the name found in the current directory.")
        log("  Either run the Get Active Members Make scenario first, or pass the file path explicitly.")
        sys.exit(1)
    # Pick the most recently modified
    candidates.sort(key=os.path.getmtime, reverse=True)
    chosen = candidates[0]
    if len(candidates) > 1:
        log(f"  Found {len(candidates)} activemembers files. Using most recently modified: {chosen}")
        for f in candidates[1:]:
            log(f"    (ignored) {f}")
    return chosen

# ─── GlueUp Authentication ────────────────────────────────────────────────────

def make_a_header(method="POST"):
    ts = str(int(time.time() * 1000))
    msg = method + GLUEUP_PK + "1.0" + ts
    d = hmac.new(GLUEUP_SK.encode(), msg.encode(), hashlib.sha256).hexdigest()
    return f"v=1.0;k={GLUEUP_PK};ts={ts};d={d}"

def glueup_auth():
    """Read token.json. Return dict of auth headers. Exit on failure."""
    if not os.path.exists(TOKEN_FILE):
        log(f"ERROR: {TOKEN_FILE} not found. Run: python3 GetToken.py")
        sys.exit(1)
    with open(TOKEN_FILE, "r") as f:
        data = json.load(f)
    token = data.get("token") or data.get("value", {}).get("token")
    if not token:
        log(f"ERROR: Could not read token from {TOKEN_FILE}. Run: python3 GetToken.py")
        sys.exit(1)
    return {"token": token, "requestOrganizationId": GLUEUP_ORG_ID}

def glueup_post(endpoint, body_dict, auth_headers):
    """POST to GlueUp API. Returns parsed JSON or raises on error."""
    a_header = make_a_header("POST")
    headers = {
        "Content-Type": "application/json",
        "a": a_header,
        "User-Agent": "Mozilla/5.0",
        **auth_headers,
    }
    body = json.dumps(body_dict).encode()
    req = Request(
        f"{GLUEUP_BASE_URL}{endpoint}",
        data=body,
        headers=headers,
        method="POST",
    )
    try:
        with urlopen(req, timeout=30) as resp:
            return json.loads(resp.read())
    except HTTPError as e:
        if e.code == 401:
            log("ERROR: GlueUp returned 401 Unauthorized. Token may be expired. Run: python3 GetToken.py")
            sys.exit(1)
        raise

# ─── GlueUp Lookups ───────────────────────────────────────────────────────────

def _extract_member_record(result):
    """Extract the first member from a membershipDirectory response."""
    items = result.get("value", [])
    return items[0] if items else None

def _unwrap(rec):
    """
    The /membershipDirectory/members list endpoint wraps each record as:
      { "membership": {...}, "individualMember": {...} }
    Unwrap to the individualMember dict if present; otherwise return as-is
    (handles both list responses and single-record responses).
    """
    return rec.get("individualMember", rec)

def build_glueup_index(all_records):
    """
    Build two in-memory lookup dicts from the full GlueUp member list.
    Called once at startup; all per-member lookups then use these dicts.

    Returns:
      by_email     — {email_lower: individualMember record}
      by_member_id — {icf_member_id_str: individualMember record}
    """
    by_email     = {}
    by_member_id = {}
    for raw in all_records:
        rec = _unwrap(raw)
        # Email index
        email_raw = rec.get("emailAddress") or {}
        email = (email_raw.get("value", "") if isinstance(email_raw, dict) else str(email_raw)).strip().lower()
        if email:
            by_email[email] = rec
        # ICF Member ID index (stored in custom properties)
        props  = rec.get("properties", {}) or {}
        icf_id = str(props.get("icfmemberid", "") or "").strip()
        if icf_id:
            by_member_id[icf_id] = rec
    log(f"  Index built: {len(by_email)} by email, {len(by_member_id)} by member ID.")
    return by_email, by_member_id

def get_all_glueup_members(auth):
    """
    Fetch ALL members from GlueUp for duplicate detection.
    Pages through with limit/offset. Returns list of all records.
    """
    log("  Fetching all GlueUp members for duplicate detection...")
    all_records = []
    offset = 0
    limit  = 100
    while True:
        try:
            result = glueup_post(
                "/v2/membershipDirectory/members",
                {"limit": limit, "offset": offset},
                auth,
            )
            batch = result.get("value", [])
            all_records.extend(batch)
            if len(batch) < limit:
                break
            offset += limit
        except Exception as e:
            log(f"  WARNING: Error fetching GlueUp members at offset {offset}: {e}")
            break
    log(f"  Retrieved {len(all_records)} total GlueUp member records.")
    return all_records

# ─── Shadow Email ─────────────────────────────────────────────────────────────

def make_shadow_email(member_id):
    return f"id_{member_id}@{SHADOW_EMAIL_DOMAIN}"

def is_shadow_email(email):
    return email.endswith(f"@{SHADOW_EMAIL_DOMAIN}")

# ─── Skip Logic ───────────────────────────────────────────────────────────────

DATE_COLS_FOR_SKIP = [
    COL_CHAPTER_START, COL_JOIN_DATE, COL_EXPIRY,
    COL_CRED_AWARD, COL_TC_AWARD,
]

def should_skip(row):
    """
    Skip shadow-email members that have no date data at all —
    they have nothing useful to import.
    Returns (True, reason_str) or (False, "").
    """
    if row.get(COL_EMAIL, "").strip():
        return False, ""
    for col in DATE_COLS_FOR_SKIP:
        if row.get(col, "").strip():
            return False, ""
    name = f"{row.get(COL_FIRST_NAME,'')} {row.get(COL_LAST_NAME,'')}".strip()
    return True, f"SKIP: {name} (ID {row.get(COL_MEMBER_ID,'')}) — shadow email, no date data"

# ─── Row Validation ──────────────────────────────────────────────────────────

# Known valid credential codes (lowercase)
VALID_CREDENTIALS = {"", "acc", "pcc", "mcc"}
VALID_TC_CREDENTIALS = {"", "actc"}

# Rough set of known country names and 2-letter codes to sanity-check the field.
# Not exhaustive — just enough to catch numeric zips or dates landing in Country.
_COUNTRY_DIGITS_RE = None  # compiled lazily

import re as _re

def _looks_like_date(val):
    """Return True if val could be a date string (MM/DD/YYYY or YYYY-MM-DD)."""
    return bool(_re.match(r'^\d{1,2}/\d{1,2}/\d{2,4}$', val)
                or _re.match(r'^\d{4}-\d{2}-\d{2}$', val))

def _looks_like_number(val):
    """Return True if val is purely numeric (could be a zip code or phone)."""
    return bool(_re.match(r'^\+?[\d\s\-\.]+$', val))

def validate_row(row):
    """
    Sanity-check a row for obvious data-shift or format errors.
    Returns (True, [list of error strings]) if invalid, (False, []) if OK.
    Hard-skip rules — any failure causes the row to be excluded from all output.
    """
    errors = []
    member_id = row.get(COL_MEMBER_ID, "").strip()
    name = f"{row.get(COL_FIRST_NAME,'').strip()} {row.get(COL_LAST_NAME,'').strip()}".strip()
    prefix = f"ID {member_id} ({name})"

    # Member ID must be numeric
    if member_id and not member_id.isdigit():
        errors.append(f"{prefix}: Member ID is not numeric: {member_id!r}")

    # Email must contain @ and a dot (if present — shadow email members have no email)
    email = row.get(COL_EMAIL, "").strip()
    if email and ("@" not in email or "." not in email.split("@")[-1]):
        errors.append(f"{prefix}: Email looks invalid: {email!r}")

    # Country must not look like a number or a date
    country = row.get(COL_COUNTRY, "").strip()
    if country:
        if _looks_like_number(country) and not any(c.isalpha() for c in country):
            errors.append(f"{prefix}: Country looks like a number (possible column shift): {country!r}")
        if _looks_like_date(country):
            errors.append(f"{prefix}: Country looks like a date (possible column shift): {country!r}")

    # Zip must not look like a country name (all alpha, length > 3)
    zip_val = row.get(COL_ZIP, "").strip()
    if zip_val and zip_val.isalpha() and len(zip_val) > 3:
        errors.append(f"{prefix}: Zip looks like a country name (possible column shift): {zip_val!r}")

    # Date fields must parse as dates if non-blank
    date_fields = [
        (COL_CHAPTER_START, "Chapter Start Date"),
        (COL_JOIN_DATE,     "Membership Join Date"),
        (COL_EXPIRY,        "Expiration Date"),
        (COL_CRED_AWARD,    "Credential Award Date"),
        (COL_CRED_EXPIRE,   "Credential Expire Date"),
        (COL_TC_AWARD,      "TC Credential Award Date"),
        (COL_TC_EXPIRE,     "TC Credential Expire Date"),
    ]
    for col, label in date_fields:
        val = row.get(col, "").strip()
        if val:
            try:
                datetime.datetime.strptime(val, "%m/%d/%Y")
            except ValueError:
                errors.append(f"{prefix}: {label} is not a valid date: {val!r}")

    # Credential must be a known code or blank
    cred = row.get(COL_CREDENTIAL, "").strip().lower()
    if cred not in VALID_CREDENTIALS:
        errors.append(f"{prefix}: Credential is not a recognised code: {row.get(COL_CREDENTIAL,'')!r} (expected one of {sorted(VALID_CREDENTIALS)})")

    # TC Credential must be a known code or blank
    tc_cred = row.get(COL_TC_CRED, "").strip().lower()
    if tc_cred not in VALID_TC_CREDENTIALS:
        errors.append(f"{prefix}: TC Credential is not a recognised code: {row.get(COL_TC_CRED,'')!r} (expected one of {sorted(VALID_TC_CREDENTIALS)})")

    return bool(errors), errors

# ─── Field Transformations ────────────────────────────────────────────────────

def expand_state(state_code):
    if not state_code:
        return ""
    return STATE_LOOKUP.get(state_code.strip().upper(), state_code.strip())

def transform_credential(value):
    """Lowercase credential code, or empty string."""
    return value.strip().lower() if value and value.strip() else ""

def transform_auto_renewal(value):
    """Yes → yes, anything else (including blank) → no."""
    return "yes" if (value or "").strip().lower() == "yes" else "no"

def get_local_region(city, zip_code):
    """
    Look up region from ZipCodes_Areas_CCC lookup.
    Try city+zip first, then zip only. Fall back to 'No Region'.
    """
    city_key = (city or "").strip().lower()
    zip_key  = (zip_code or "").strip()
    return (
        ZIP_REGION_LOOKUP.get((city_key, zip_key))
        or ZIP_ONLY_LOOKUP.get(zip_key)
        or "No Region"
    )

# ─── Zip/Region lookup loader ─────────────────────────────────────────────────

def load_zip_region_lookup(xlsx_path=None):
    """
    Load ZipCodes_Areas_CCC.xlsx into ZIP_REGION_LOOKUP and ZIP_ONLY_LOOKUP.
    Columns: A=Region, B=City, C=Zip
    Silently skips if file not found — all members get 'No Region'.
    """
    global ZIP_REGION_LOOKUP, ZIP_ONLY_LOOKUP
    candidates = [
        xlsx_path,
        "ZipCodes_Areas_CCC.xlsx",
        os.path.join(os.path.dirname(__file__), "ZipCodes_Areas_CCC.xlsx"),
    ]
    path = next((p for p in candidates if p and os.path.exists(p)), None)
    if not path:
        log("  WARNING: ZipCodes_Areas_CCC.xlsx not found — Local Region will be 'No Region' for all members.")
        return
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 3:
                continue
            region = str(row[0] or "").strip()
            city   = str(row[1] or "").strip().lower()
            zip_   = str(row[2] or "").strip()
            if region and zip_:
                ZIP_ONLY_LOOKUP[zip_] = region
                if city:
                    ZIP_REGION_LOOKUP[(city, zip_)] = region
                count += 1
        wb.close()
        log(f"  Loaded {count} zip/region entries from {path}.")
    except Exception as e:
        log(f"  WARNING: Could not load ZipCodes_Areas_CCC.xlsx: {e} — Local Region will be 'No Region'.")

# ─── CSV Loading ──────────────────────────────────────────────────────────────

REQUIRED_COLS = [
    COL_MEMBER_ID, COL_FIRST_NAME, COL_LAST_NAME, COL_EMAIL, COL_EXPIRY,
]
OPTIONAL_COLS = [COL_MEMBER_TYPE]

def load_csv(path):
    """
    Read input CSV. Returns list of row dicts.
    Exits if required columns are missing.
    Warns (but continues) if optional columns are absent.
    """
    if not os.path.exists(path):
        log(f"ERROR: Input file not found: {path}")
        sys.exit(1)
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        missing_required = [c for c in REQUIRED_COLS if c not in headers]
        if missing_required:
            log(f"ERROR: Input CSV missing required columns: {missing_required}")
            log(f"  Found columns: {list(headers)}")
            sys.exit(1)
        missing_optional = [c for c in OPTIONAL_COLS if c not in headers]
        for col in missing_optional:
            log(f"  NOTE: Optional column '{col}' not present — field will be blank in output.")
        for row in reader:
            if not any(v.strip() for v in row.values()):
                continue  # skip blank rows
            # Ensure optional cols exist in every row dict (blank if absent)
            for col in missing_optional:
                row.setdefault(col, "")
            rows.append(row)
    log(f"  Loaded {len(rows)} data rows from {path}.")
    return rows

# ─── GlueUp record field extraction ──────────────────────────────────────────

def extract_glueup_name(rec):
    """Return (first_name, last_name) from a GlueUp member record."""
    first = (rec.get("givenName") or "").strip()
    last  = (rec.get("familyName") or "").strip()
    return first, last

def _glueup_addr(rec):
    """Return the address sub-object, checking both top-level and nested."""
    return rec.get("address") or {}

def _glueup_city(rec):
    """City is stored as address.cityName in GlueUp API responses."""
    addr = _glueup_addr(rec)
    return str(addr.get("cityName") or addr.get("city") or rec.get("city") or "").strip()

def _glueup_zip(rec):
    """Zip is stored as address.zipCode in GlueUp API responses."""
    addr = _glueup_addr(rec)
    return str(addr.get("zipCode") or rec.get("zipCode") or "").strip()

def _extract_prop_value(val):
    """
    GlueUp custom properties that are single-select fields return a dict:
      {"code": "none", "title": {"en": "None"}}
    Extract the code string in that case.
    Dates as epoch-ms ints are converted to MM/DD/YYYY.
    Dates as YYYY-MM-DD strings are converted to MM/DD/YYYY.
    Everything else is returned as a stripped string.
    """
    if val is None:
        return ""
    # Single-select / dropdown field — extract the code
    if isinstance(val, dict):
        return str(val.get("code", "") or "").strip()
    # Epoch-ms timestamp
    if isinstance(val, (int, float)) and val > 1_000_000_000_000:
        try:
            return datetime.datetime.utcfromtimestamp(val / 1000).strftime("%m/%d/%Y")
        except Exception:
            pass
    # YYYY-MM-DD string
    if isinstance(val, str) and len(val) == 10 and val[4] == "-":
        try:
            return datetime.datetime.strptime(val, "%Y-%m-%d").strftime("%m/%d/%Y")
        except Exception:
            pass
    return str(val).strip()

def extract_glueup_field(rec, prop_key):
    """
    Extract a field value from a GlueUp record.
    Handles top-level contact fields and custom properties dict.
    """
    # Top-level contact fields
    if prop_key == "givenName":
        return str(rec.get("givenName") or "").strip()
    if prop_key == "familyName":
        return str(rec.get("familyName") or "").strip()
    if prop_key == "emailAddress":
        raw = rec.get("emailAddress") or {}
        return (raw.get("value", "") if isinstance(raw, dict) else str(raw)).strip()
    if prop_key == "city":
        return _glueup_city(rec)
    if prop_key == "zipCode":
        return _glueup_zip(rec)

    # Custom properties
    props = rec.get("properties", {}) or {}
    return _extract_prop_value(props.get(prop_key))

def extract_glueup_import_date(rec):
    """Extract icfimportdate from GlueUp record, for duplicate detection."""
    return extract_glueup_field(rec, "icfimportdate")

# ─── Row Builders ─────────────────────────────────────────────────────────────

def build_contact_row(row, glueup_rec, effective_email, has_real_email):
    """
    Build one row dict for the Contact import XLSX.
    All 33 columns from sample_contact_import.xlsx must be present.
    Blue (ignored) fields are left blank. Active fields are populated.
    """
    shadow = not has_real_email
    if glueup_rec and not shadow:
        first, last = extract_glueup_name(glueup_rec)
    else:
        first = row.get(COL_FIRST_NAME, "").strip()
        last  = row.get(COL_LAST_NAME,  "").strip()

    credential = transform_credential(row.get(COL_CREDENTIAL, ""))
    tc_cred    = transform_credential(row.get(COL_TC_CRED, ""))

    return {
        "First Name":                               first,
        "Last Name":                                last,
        "Address":                                  "",          # blue — ignored
        "City":                                     row.get(COL_CITY, "").strip(),
        "State/Province":                           expand_state(row.get(COL_STATE, "")),
        "Postal Code/Zip Code":                     row.get(COL_ZIP, "").strip(),
        "Email":                                    effective_email,
        "Phone":                                    row.get(COL_PHONE, "").strip(),
        "Company":                                  "",          # blue — ignored
        "Title/Position":                           "",          # blue — ignored
        "Volunteer Role":                           "",          # blue — ignored
        "Findable":                                 "",          # blue — ignored
        "Directory Listing Text":                   "",          # blue — ignored
        "Coach Industry":                           "",          # blue — ignored
        "Coaching Specialization":                  "",          # blue — ignored
        "Has Email":                                "yes" if has_real_email else "no",
        "ICF Chapter Start Date":                   row.get(COL_CHAPTER_START, "").strip(),
        "ICF Credential":                           credential,
        "ICF Credential Award Date":                row.get(COL_CRED_AWARD, "").strip() if credential else "",
        "ICF Credential Expire Date":               row.get(COL_CRED_EXPIRE, "").strip() if credential else "",
        "ICF Global Auto Renewal":                  "",          # blue — ignored (belongs to membership)
        "ICF Global Member ID":                     row.get(COL_MEMBER_ID, "").strip(),
        "ICF Global Member Status":                 row.get(COL_STATUS, "").strip(),
        "ICF Team Coaching Credential":             tc_cred,
        "ICF Team Coaching Credential Award Date":  row.get(COL_TC_AWARD, "").strip() if tc_cred else "",
        "ICF Team Coaching Credential Expire Date": row.get(COL_TC_EXPIRE, "").strip() if tc_cred else "",
        "ICF Global Import Date":                   RUN_DATE_MMDDYYYY,
        "Local Region":                             get_local_region(row.get(COL_CITY, ""), row.get(COL_ZIP, "")),
        "ICF Global Member Type":                   row.get(COL_MEMBER_TYPE, "").strip(),
        "ICF Global Membership End Date":           "",          # blue — ignored (belongs to membership)
        "ICF Global Membership Restart Date":       "",          # blue — ignored (belongs to membership)
        "ICF Global Membership Start Date":         "",          # blue — ignored (belongs to membership)
        "ICF Global Membership Type":               "",          # blue — ignored (belongs to membership)
    }

# Exact column order from sample_contact_import.xlsx (33 columns — all must be present)
CONTACT_FIELDNAMES = [
    "First Name", "Last Name", "Address", "City", "State/Province",
    "Postal Code/Zip Code", "Email", "Phone", "Company", "Title/Position",
    "Volunteer Role", "Findable", "Directory Listing Text", "Coach Industry",
    "Coaching Specialization", "Has Email", "ICF Chapter Start Date",
    "ICF Credential", "ICF Credential Award Date", "ICF Credential Expire Date",
    "ICF Global Auto Renewal", "ICF Global Member ID", "ICF Global Member Status",
    "ICF Team Coaching Credential", "ICF Team Coaching Credential Award Date",
    "ICF Team Coaching Credential Expire Date", "ICF Global Import Date",
    "Local Region", "ICF Global Member Type", "ICF Global Membership End Date",
    "ICF Global Membership Restart Date", "ICF Global Membership Start Date",
    "ICF Global Membership Type",
]

def build_membership_row(row, glueup_rec, effective_email, has_real_email):
    """
    Build one row dict for the Membership import XLSX.
    All 34 columns from sample_membership_import.xlsx must be present.
    Blue (ignored) fields are left blank. Active fields are populated.
    First Name and Last Name are required by GlueUp for record matching.
    """
    shadow = not has_real_email
    if glueup_rec and not shadow:
        first, last = extract_glueup_name(glueup_rec)
    else:
        first = row.get(COL_FIRST_NAME, "").strip()
        last  = row.get(COL_LAST_NAME,  "").strip()

    return {
        "Membership Start Date":                        row.get(COL_JOIN_DATE, "").strip(),
        "Membership End Date":                          row.get(COL_EXPIRY, "").strip(),
        "Currency":                                     "",          # blue — ignored
        "First Name":                                   first,
        "Last Name":                                    last,
        "Email":                                        effective_email,
        "Phone":                                        "",          # blue — ignored (in contact import)
        "Postal Code/Zip Code":                         "",          # blue — ignored (in contact import)
        "Address":                                      "",          # blue — ignored
        "City":                                         "",          # blue — ignored (in contact import)
        "State/Province":                               "",          # blue — ignored (in contact import)
        "Country/Region":                               "",          # blue — ignored
        "Volunteer Role":                               "",          # blue — ignored
        "Coach Industry":                               "",          # blue — ignored
        "Coaching Specialization":                      "",          # blue — ignored
        "Company Name":                                 "",          # blue — ignored
        "Function":                                     "",          # blue — ignored
        "Title/Position":                               "",          # blue — ignored
        "Findable":                                     "",          # blue — ignored
        "Directory Listing Text":                       "",          # blue — ignored
        "ICF Credential":                               "",          # blue — ignored (in contact import)
        "ICF Credential Award Date":                    "",          # blue — ignored (in contact import)
        "ICF Credential Expire Date":                   "",          # blue — ignored (in contact import)
        "ICF Global Auto Renewal":                      transform_auto_renewal(row.get(COL_AUTO_RENEWAL, "")),
        "ICF Global Member ID":                         row.get(COL_MEMBER_ID, "").strip(),
        "ICF Global Membership End Date":               row.get(COL_EXPIRY, "").strip(),
        "ICF Global Membership Start Date":             row.get(COL_JOIN_DATE, "").strip(),
        "ICF Global Membership Type":                   "individual",
        "ICF Team Coaching Credential":                 "",          # blue — ignored (in contact import)
        "ICF Team Coaching Credential Award Date":      "",          # blue — ignored (in contact import)
        "ICF Team Coaching Credential Expire Date":     "",          # blue — ignored (in contact import)
        "Has Email":                                    "",          # blue — ignored (in contact import)
        "ICF Global Import Date":                       RUN_DATE_MMDDYYYY,
        "ICF Global Membership Restart Date":           row.get(COL_REJOIN, "").strip(),
    }

# Exact column order from sample_membership_import.xlsx (34 columns — all must be present)
MEMBERSHIP_FIELDNAMES = [
    "Membership Start Date", "Membership End Date", "Currency",
    "First Name", "Last Name", "Email", "Phone", "Postal Code/Zip Code",
    "Address", "City", "State/Province", "Country/Region",
    "Volunteer Role", "Coach Industry", "Coaching Specialization",
    "Company Name", "Function", "Title/Position", "Findable",
    "Directory Listing Text",
    "ICF Credential", "ICF Credential Award Date", "ICF Credential Expire Date",
    "ICF Global Auto Renewal", "ICF Global Member ID",
    "ICF Global Membership End Date", "ICF Global Membership Start Date",
    "ICF Global Membership Type",
    "ICF Team Coaching Credential",
    "ICF Team Coaching Credential Award Date",
    "ICF Team Coaching Credential Expire Date",
    "Has Email", "ICF Global Import Date", "ICF Global Membership Restart Date",
]

# ─── Comparison Logic ─────────────────────────────────────────────────────────

# GlueUp stores "no credential" as code "none" and "no auto-renewal" as "no".
# ICF Global sends blank for both. Treat GlueUp "none" as equivalent to blank
# for credential fields, so we don't generate false CHANGED records.
_GLUEUP_BLANK_CODES = {"none", "unspecified", ""}

def normalize_for_compare(value, transform):
    """
    Normalize a value for comparison between ICF Global and GlueUp.
    Handles GlueUp's explicit 'none'/'unspecified' codes as equivalent to blank.
    """
    v = str(value or "").strip()
    if transform == "lower":
        # Treat GlueUp 'none' code as blank (= no credential in ICF)
        lowered = v.lower()
        return "" if lowered in _GLUEUP_BLANK_CODES else lowered
    if transform == "autorenewal":
        return "yes" if v.lower() == "yes" else "no"
    if transform == "date":
        try:
            return datetime.datetime.strptime(v, "%m/%d/%Y").strftime("%Y-%m-%d")
        except Exception:
            return v
    return v

def compare_record(row, glueup_rec, effective_email):
    """
    Compare ICF Global row against GlueUp record field by field.
    Returns list of dicts: [{field, icf_value, glueup_value, differs}]

    Non-overwrite rule: if ICF Global sends a blank value and GlueUp has a
    non-blank value, treat them as equal. We never want to blank out historical
    data in GlueUp (e.g. expired credential dates that ICF Global no longer sends).
    """
    diffs = []
    for label, icf_col, glueup_prop, transform in COMPARISON_FIELDS:
        icf_raw    = effective_email if icf_col == "_effective_email" else row.get(icf_col, "")
        glueup_raw = extract_glueup_field(glueup_rec, glueup_prop) if glueup_rec else ""

        icf_norm    = normalize_for_compare(icf_raw, transform)
        glueup_norm = normalize_for_compare(glueup_raw, transform)

        # Non-overwrite rule: ICF blank + GlueUp has value → not a difference
        if not icf_norm and glueup_norm:
            differs = False
        else:
            differs = icf_norm != glueup_norm

        diffs.append({
            "field":        label,
            "icf_value":    str(icf_raw).strip(),
            "glueup_value": str(glueup_raw).strip(),
            "differs":      differs,
        })
    return diffs

# ─── CSV Writing ──────────────────────────────────────────────────────────────

def write_csv(rows, path, fieldnames):
    """Write list-of-dicts to CSV with header row."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    log(f"  Written: {path} ({len(rows)} rows)")

def write_import_xlsx(rows, path, fieldnames):
    """
    Write list-of-dicts to an import-ready XLSX file with header row.
    Header row uses the dark blue style matching the other report files.
    All values written as plain strings to avoid GlueUp import type issues.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Import"

    for col, h in enumerate(fieldnames, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = FILL_HEADER
        c.font = FONT_WHITE_BOLD
        c.alignment = Alignment(wrap_text=False, vertical="center")
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20

    for row_num, row in enumerate(rows, start=2):
        for col, field in enumerate(fieldnames, start=1):
            ws.cell(row=row_num, column=col, value=row.get(field, ""))

    wb.save(path)
    log(f"  Written: {path} ({len(rows)} rows)")

# ─── Excel Report Writing ─────────────────────────────────────────────────────

FILL_GREEN       = PatternFill("solid", fgColor="C6EFCE")
FILL_YELLOW      = PatternFill("solid", fgColor="FFEB9C")
FILL_GRAY        = PatternFill("solid", fgColor="F2F2F2")
FILL_RED         = PatternFill("solid", fgColor="FFC7CE")
FILL_HEADER      = PatternFill("solid", fgColor="1F3864")
FILL_DUPE_KEEP   = PatternFill("solid", fgColor="C6EFCE")
FILL_DUPE_REVIEW = PatternFill("solid", fgColor="FFC7CE")
FONT_WHITE_BOLD  = Font(bold=True, color="FFFFFF")
FONT_BOLD        = Font(bold=True)

def _hdr_cell(ws, row, col, value, fill=None, font=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill:
        c.fill = fill
    if font:
        c.font = font
    c.alignment = Alignment(wrap_text=True, vertical="top")
    return c

def write_comparison_report(comparison_data, path):
    """
    Write comparison_report XLSX with three tabs:
      Summary  — NEW / CHANGED / SAME / SKIPPED counts
      Detail   — every member, every comparison field, colour-coded
      Changed  — only CHANGED members
    """
    wb = openpyxl.Workbook()

    # ── Summary ──────────────────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Summary"
    counts = {}
    for r in comparison_data:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    ws_sum.column_dimensions["A"].width = 20
    ws_sum.column_dimensions["B"].width = 12
    _hdr_cell(ws_sum, 1, 1,
              f"GlobalGlueUpSync — Comparison Report — {RUN_TS.strftime('%Y-%m-%d %H:%M')}",
              FILL_HEADER, FONT_WHITE_BOLD)
    ws_sum.merge_cells("A1:B1")
    _hdr_cell(ws_sum, 3, 1, "Status", FILL_GRAY, FONT_BOLD)
    _hdr_cell(ws_sum, 3, 2, "Count",  FILL_GRAY, FONT_BOLD)
    status_fills = {"NEW": FILL_GREEN, "CHANGED": FILL_YELLOW}
    for i, status in enumerate(["NEW", "CHANGED", "SAME", "SKIPPED"], start=4):
        c = ws_sum.cell(row=i, column=1, value=status)
        if status in status_fills:
            c.fill = status_fills[status]
        ws_sum.cell(row=i, column=2, value=counts.get(status, 0))

    # ── Detail & Changed tabs ────────────────────────────────────────────────
    field_labels = [f[0] for f in COMPARISON_FIELDS]
    headers = (["Status", "ICF Member ID", "Name", "Email"]
               + [f"{lbl} (ICF)" for lbl in field_labels]
               + [f"{lbl} (GlueUp)" for lbl in field_labels])

    for tab_name, filter_fn in [
        ("Detail",  lambda e: True),
        ("Changed", lambda e: e["status"] == "CHANGED"),
    ]:
        ws = wb.create_sheet(tab_name)
        for col, h in enumerate(headers, start=1):
            _hdr_cell(ws, 1, col, h, FILL_HEADER, FONT_WHITE_BOLD)
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18

        row_num = 2
        for entry in comparison_data:
            if not filter_fn(entry):
                continue
            status = entry["status"]
            row_fill = FILL_GREEN if status == "NEW" else FILL_YELLOW if status == "CHANGED" else None
            diffs = entry.get("diffs", [])
            base_vals   = [status, entry["member_id"], entry["name"], entry["email"]]
            icf_vals    = [d["icf_value"]    for d in diffs]
            glueup_vals = [d["glueup_value"] for d in diffs]
            differs     = [d["differs"]      for d in diffs]

            for col, val in enumerate(base_vals + icf_vals + glueup_vals, start=1):
                c = ws.cell(row=row_num, column=col, value=val)
                c.alignment = Alignment(vertical="top")
                if row_fill:
                    c.fill = row_fill
                # Highlight individual differing cells
                if col > 4 and len(differs) > 0:
                    field_idx = (col - 5) % len(field_labels)
                    if field_idx < len(differs) and differs[field_idx]:
                        c.fill = FILL_RED
            row_num += 1

    wb.save(path)
    log(f"  Written: {path}")

def write_duplicate_report(all_glueup_members, path):
    """
    Write duplicate_report XLSX.
    Groups by icfmemberid; flags groups with >1 record.
    Most recent icfimportdate = KEEP; older = REVIEW — CANCEL IN GLUEUP.
    """
    by_icf_id = defaultdict(list)
    for rec in all_glueup_members:
        props  = rec.get("properties", {}) or {}
        icf_id = str(props.get("icfmemberid", "") or "").strip()
        if icf_id:
            by_icf_id[icf_id].append(rec)

    dupes = {k: v for k, v in by_icf_id.items() if len(v) > 1}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Duplicates"

    col_headers = ["Action", "ICF Member ID", "GlueUp Membership ID",
                   "Name", "Email", "Import Date", "Notes"]
    last_col = openpyxl.utils.get_column_letter(len(col_headers))

    _hdr_cell(ws, 1, 1,
              f"Duplicate Report — {RUN_TS.strftime('%Y-%m-%d %H:%M')}",
              FILL_HEADER, FONT_WHITE_BOLD)
    ws.merge_cells(f"A1:{last_col}1")

    for col, h in enumerate(col_headers, start=1):
        _hdr_cell(ws, 2, col, h, FILL_HEADER, FONT_WHITE_BOLD)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 24

    def parse_import_date(rec):
        d = extract_glueup_import_date(rec)
        # Try MM/DD/YYYY first (after _extract_prop_value conversion),
        # then YYYY-MM-DD as fallback if stored as plain string
        for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(d, fmt)
            except Exception:
                pass
        return datetime.datetime.min

    row_num    = 3
    dupe_count = 0
    for icf_id, recs in sorted(dupes.items()):
        recs_sorted = sorted(recs, key=parse_import_date, reverse=True)
        for i, rec in enumerate(recs_sorted):
            action    = "KEEP" if i == 0 else "REVIEW — CANCEL IN GLUEUP"
            fill      = FILL_DUPE_KEEP if i == 0 else FILL_DUPE_REVIEW
            name      = f"{rec.get('givenName','')} {rec.get('familyName','')}".strip()
            email_raw = rec.get("emailAddress") or {}
            email     = email_raw.get("value", "") if isinstance(email_raw, dict) else str(email_raw)
            imp_date  = extract_glueup_import_date(rec)
            glu_id    = str(rec.get("id", ""))
            notes     = ("Most recent import date" if i == 0
                         else "Older record — cancel membership in GlueUp Admin UI")

            for col, val in enumerate([action, icf_id, glu_id, name, email, imp_date, notes], start=1):
                c = ws.cell(row=row_num, column=col, value=val)
                c.fill = fill
                c.alignment = Alignment(vertical="top")
            row_num    += 1
            dupe_count += 1

    if dupe_count == 0:
        ws.cell(row=3, column=1, value="No duplicates found.")

    wb.save(path)
    log(f"  Written: {path} ({len(dupes)} duplicate groups, {dupe_count} records flagged)")

# ─── Main Processing Loop ─────────────────────────────────────────────────────

def process_rows(rows, glueup_by_email, glueup_by_member_id, dry_run=False):
    """
    Main per-record loop. Returns:
      contact_rows, membership_rows, comparison_data
    """
    contact_rows    = []
    membership_rows = []
    comparison_data = []
    counts = {"total": 0, "skipped": 0, "new": 0, "new_expired": 0,
              "existing_real": 0, "existing_shadow": 0, "errors": 0, "fallback": 0}
    total = len(rows)

    for row in rows:
        counts["total"] += 1
        n         = counts["total"]
        member_id = row.get(COL_MEMBER_ID, "").strip()
        raw_email = row.get(COL_EMAIL, "").strip()
        name_preview = f"{row.get(COL_FIRST_NAME,'')} {row.get(COL_LAST_NAME,'')}".strip()

        # Step 1: Skip check
        skip, skip_msg = should_skip(row)
        if skip:
            log(f"  {skip_msg}")
            counts["skipped"] += 1
            comparison_data.append({
                "member_id": member_id,
                "name":   name_preview,
                "email":  make_shadow_email(member_id),
                "status": "SKIPPED",
                "diffs":  [],
            })
            continue

        # Step 2: Effective email
        has_real_email  = bool(raw_email)
        effective_email = raw_email.lower() if has_real_email else make_shadow_email(member_id)

        # Steps 3–5: GlueUp lookup (in-memory dict — no API call per member)
        glueup_rec = glueup_by_email.get(effective_email.lower())
        if not glueup_rec:
            counts["fallback"] += 1
            glueup_rec = glueup_by_member_id.get(str(member_id))

        # Step 6: Status
        is_new    = glueup_rec is None
        diffs     = compare_record(row, glueup_rec, effective_email)
        any_diff  = any(d["differs"] for d in diffs)

        if is_new:
            status = "NEW"
            counts["new"] += 1
        else:
            status = "CHANGED" if any_diff else "SAME"
            if has_real_email:
                counts["existing_real"] += 1
            else:
                counts["existing_shadow"] += 1

        name = f"{row.get(COL_FIRST_NAME,'')} {row.get(COL_LAST_NAME,'')}".strip()
        comparison_data.append({
            "member_id": member_id,
            "name":      name,
            "email":     effective_email,
            "status":    status,
            "diffs":     diffs,
        })

        # Step 8: Contact import — NEW and CHANGED only
        if status in ("NEW", "CHANGED"):
            contact_rows.append(
                build_contact_row(row, glueup_rec, effective_email, has_real_email)
            )

        # Step 9: Membership import:
        #   NEW + Active status → create membership
        #   NEW + non-Active (e.g. Expired) → contact only, no membership
        #   CHANGED + membership field changed → update membership
        member_status = row.get(COL_STATUS, "").strip()
        is_active = member_status.lower() == "active"

        if status == "NEW":
            if not is_active:
                # ICF Global has already marked them expired — contact only
                counts["new_expired"] += 1
            else:
                # Active per ICF Global — check expiration date
                expiry_str = row.get(COL_EXPIRY, "").strip()
                expiry_passed = False
                if expiry_str:
                    try:
                        expiry_dt = datetime.datetime.strptime(expiry_str, "%m/%d/%Y")
                        expiry_passed = expiry_dt.date() < RUN_TS.date()
                    except ValueError:
                        pass
                if expiry_passed:
                    # ICF still says Active but expiry date has passed —
                    # member is in grace period; let GlueUp handle it
                    counts["new_expired"] += 1
                else:
                    membership_rows.append(
                        build_membership_row(row, glueup_rec, effective_email, has_real_email)
                    )
        elif status == "CHANGED":
            changed_field_labels = {d["field"] for d in diffs if d["differs"]}
            membership_triggered = any(
                lbl in {"Membership Expiration Date", "Auto Renewal"}
                for lbl in changed_field_labels
            )
            if membership_triggered:
                membership_rows.append(
                    build_membership_row(row, glueup_rec, effective_email, has_real_email)
                )

    log("")
    log("── Run Summary ──────────────────────────────────────────────")
    log(f"  Total rows read:           {counts['total']}")
    log(f"  Skipped (no data):         {counts['skipped']}")
    log(f"  New — Active:              {counts['new'] - counts['new_expired']}")
    log(f"  New — no membership created (expired or in grace period): {counts['new_expired']}")
    log(f"  Existing — real email:     {counts['existing_real']}")
    log(f"  Existing — shadow email:   {counts['existing_shadow']}")
    log(f"  Errors:                    {counts['errors']}")
    log(f"  Contact rows to import:    {len(contact_rows)}")
    log(f"  Membership rows to import: {len(membership_rows)}")
    log("─────────────────────────────────────────────────────────────")

    return contact_rows, membership_rows, comparison_data

# ─── Google Drive Upload ──────────────────────────────────────────────────────

def _get_drive_service():
    """
    Build and return an authenticated Google Drive service object.
    Uses OAuth 2.0 with credentials.json (Desktop App type).
    Caches the token in drive_token.json for subsequent runs.

    Returns the service object, or None if auth fails.

    One-time setup:
      1. Enable the Google Drive API in Google Cloud Console.
      2. Create OAuth 2.0 credentials (Desktop App type).
      3. Download as credentials.json into the GG-Integration directory.
      4. First run opens a browser for consent; subsequent runs use drive_token.json.

    Future migration path (Open Item 15):
      Replace OAuth with a service account:
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_file(
            "service_account.json", scopes=DRIVE_SCOPES)
      Share the Drive folder with the service account email and remove
      the credentials.json / drive_token.json flow below.
    """
    if not _DRIVE_LIBS_AVAILABLE:
        log("  ERROR: Google Drive libraries not installed.")
        log("  Run: pip3 install google-api-python-client google-auth-httplib2 google-auth-oauthlib --break-system-packages")
        return None

    creds = None

    # Load cached token if available
    if os.path.exists(DRIVE_TOKEN_FILE):
        try:
            creds = _GCredentials.from_authorized_user_file(DRIVE_TOKEN_FILE, DRIVE_SCOPES)
        except Exception as e:
            log(f"  WARNING: Could not load {DRIVE_TOKEN_FILE}: {e}  — will re-authenticate.")

    # Refresh or obtain new credentials
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(_GRequest())
            except Exception as e:
                log(f"  WARNING: Token refresh failed: {e}  — will re-authenticate.")
                creds = None

        if not creds:
            if not os.path.exists(DRIVE_CREDENTIALS):
                log(f"  ERROR: {DRIVE_CREDENTIALS} not found.")
                log("  See the Design Spec (Google Drive Setup section) for one-time setup instructions.")
                return None
            try:
                flow = _InstalledAppFlow.from_client_secrets_file(DRIVE_CREDENTIALS, DRIVE_SCOPES)
                creds = flow.run_local_server(port=0)
            except Exception as e:
                log(f"  ERROR: OAuth flow failed: {e}")
                return None

        # Save refreshed / new token for next run
        try:
            with open(DRIVE_TOKEN_FILE, "w") as f:
                f.write(creds.to_json())
        except Exception as e:
            log(f"  WARNING: Could not save {DRIVE_TOKEN_FILE}: {e}")

    try:
        service = _gdrive_build("drive", "v3", credentials=creds)
        return service
    except Exception as e:
        log(f"  ERROR: Failed to build Drive service: {e}")
        return None


def upload_to_drive(file_paths):
    """
    Upload output files to a new dated subfolder in the Sync Output Drive folder.

    Creates:  Sync_YYYYMMDD_HHMM  inside DRIVE_FOLDER_ID
    Uploads:  all files in file_paths into that subfolder

    Logs each file uploaded and the shareable folder link.
    On any failure: logs the error and continues (output files are always local).

    Args:
        file_paths: list of local file paths to upload
    """
    log("Uploading output files to Google Drive...")

    service = _get_drive_service()
    if service is None:
        log("  Drive upload skipped — authentication failed (see above).")
        log("  Output files are saved locally. Re-run with a working credentials.json to upload.")
        return

    # Create a dated subfolder inside the Sync Output folder
    subfolder_name = f"Sync_{RUN_TS_STR}"
    try:
        folder_meta = {
            "name":     subfolder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents":  [DRIVE_FOLDER_ID],
        }
        folder = service.files().create(
            body=folder_meta,
            fields="id,webViewLink",
            supportsAllDrives=True,
        ).execute()
        folder_id   = folder.get("id")
        folder_link = folder.get("webViewLink", "")
        log(f"  Created subfolder: {subfolder_name}")
        log(f"  Folder link: {folder_link}")
    except Exception as e:
        log(f"  ERROR: Could not create Drive subfolder '{subfolder_name}': {e}")
        log("  Output files are saved locally.")
        return

    # Upload each file
    MIME_MAP = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".txt":  "text/plain",
    }
    uploaded = 0
    for path in file_paths:
        if not os.path.exists(path):
            log(f"  WARNING: Skipping missing file: {path}")
            continue
        ext      = os.path.splitext(path)[1].lower()
        mimetype = MIME_MAP.get(ext, "application/octet-stream")
        fname    = os.path.basename(path)
        try:
            file_meta = {"name": fname, "parents": [folder_id]}
            media     = _MediaFileUpload(path, mimetype=mimetype, resumable=False)
            service.files().create(
                body=file_meta,
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            ).execute()
            log(f"  Uploaded: {fname}")
            uploaded += 1
        except Exception as e:
            log(f"  WARNING: Failed to upload {fname}: {e}")

    log(f"  Drive upload complete: {uploaded}/{len(file_paths)} files uploaded to {subfolder_name}.")


# ─── Entry Point ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="GlobalGlueUpSync — ICF Washington State GlueUp Sync"
    )
    parser.add_argument(
        "input_csv", nargs="?", default=None,
        help="Path to activemembers CSV. If omitted, auto-discovers the most "
             "recently modified file with 'activemembers' in the name."
    )
    parser.add_argument("--no-drive", action="store_true",
                        help="Skip Google Drive upload")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Process all rows and print summary; do not write output files")
    args = parser.parse_args()

    log("=" * 61)
    log(f"  GlobalGlueUpSync.py  —  {RUN_TS.strftime('%Y-%m-%d %H:%M')}")
    log("=" * 61)

    # Resolve input file
    if args.input_csv:
        input_path = args.input_csv
        log(f"  Input:    {input_path} (specified)")
    else:
        log("  Input:    (auto-discovering most recent activemembers file...)")
        input_path = find_input_file()
        log(f"  Input:    {input_path}")

    log(f"  Dry run:  {args.dry_run}")
    log(f"  No Drive: {args.no_drive}")
    log("")

    log("Loading region lookup...")
    load_zip_region_lookup()

    log("Loading input CSV...")
    rows = load_csv(input_path)

    log("Loading GlueUp auth token...")
    auth = glueup_auth()

    log("Fetching all GlueUp members...")
    all_glueup = get_all_glueup_members(auth)
    glueup_by_email, glueup_by_member_id = build_glueup_index(all_glueup)

    log("Processing members...")
    contact_rows, membership_rows, comparison_data = process_rows(
        rows, glueup_by_email, glueup_by_member_id, dry_run=args.dry_run
    )

    if args.dry_run:
        log("Dry run — no files written.")
        return

    contact_path    = f"contact_import_{RUN_TS_STR}.xlsx"
    membership_path = f"membership_import_{RUN_TS_STR}.xlsx"
    comparison_path = f"comparison_report_{RUN_TS_STR}.xlsx"
    duplicate_path  = f"duplicate_report_{RUN_TS_STR}.xlsx"
    log_path        = f"run_log_{RUN_TS_STR}.txt"

    log("")
    log("Writing output files...")
    write_import_xlsx(contact_rows,    contact_path,    CONTACT_FIELDNAMES)
    write_import_xlsx(membership_rows, membership_path, MEMBERSHIP_FIELDNAMES)
    write_comparison_report(comparison_data, comparison_path)

    log("Writing duplicate report...")
    write_duplicate_report(all_glueup, duplicate_path)

    save_log(log_path)
    log(f"  Written: {log_path}")

    if not args.no_drive:
        log("")
        upload_to_drive([contact_path, membership_path, comparison_path, duplicate_path, log_path])

    log("")
    log("Done.")

if __name__ == "__main__":
    main()
