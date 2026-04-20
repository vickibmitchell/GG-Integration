#!/usr/bin/env python3
"""
GlobalGlueUpSync.py
ICF Washington State Chapter — GlueUp Member Sync
Version: 1.0

Reads the ICF Global active-member CSV (produced by the Get Active Members
Make scenario), looks up each member in GlueUp by email then by ICF Member ID,
and produces five output files:
  - contact_import_YYYYMMDD_HHMM.csv
  - membership_import_YYYYMMDD_HHMM.csv
  - comparison_report_YYYYMMDD_HHMM.xlsx
  - duplicate_report_YYYYMMDD_HHMM.xlsx
  - run_log_YYYYMMDD_HHMM.txt

Usage:
  python3 GlobalGlueUpSync.py <file.csv>
  python3 GlobalGlueUpSync.py <file.csv> --no-drive
  python3 GlobalGlueUpSync.py <file.csv> --dry-run

See GlueUp_Sync_Design_Spec for full field and logic documentation.
"""

import argparse
import csv
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

# ─── Configuration ────────────────────────────────────────────────────────────

GLUEUP_BASE_URL     = "https://api-services.glueup.com"
GLUEUP_ORG_ID       = "7912"
GLUEUP_PK           = "icfwshts"
GLUEUP_SK           = "MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPfXb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfUCCAP67vZiAQ55"
GLUEUP_MEMBERSHIP_TYPE_ID = 37600

SHADOW_EMAIL_DOMAIN = "members.icfwashingtonstate.org"
TOKEN_FILE          = "token.json"
DRIVE_FOLDER_ID     = None   # Open Item 6 — set when Drive upload is re-enabled

RUN_TS              = datetime.datetime.now()
RUN_DATE_MMDDYYYY   = RUN_TS.strftime("%m/%d/%Y")
RUN_TS_STR          = RUN_TS.strftime("%Y%m%d_%H%M")

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
ZIP_REGION_LOOKUP   = {}   # populated by load_zip_region_lookup()
ZIP_ONLY_LOOKUP     = {}   # populated by load_zip_region_lookup()

# ─── Comparison fields (ICF Global col → GlueUp property key) ─────────────────
# Used to build the comparison report diff.
COMPARISON_FIELDS = [
    # (label, icf_csv_col, glueup_property, transform_fn_name)
    ("First Name",                  "First_Name",                   "givenName",                        "none"),
    ("Last Name",                   "Last_Name",                    "familyName",                       "none"),
    ("Email",                       "_effective_email",             "emailAddress",                     "none"),
    ("City",                        "City",                         "city",                             "none"),
    ("Zip",                         "Zip",                          "zipCode",                          "none"),
    ("Membership Expiration Date",  "Membership_Expiration_Date",   "icfglobalmembershipenddate",       "date"),
    ("ICF Credential",              "Flagship_Credential",          "icfcredential",                    "lower"),
    ("Credential Award Date",       "Credential_Award_Date",        "icfcredentialawarddate",           "date"),
    ("Credential Expire Date",      "Credential_Expire_Date",       "icfcredentialexpiredate",          "date"),
    ("TC Credential",               "ACTC_Credential",              "icfteamcoachingcredential",        "lower"),
    ("TC Credential Award Date",    "ACTC_Credential_Award_Date",   "icfteamcoachingcredentialaward",   "date"),
    ("TC Credential Expire Date",   "ACTC_Credential_Expire_Date",  "icfteamcoachingcredentialexpir",   "date"),
    ("Auto Renewal",                "Auto_Renewal",                 "icfglobalautorenewal",             "autorenewal"),
]

# ─── Logger ───────────────────────────────────────────────────────────────────

_log_lines = []

def log(msg):
    print(msg)
    _log_lines.append(msg)

def save_log(path):
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(_log_lines))

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
    """POST to GlueUp API. Returns parsed JSON value or raises on error."""
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
    if not items:
        return None
    return items[0]

def lookup_by_email(email, auth):
    """Search GlueUp members by email. Returns record dict or None."""
    try:
        result = glueup_post(
            "/v2/membershipDirectory/members",
            {
                "filter": [{"projection": "emailAddress", "operator": "eq", "values": [email]}],
                "limit": 5,
                "offset": 0,
            },
            auth,
        )
        return _extract_member_record(result)
    except (HTTPError, URLError, Exception) as e:
        log(f"  WARNING: GlueUp email lookup failed for {email}: {e}")
        return None

def lookup_by_member_id(icf_id, auth):
    """Search GlueUp members by icfmemberid custom field. Returns record dict or None."""
    try:
        result = glueup_post(
            "/v2/membershipDirectory/members",
            {
                "filter": [{"projection": "properties.icfmemberid", "operator": "eq", "values": [str(icf_id)]}],
                "limit": 5,
                "offset": 0,
            },
            auth,
        )
        return _extract_member_record(result)
    except (HTTPError, URLError, Exception) as e:
        log(f"  WARNING: GlueUp member-ID lookup failed for {icf_id}: {e}")
        return None

def get_all_glueup_members(auth):
    """
    Fetch ALL members from GlueUp for duplicate detection.
    Pages through with limit/offset. Returns list of all records.
    """
    log("  Fetching all GlueUp members for duplicate detection...")
    all_records = []
    offset = 0
    limit = 100
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
    "Chapter_Start_Date", "Membership_Join_Date", "Membership_Expiration_Date",
    "Credential_Award_Date", "ACTC_Credential_Award_Date",
]

def should_skip(row):
    """
    Skip shadow-email members that have no date data at all —
    they have nothing useful to import.
    Returns (True, reason) or (False, "").
    """
    if row.get("Email", "").strip():
        return False, ""
    # shadow email member — check for any date data
    for col in DATE_COLS_FOR_SKIP:
        if row.get(col, "").strip():
            return False, ""
    return True, f"SKIP: {row.get('First_Name','')} {row.get('Last_Name','')} (ID {row.get('Member_ID','')}) — shadow email, no date data"

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
    result = ZIP_REGION_LOOKUP.get((city_key, zip_key))
    if result:
        return result
    result = ZIP_ONLY_LOOKUP.get(zip_key)
    if result:
        return result
    return "No Region"

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
    "Member_ID", "First_Name", "Last_Name", "Email",
    "Membership_Expiration_Date", "Member_Type",
]

def load_csv(path):
    """
    Read input CSV. Returns list of row dicts.
    Exits if required columns are missing.
    """
    if not os.path.exists(path):
        log(f"ERROR: Input file not found: {path}")
        sys.exit(1)
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        missing = [c for c in REQUIRED_COLS if c not in headers]
        if missing:
            log(f"ERROR: Input CSV missing required columns: {missing}")
            log(f"  Found columns: {headers}")
            sys.exit(1)
        for i, row in enumerate(reader, start=2):
            if not any(v.strip() for v in row.values()):
                continue  # skip blank rows
            rows.append(row)
    log(f"  Loaded {len(rows)} data rows from {path}.")
    return rows

# ─── GlueUp record field extraction ──────────────────────────────────────────

def extract_glueup_name(rec):
    """Return (first_name, last_name) from a GlueUp member record."""
    first = (rec.get("givenName") or "").strip()
    last  = (rec.get("familyName") or "").strip()
    return first, last

def extract_glueup_field(rec, prop_key):
    """
    Extract a field value from a GlueUp record.
    Handles top-level fields and nested properties dict.
    Dates stored as timestamps (ms) are converted to MM/DD/YYYY.
    """
    # Top-level contact fields
    top_map = {
        "givenName":    rec.get("givenName", ""),
        "familyName":   rec.get("familyName", ""),
        "emailAddress": (rec.get("emailAddress") or {}).get("value", "") if isinstance(rec.get("emailAddress"), dict) else rec.get("emailAddress", ""),
        "city":         rec.get("city", "") or (rec.get("address") or {}).get("city", ""),
        "zipCode":      rec.get("zipCode", "") or (rec.get("address") or {}).get("zipCode", ""),
    }
    if prop_key in top_map:
        return str(top_map[prop_key] or "").strip()

    # Custom properties
    props = rec.get("properties", {}) or {}
    val = props.get(prop_key, "")
    if val is None:
        return ""
    # GlueUp stores dates as epoch-ms integers for custom date fields
    if isinstance(val, (int, float)) and val > 1_000_000_000_000:
        try:
            dt = datetime.datetime.utcfromtimestamp(val / 1000)
            return dt.strftime("%m/%d/%Y")
        except Exception:
            pass
    # GlueUp may store dates as YYYY-MM-DD strings
    if isinstance(val, str) and len(val) == 10 and val[4] == "-":
        try:
            dt = datetime.datetime.strptime(val, "%Y-%m-%d")
            return dt.strftime("%m/%d/%Y")
        except Exception:
            pass
    return str(val).strip()

def extract_glueup_import_date(rec):
    """Extract icfimportdate from GlueUp record, for duplicate detection."""
    return extract_glueup_field(rec, "icfimportdate")

# ─── Row Builders ─────────────────────────────────────────────────────────────

def build_contact_row(row, glueup_rec, effective_email, has_real_email):
    """Build one row dict for the Contact import CSV."""
    shadow = not has_real_email

    # Source of truth: GlueUp for existing real-email members
    if glueup_rec and not shadow:
        first, last = extract_glueup_name(glueup_rec)
    else:
        first = row.get("First_Name", "").strip()
        last  = row.get("Last_Name", "").strip()

    return {
        "First Name":             first,
        "Last Name":              last,
        "Email":                  effective_email,
        "Has Email":              "yes" if has_real_email else "no",
        "ICF Global Member ID":   row.get("Member_ID", "").strip(),
        "Phone":                  row.get("Phone", "").strip(),
        "City":                   row.get("City", "").strip(),
        "Postal Code/Zip Code":   row.get("Zip", "").strip(),
        "State/Province":         expand_state(row.get("State", "")),
        "Country":                row.get("Country", "").strip(),
        "Local Region":           get_local_region(row.get("City", ""), row.get("Zip", "")),
    }

CONTACT_FIELDNAMES = [
    "First Name", "Last Name", "Email", "Has Email",
    "ICF Global Member ID", "Phone", "City", "Postal Code/Zip Code",
    "State/Province", "Country", "Local Region",
]

def build_membership_row(row, glueup_rec, effective_email, has_real_email):
    """Build one row dict for the Membership import CSV."""
    shadow = not has_real_email

    # Name: GlueUp for existing real-email members, ICF Global for new/shadow
    if glueup_rec and not shadow:
        first, last = extract_glueup_name(glueup_rec)
    else:
        first = row.get("First_Name", "").strip()
        last  = row.get("Last_Name", "").strip()

    credential  = transform_credential(row.get("Flagship_Credential", ""))
    tc_cred     = transform_credential(row.get("ACTC_Credential", ""))

    return {
        "Membership Start Date":            "",
        "Membership End Date":              row.get("Membership_Expiration_Date", "").strip(),
        "Currency":                         "",
        "First Name":                       first,
        "Last Name":                        last,
        "Email":                            effective_email,
        "Postal Code/Zip Code":             row.get("Zip", "").strip(),
        "Address":                          "",
        "City":                             row.get("City", "").strip(),
        "Volunteer Role":                   "",
        "Coach Industry":                   "",
        "Coach Specialty":                  "",
        "Findable":                         "",
        "Directory Listing Text":           "",
        "ICF Credential":                   credential,
        "ICF Credential Award Date":        row.get("Credential_Award_Date", "").strip() if credential else "",
        "ICF Credential Expire Date":       row.get("Credential_Expire_Date", "").strip() if credential else "",
        "ICF Global Auto Renewal":          transform_auto_renewal(row.get("Auto_Renewal", "")),
        "ICF Global Member ID":             row.get("Member_ID", "").strip(),
        "ICF Global Membership End Date":   row.get("Membership_Expiration_Date", "").strip(),
        "ICF Global Membership Start Date": row.get("Membership_Join_Date", "").strip(),
        "ICF Global Membership Type":       "individual",
        "ICF Team Coaching Credential":     tc_cred,
        "ICF TC Credential Award Date":     row.get("ACTC_Credential_Award_Date", "").strip() if tc_cred else "",
        "ICF TC Credential Expire Date":    row.get("ACTC_Credential_Expire_Date", "").strip() if tc_cred else "",
        "ICF Global Member Type":           row.get("Member_Type", "").strip(),
        "ICF Global Import Date":           RUN_DATE_MMDDYYYY,
    }

MEMBERSHIP_FIELDNAMES = [
    "Membership Start Date", "Membership End Date", "Currency",
    "First Name", "Last Name", "Email", "Postal Code/Zip Code",
    "Address", "City", "Volunteer Role", "Coach Industry", "Coach Specialty",
    "Findable", "Directory Listing Text",
    "ICF Credential", "ICF Credential Award Date", "ICF Credential Expire Date",
    "ICF Global Auto Renewal", "ICF Global Member ID",
    "ICF Global Membership End Date", "ICF Global Membership Start Date",
    "ICF Global Membership Type", "ICF Team Coaching Credential",
    "ICF TC Credential Award Date", "ICF TC Credential Expire Date",
    "ICF Global Member Type", "ICF Global Import Date",
]

# ─── Comparison Logic ─────────────────────────────────────────────────────────

def normalize_for_compare(value, transform):
    """Normalize a value for comparison (strip, lowercase where needed)."""
    v = str(value or "").strip()
    if transform == "lower":
        return v.lower()
    if transform == "autorenewal":
        return "yes" if v.lower() == "yes" else "no"
    if transform == "date":
        # Normalize MM/DD/YYYY → YYYY-MM-DD for comparison
        try:
            return datetime.datetime.strptime(v, "%m/%d/%Y").strftime("%Y-%m-%d")
        except Exception:
            return v
    return v

def compare_record(row, glueup_rec, effective_email):
    """
    Compare ICF Global row against GlueUp record field by field.
    Returns list of dicts: [{field, icf_value, glueup_value, differs}]
    """
    diffs = []
    for label, icf_col, glueup_prop, transform in COMPARISON_FIELDS:
        if icf_col == "_effective_email":
            icf_raw = effective_email
        else:
            icf_raw = row.get(icf_col, "")

        glueup_raw = extract_glueup_field(glueup_rec, glueup_prop) if glueup_rec else ""

        icf_norm    = normalize_for_compare(icf_raw, transform)
        glueup_norm = normalize_for_compare(glueup_raw, transform)

        diffs.append({
            "field":        label,
            "icf_value":    str(icf_raw).strip(),
            "glueup_value": str(glueup_raw).strip(),
            "differs":      icf_norm != glueup_norm,
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

# ─── Excel Report Writing ─────────────────────────────────────────────────────

# Colour palette
FILL_GREEN  = PatternFill("solid", fgColor="C6EFCE")   # NEW
FILL_YELLOW = PatternFill("solid", fgColor="FFEB9C")   # CHANGED
FILL_GRAY   = PatternFill("solid", fgColor="F2F2F2")   # SAME / header
FILL_RED    = PatternFill("solid", fgColor="FFC7CE")   # diff cell
FILL_HEADER = PatternFill("solid", fgColor="1F3864")   # dark blue header
FILL_DUPE_KEEP   = PatternFill("solid", fgColor="C6EFCE")
FILL_DUPE_REVIEW = PatternFill("solid", fgColor="FFC7CE")

FONT_WHITE_BOLD = Font(bold=True, color="FFFFFF")
FONT_BOLD       = Font(bold=True)

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
    Write comparison_report XLSX.
    Tabs:
      Summary  — NEW / CHANGED / SAME / SKIPPED counts
      Detail   — every member, every comparison field, colour-coded
      Changed  — only CHANGED members
    comparison_data: list of dicts with keys:
      member_id, name, email, status, diffs (list from compare_record())
    """
    wb = openpyxl.Workbook()

    # ── Summary tab ──────────────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Summary"
    counts = {"NEW": 0, "CHANGED": 0, "SAME": 0, "SKIPPED": 0}
    for r in comparison_data:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    ws_sum.column_dimensions["A"].width = 20
    ws_sum.column_dimensions["B"].width = 12
    _hdr_cell(ws_sum, 1, 1, f"GlobalGlueUpSync — Comparison Report — {RUN_TS.strftime('%Y-%m-%d %H:%M')}", FILL_HEADER, FONT_WHITE_BOLD)
    ws_sum.merge_cells("A1:B1")
    _hdr_cell(ws_sum, 3, 1, "Status", FILL_GRAY, FONT_BOLD)
    _hdr_cell(ws_sum, 3, 2, "Count", FILL_GRAY, FONT_BOLD)
    status_fills = {"NEW": FILL_GREEN, "CHANGED": FILL_YELLOW, "SAME": None, "SKIPPED": None}
    for i, status in enumerate(["NEW", "CHANGED", "SAME", "SKIPPED"], start=4):
        ws_sum.cell(row=i, column=1, value=status).fill = status_fills.get(status) or PatternFill()
        ws_sum.cell(row=i, column=2, value=counts[status])

    # ── Detail tab ───────────────────────────────────────────────────────────
    ws_det = wb.create_sheet("Detail")
    field_labels = [f["field"] for f in comparison_data[0]["diffs"]] if comparison_data and comparison_data[0].get("diffs") else [c[0] for c in COMPARISON_FIELDS]
    headers = ["Status", "ICF Member ID", "Name", "Email"] + \
              [f"{lbl} (ICF)" for lbl in field_labels] + \
              [f"{lbl} (GlueUp)" for lbl in field_labels]

    for col, h in enumerate(headers, start=1):
        _hdr_cell(ws_det, 1, col, h, FILL_HEADER, FONT_WHITE_BOLD)
        ws_det.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18

    row_num = 2
    for entry in comparison_data:
        status = entry["status"]
        row_fill = FILL_GREEN if status == "NEW" else FILL_YELLOW if status == "CHANGED" else None

        base_vals = [status, entry["member_id"], entry["name"], entry["email"]]
        icf_vals    = [d["icf_value"]    for d in entry.get("diffs", [])]
        glueup_vals = [d["glueup_value"] for d in entry.get("diffs", [])]
        differs     = [d["differs"]      for d in entry.get("diffs", [])]

        all_vals = base_vals + icf_vals + glueup_vals
        for col, val in enumerate(all_vals, start=1):
            c = ws_det.cell(row=row_num, column=col, value=val)
            c.alignment = Alignment(vertical="top")
            if row_fill:
                c.fill = row_fill
            # Highlight individual diff cells in the value columns
            if col > 4:
                field_idx = (col - 5) % len(field_labels) if len(field_labels) else 0
                if field_idx < len(differs) and differs[field_idx]:
                    c.fill = FILL_RED
        row_num += 1

    # ── Changed tab ───────────────────────────────────────────────────────────
    ws_chg = wb.create_sheet("Changed")
    for col, h in enumerate(headers, start=1):
        _hdr_cell(ws_chg, 1, col, h, FILL_HEADER, FONT_WHITE_BOLD)
        ws_chg.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18

    row_num = 2
    for entry in comparison_data:
        if entry["status"] != "CHANGED":
            continue
        base_vals   = ["CHANGED", entry["member_id"], entry["name"], entry["email"]]
        icf_vals    = [d["icf_value"]    for d in entry.get("diffs", [])]
        glueup_vals = [d["glueup_value"] for d in entry.get("diffs", [])]
        differs     = [d["differs"]      for d in entry.get("diffs", [])]
        all_vals    = base_vals + icf_vals + glueup_vals
        for col, val in enumerate(all_vals, start=1):
            c = ws_chg.cell(row=row_num, column=col, value=val)
            c.alignment = Alignment(vertical="top")
            c.fill = FILL_YELLOW
            if col > 4:
                field_idx = (col - 5) % len(field_labels) if len(field_labels) else 0
                if field_idx < len(differs) and differs[field_idx]:
                    c.fill = FILL_RED
        row_num += 1

    wb.save(path)
    log(f"  Written: {path}")

def write_duplicate_report(all_glueup_members, path):
    """
    Write duplicate_report XLSX.
    Finds contacts with >1 active membership for the same icfmemberid.
    Uses icfimportdate to flag: most recent = KEEP, older = REVIEW/CANCEL.
    """
    # Group by icfmemberid
    by_icf_id = defaultdict(list)
    for rec in all_glueup_members:
        props = rec.get("properties", {}) or {}
        icf_id = str(props.get("icfmemberid", "") or "").strip()
        if icf_id:
            by_icf_id[icf_id].append(rec)

    dupes = {k: v for k, v in by_icf_id.items() if len(v) > 1}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Duplicates"

    headers = ["Action", "ICF Member ID", "GlueUp Membership ID", "Name", "Email", "Import Date", "Notes"]
    for col, h in enumerate(headers, start=1):
        _hdr_cell(ws, 1, col, h, FILL_HEADER, FONT_WHITE_BOLD)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 22

    _hdr_cell(ws, 1, 1, f"Duplicate Report — {RUN_TS.strftime('%Y-%m-%d %H:%M')}", FILL_HEADER, FONT_WHITE_BOLD)
    ws.merge_cells(f"A1:{openpyxl.utils.get_column_letter(len(headers))}1")

    # re-write actual headers on row 2
    for col, h in enumerate(headers, start=1):
        _hdr_cell(ws, 2, col, h, FILL_HEADER, FONT_WHITE_BOLD)

    row_num = 3
    dupe_count = 0
    for icf_id, recs in sorted(dupes.items()):
        # Sort by import date descending — most recent = keep
        def parse_import_date(r):
            d = extract_glueup_import_date(r)
            try:
                return datetime.datetime.strptime(d, "%m/%d/%Y")
            except Exception:
                return datetime.datetime.min

        recs_sorted = sorted(recs, key=parse_import_date, reverse=True)
        for i, rec in enumerate(recs_sorted):
            action = "KEEP" if i == 0 else "REVIEW — CANCEL IN GLUEUP"
            fill   = FILL_DUPE_KEEP if i == 0 else FILL_DUPE_REVIEW
            name   = f"{rec.get('givenName','')} {rec.get('familyName','')}".strip()
            email_raw = rec.get("emailAddress") or {}
            email  = email_raw.get("value", "") if isinstance(email_raw, dict) else str(email_raw)
            imp_date = extract_glueup_import_date(rec)
            glu_id   = str(rec.get("id", ""))
            notes    = "Most recent import date" if i == 0 else "Older record — cancel membership in GlueUp Admin UI"

            row_data = [action, icf_id, glu_id, name, email, imp_date, notes]
            for col, val in enumerate(row_data, start=1):
                c = ws.cell(row=row_num, column=col, value=val)
                c.fill = fill
                c.alignment = Alignment(vertical="top")
            row_num += 1
            dupe_count += 1

    if dupe_count == 0:
        ws.cell(row=3, column=1, value="No duplicates found.")

    wb.save(path)
    log(f"  Written: {path} ({len(dupes)} duplicate groups, {dupe_count} records flagged)")

# ─── Main Processing Loop ─────────────────────────────────────────────────────

def process_rows(rows, auth, dry_run=False):
    """
    Main per-record loop. Returns:
      contact_rows, membership_rows, comparison_data
    """
    contact_rows     = []
    membership_rows  = []
    comparison_data  = []

    counts = {"total": 0, "skipped": 0, "new": 0,
              "existing_real": 0, "existing_shadow": 0, "errors": 0}

    for row in rows:
        counts["total"] += 1
        member_id = row.get("Member_ID", "").strip()
        raw_email = row.get("Email", "").strip()

        # Step 1: Skip check
        skip, skip_msg = should_skip(row)
        if skip:
            log(f"  {skip_msg}")
            counts["skipped"] += 1
            comparison_data.append({
                "member_id": member_id,
                "name": f"{row.get('First_Name','')} {row.get('Last_Name','')}".strip(),
                "email": make_shadow_email(member_id),
                "status": "SKIPPED",
                "diffs": [],
            })
            continue

        # Step 2: Effective email
        has_real_email = bool(raw_email)
        effective_email = raw_email.lower() if has_real_email else make_shadow_email(member_id)

        # Steps 3-5: GlueUp lookup
        glueup_rec = None
        lookup_method = None

        if not dry_run:
            glueup_rec = lookup_by_email(effective_email, auth)
            if glueup_rec:
                lookup_method = "email"
            else:
                glueup_rec = lookup_by_member_id(member_id, auth)
                if glueup_rec:
                    lookup_method = "member_id"

        # Step 6: Status
        is_new = glueup_rec is None

        # Step 7: Comparison
        diffs = compare_record(row, glueup_rec, effective_email)
        any_diff = any(d["differs"] for d in diffs)

        if is_new:
            status = "NEW"
            counts["new"] += 1
        elif any_diff:
            status = "CHANGED"
            if has_real_email:
                counts["existing_real"] += 1
            else:
                counts["existing_shadow"] += 1
        else:
            status = "SAME"
            if has_real_email:
                counts["existing_real"] += 1
            else:
                counts["existing_shadow"] += 1

        name = f"{row.get('First_Name','')} {row.get('Last_Name','')}".strip()

        comparison_data.append({
            "member_id": member_id,
            "name":      name,
            "email":     effective_email,
            "status":    status,
            "diffs":     diffs,
        })

        # Steps 8-10: Build import rows (always, regardless of status)
        contact_rows.append(
            build_contact_row(row, glueup_rec, effective_email, has_real_email)
        )
        membership_rows.append(
            build_membership_row(row, glueup_rec, effective_email, has_real_email)
        )

    log("")
    log("── Run Summary ──────────────────────────────────────────────")
    log(f"  Total rows read:          {counts['total']}")
    log(f"  Skipped (no data):        {counts['skipped']}")
    log(f"  New (no GlueUp match):    {counts['new']}")
    log(f"  Existing — real email:    {counts['existing_real']}")
    log(f"  Existing — shadow email:  {counts['existing_shadow']}")
    log(f"  Errors:                   {counts['errors']}")
    log(f"  Contact rows to import:   {len(contact_rows)}")
    log(f"  Membership rows to import:{len(membership_rows)}")
    log("─────────────────────────────────────────────────────────────")

    return contact_rows, membership_rows, comparison_data

# ─── Entry Point ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="GlobalGlueUpSync — ICF Washington State GlueUp Sync")
    parser.add_argument("input_csv", help="Path to activemembers CSV from Make scenario")
    parser.add_argument("--no-drive", action="store_true", help="Skip Google Drive upload")
    parser.add_argument("--dry-run",  action="store_true", help="Process rows and print summary only; no output files written")
    args = parser.parse_args()

    log("=" * 61)
    log(f"  GlobalGlueUpSync.py  —  {RUN_TS.strftime('%Y-%m-%d %H:%M')}")
    log("=" * 61)
    log(f"  Input:    {args.input_csv}")
    log(f"  Dry run:  {args.dry_run}")
    log(f"  No Drive: {args.no_drive}")
    log("")

    # Load zip/region lookup
    log("Loading region lookup...")
    load_zip_region_lookup()

    # Load input CSV
    log("Loading input CSV...")
    rows = load_csv(args.input_csv)

    # Auth
    log("Loading GlueUp auth token...")
    auth = glueup_auth()

    # Process
    log("Processing members...")
    contact_rows, membership_rows, comparison_data = process_rows(rows, auth, dry_run=args.dry_run)

    if args.dry_run:
        log("Dry run — no files written.")
        return

    # Output paths
    contact_path    = f"contact_import_{RUN_TS_STR}.csv"
    membership_path = f"membership_import_{RUN_TS_STR}.csv"
    comparison_path = f"comparison_report_{RUN_TS_STR}.xlsx"
    duplicate_path  = f"duplicate_report_{RUN_TS_STR}.xlsx"
    log_path        = f"run_log_{RUN_TS_STR}.txt"

    # Write import CSVs
    log("")
    log("Writing output files...")
    write_csv(contact_rows,    contact_path,    CONTACT_FIELDNAMES)
    write_csv(membership_rows, membership_path, MEMBERSHIP_FIELDNAMES)

    # Write comparison report
    write_comparison_report(comparison_data, comparison_path)

    # Fetch all GlueUp members for duplicate detection
    log("Fetching GlueUp members for duplicate report...")
    all_glueup = get_all_glueup_members(auth)
    write_duplicate_report(all_glueup, duplicate_path)

    # Save log
    save_log(log_path)
    log(f"  Written: {log_path}")

    # Drive upload (stub — Open Item 6)
    if not args.no_drive:
        log("")
        log("Drive upload: skipped (Open Item 6 — Drive folder ID not configured).")

    log("")
    log("Done.")

if __name__ == "__main__":
    main()