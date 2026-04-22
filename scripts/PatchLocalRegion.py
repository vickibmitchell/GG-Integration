#!/usr/bin/env python3
"""
patch_local_region.py
ICF Washington State Chapter — One-Time Local Region Patch Utility

Pulls ALL active members from GlueUp, looks up their Local Region from
ZipCodes_Areas_CCC.xlsx using City + Zip, and writes a contact_import
XLSX containing only members whose Local Region is currently blank or
missing.

The output file is ready to upload directly via:
  GlueUp Admin UI → Contacts → Import

Usage:
  python3 patch_local_region.py
  python3 patch_local_region.py --zip-file /path/to/ZipCodes_Areas_CCC.xlsx
  python3 patch_local_region.py --dry-run    # print summary, no output file
  python3 patch_local_region.py --all        # include members who already have a region

Prerequisites:
  pip3 install openpyxl --break-system-packages
  A valid glueup_token.json in the same directory (run GetToken.py first).
"""

import argparse
import datetime
import hmac
import hashlib
import json
import os
import sys
import time
from urllib.request import urlopen, Request
from urllib.error import HTTPError, URLError

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip3 install openpyxl --break-system-packages")
    sys.exit(1)

# ─── Configuration ─────────────────────────────────────────────────────────────

GLUEUP_BASE_URL = "https://api-services.glueup.com"
GLUEUP_ORG_ID   = "7912"
GLUEUP_PK       = "icfwshts"
GLUEUP_SK       = "MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPfXb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfUCCAP67vZiAQ55"
TOKEN_FILE      = "glueup_token.json"

ZIP_LOOKUP_FILE = "ZipCodes_Areas_CCC.xlsx"

RUN_TS          = datetime.datetime.now()
RUN_DATE_MMDDYYYY = RUN_TS.strftime("%m/%d/%Y")
RUN_TS_STR      = RUN_TS.strftime("%Y%m%d_%H%M")

# ─── Logging ───────────────────────────────────────────────────────────────────

_log_lines = []

def log(msg):
    print(msg)
    _log_lines.append(msg)

# ─── GlueUp Auth & HTTP ────────────────────────────────────────────────────────

def make_a_header():
    ts  = str(int(time.time() * 1000))
    msg = "POST" + GLUEUP_PK + "1.0" + ts
    d   = hmac.new(GLUEUP_SK.encode(), msg.encode(), hashlib.sha256).hexdigest()
    return f"v=1.0;k={GLUEUP_PK};ts={ts};d={d}"

def load_auth():
    """Read token.json and return auth headers dict. Exits on failure."""
    if not os.path.exists(TOKEN_FILE):
        log(f"ERROR: {TOKEN_FILE} not found. Run: python3 GetToken.py")
        sys.exit(1)
    with open(TOKEN_FILE) as f:
        data = json.load(f)
    token = data.get("token") or data.get("value", {}).get("token")
    if not token:
        log(f"ERROR: Could not read token from {TOKEN_FILE}. Run: python3 GetToken.py")
        sys.exit(1)
    return {"token": token, "requestOrganizationId": GLUEUP_ORG_ID}

def glueup_post(endpoint, body_dict, auth):
    """POST to GlueUp API. Returns parsed JSON or raises."""
    headers = {
        "Content-Type": "application/json",
        "a": make_a_header(),
        "User-Agent": "Mozilla/5.0",
        **auth,
    }
    body = json.dumps(body_dict).encode()
    req  = Request(f"{GLUEUP_BASE_URL}{endpoint}", data=body, headers=headers, method="POST")
    try:
        with urlopen(req, timeout=30) as resp:
            return json.loads(resp.read())
    except HTTPError as e:
        if e.code == 401:
            log("ERROR: GlueUp returned 401 Unauthorized — token may be expired. Run: python3 GetToken.py")
            sys.exit(1)
        raise

# ─── GlueUp Data Retrieval ─────────────────────────────────────────────────────

def fetch_all_members(auth):
    """
    Page through /v2/membershipDirectory/members and return every record.
    The endpoint returns active members only by default.
    """
    log("Fetching all active members from GlueUp...")
    all_records = []
    offset, limit = 0, 100
    while True:
        try:
            result = glueup_post(
                "/v2/membershipDirectory/members",
                {"limit": limit, "offset": offset},
                auth,
            )
        except Exception as e:
            log(f"  WARNING: API error at offset {offset}: {e}")
            break

        batch = result.get("value", [])
        all_records.extend(batch)
        total = result.get("metadata", {}).get("pagination", {}).get("total", "?")
        log(f"  Fetched {len(all_records)} / {total} members...")

        if len(batch) < limit:
            break
        offset += limit

    log(f"  Done. {len(all_records)} total records retrieved.")
    return all_records

# ─── Record Field Extraction ───────────────────────────────────────────────────

def unwrap(raw):
    """
    List endpoint wraps each record as {"membership": {...}, "individualMember": {...}}.
    Unwrap to the individualMember dict.
    """
    return raw.get("individualMember", raw)

def get_email(rec):
    raw = rec.get("emailAddress") or {}
    return (raw.get("value", "") if isinstance(raw, dict) else str(raw)).strip()

def get_city(rec):
    addr = rec.get("address") or {}
    return str(addr.get("cityName") or addr.get("city") or rec.get("city") or "").strip()

def get_zip(rec):
    addr = rec.get("address") or {}
    return str(addr.get("zipCode") or rec.get("zipCode") or "").strip()

def get_name(rec):
    return rec.get("givenName", "").strip(), rec.get("familyName", "").strip()

def get_prop(rec, key):
    """Extract a custom property value from the properties dict."""
    props = rec.get("properties") or {}
    val   = props.get(key)
    if val is None:
        return ""
    if isinstance(val, dict):
        # single-select field: {"code": "...", "title": {...}}
        return str(val.get("code", "") or "").strip()
    if isinstance(val, (int, float)) and val > 1_000_000_000_000:
        try:
            return datetime.datetime.utcfromtimestamp(val / 1000).strftime("%m/%d/%Y")
        except Exception:
            pass
    return str(val).strip()

def get_current_local_region(rec):
    return get_prop(rec, "localregion")   # adjust key if yours differs

# ─── Zip → Region Lookup ───────────────────────────────────────────────────────

ZIP_REGION_LOOKUP = {}   # (city_lower, zip) → region
ZIP_ONLY_LOOKUP   = {}   # zip → region

def load_zip_lookup(path=None):
    global ZIP_REGION_LOOKUP, ZIP_ONLY_LOOKUP
    candidates = [
        path,
        ZIP_LOOKUP_FILE,
        os.path.join(os.path.dirname(__file__), ZIP_LOOKUP_FILE),
    ]
    xlsx = next((p for p in candidates if p and os.path.exists(p)), None)
    if not xlsx:
        log(f"WARNING: {ZIP_LOOKUP_FILE} not found — all members will get 'No Region'.")
        return
    try:
        wb    = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
        ws    = wb.active
        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 3:
                continue
            region = str(row[0] or "").strip()
            city   = str(row[1] or "").strip().lower()
            zip_   = str(row[2] or "").strip()
            if region and zip_:
                ZIP_ONLY_LOOKUP[zip_]          = region
                if city:
                    ZIP_REGION_LOOKUP[(city, zip_)] = region
                count += 1
        wb.close()
        log(f"  Loaded {count} zip/region entries from {xlsx}.")
    except Exception as e:
        log(f"WARNING: Could not load {ZIP_LOOKUP_FILE}: {e} — all members will get 'No Region'.")

def lookup_region(city, zip_code):
    city_key = (city or "").strip().lower()
    zip_key  = (zip_code or "").strip()
    return (
        ZIP_REGION_LOOKUP.get((city_key, zip_key))
        or ZIP_ONLY_LOOKUP.get(zip_key)
        or "No Region"
    )

# ─── Import Row Builder ────────────────────────────────────────────────────────

# All 33 columns required by the GlueUp contact import template.
# Fields not relevant to this patch are left blank — GlueUp treats blank
# cells as "no change" during import (only populated cells are updated).
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

def build_import_row(rec, region):
    """
    Build a minimal contact import row that sets only Local Region
    (plus the identifying fields GlueUp needs to match the record).
    """
    first, last = get_name(rec)
    email       = get_email(rec)
    city        = get_city(rec)
    zip_code    = get_zip(rec)
    icf_id      = get_prop(rec, "icfmemberid")
    has_email   = "no" if "@members.icfwashingtonstate.org" in email else "yes"

    row = {col: "" for col in CONTACT_FIELDNAMES}   # start blank
    row.update({
        "First Name":            first,
        "Last Name":             last,
        "Email":                 email,
        "City":                  city,
        "Postal Code/Zip Code":  zip_code,
        "Has Email":             has_email,
        "ICF Global Member ID":  icf_id,
        "ICF Global Import Date": RUN_DATE_MMDDYYYY,
        "Local Region":          region,
    })
    return row

# ─── XLSX Writer ───────────────────────────────────────────────────────────────

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)

def write_import_xlsx(rows, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "crm_contacts"

    # Header row
    for col_idx, col_name in enumerate(CONTACT_FIELDNAMES, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")

    # Data rows — yellow (CHANGED) to match the existing sync convention
    for row_idx, row_dict in enumerate(rows, start=2):
        for col_idx, col_name in enumerate(CONTACT_FIELDNAMES, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row_dict.get(col_name, ""))
            cell.fill = YELLOW_FILL

    # Auto-fit column widths (approximate)
    for col_idx, col_name in enumerate(CONTACT_FIELDNAMES, start=1):
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(col_idx)
        ].width = max(12, min(len(col_name) + 2, 40))

    wb.save(path)
    log(f"  Saved: {path}  ({len(rows)} rows)")

# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Patch Local Region for GlueUp members missing it.")
    parser.add_argument("--zip-file",  default=None,  help="Path to ZipCodes_Areas_CCC.xlsx")
    parser.add_argument("--dry-run",   action="store_true", help="Print summary but do not write output file")
    parser.add_argument("--all",       action="store_true", help="Include members who already have a Local Region")
    args = parser.parse_args()

    log("=" * 60)
    log("  patch_local_region.py")
    log(f"  Run: {RUN_TS.strftime('%Y-%m-%d %H:%M:%S')}")
    log("=" * 60)

    # 1. Load zip → region lookup
    log("\nLoading zip/region lookup...")
    load_zip_lookup(args.zip_file)

    # 2. Authenticate
    log("\nLoading GlueUp auth token...")
    auth = load_auth()
    log("  Token loaded OK.")

    # 3. Fetch all members
    log("")
    all_raw = fetch_all_members(auth)

    # 4. Process each member
    log("\nAnalysing members...")
    import_rows        = []
    skipped_has_region = 0
    skipped_no_region  = 0   # lookup returned 'No Region' — nothing useful to set
    count_total        = len(all_raw)

    for raw in all_raw:
        rec = unwrap(raw)

        city     = get_city(rec)
        zip_code = get_zip(rec)
        current  = get_current_local_region(rec)

        # Determine what region we would assign
        new_region = lookup_region(city, zip_code)

        # Skip members who already have a region (unless --all)
        if current and not args.all:
            skipped_has_region += 1
            continue

        # Skip members where the lookup can't find a region — nothing to set
        if new_region == "No Region":
            skipped_no_region += 1
            first, last = get_name(rec)
            log(f"  SKIP (no lookup match): {first} {last} | city={city!r} zip={zip_code!r}")
            continue

        import_rows.append(build_import_row(rec, new_region))

    # 5. Summary
    log(f"\n{'─'*60}")
    log(f"  Total GlueUp records:         {count_total}")
    log(f"  Already have a Local Region:  {skipped_has_region}")
    log(f"  No lookup match (skipped):    {skipped_no_region}")
    log(f"  Will be updated:              {len(import_rows)}")
    log(f"{'─'*60}")

    if not import_rows:
        log("\nNothing to do — no members need a Local Region update.")
        return

    if args.dry_run:
        log("\n--dry-run set. No file written.")
        log("Members that WOULD be updated:")
        for r in import_rows:
            log(f"  {r['First Name']} {r['Last Name']} | {r['Email']} → {r['Local Region']}")
        return

    # 6. Write output
    out_path = f"local_region_patch_{RUN_TS_STR}.xlsx"
    log(f"\nWriting contact import file...")
    write_import_xlsx(import_rows, out_path)

    log(f"""
Done.

Next step:
  Upload '{out_path}' via GlueUp Admin UI:
    Contacts → Import → choose file → follow the 5-step wizard
  Yellow rows = CHANGED (Local Region being set).
""")

if __name__ == "__main__":
    main()