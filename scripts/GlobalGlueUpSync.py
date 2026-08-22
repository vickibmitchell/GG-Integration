#!/usr/bin/env python3
"""
GlobalGlueUpSync.py
ICF Washington State Chapter — GlueUp Member Sync
Version: 2.3.0

CHANGELOG
---------
v2.3.0  2026-08-22
  - New Contact-scoped custom field "ICF Global Email" (internal name
    icfglobalemail) added to CONTACT_FIELDNAMES and always populated in
    build_contact_row() with the current ICF Global email for every
    processed row. This field always mirrors ICF Global's email, whatever
    it is this run.
  - Root-cause fix for duplicate Contact creation: GlueUp's Contacts Import
    matches/dedupes incoming rows by the "Email" column. Previously the
    script sent ICF Global's email into "Email" for every row, so any
    member whose ICF Global email differed from what GlueUp had on file
    would fail to match the existing Contact and create a new orphan
    Contact (no Membership attached) instead of updating the real one.
    build_contact_row() now only sends effective_email into "Email" when
    the contact is genuinely NEW (glueup_rec is None); for an existing/
    matched contact it resends GlueUp's own on-file email (via
    extract_glueup_field(glueup_rec, "emailAddress")) so the identifying
    "Email" field is never changed by the sync, while "ICF Global Email"
    still tracks any change on the ICF Global side. GlueUp's primary
    Email/identifier is otherwise only ever set once, at true NEW-member
    creation.
  - Dropped member detection now also appends a minimal Contact-import row
    (First Name, Last Name, existing GlueUp Email, ICF Global Member ID,
    "ICF Global Member Status" = "Expired", all else blank) to
    contact_rows for every dropped member, so contact_import_*.xlsx sets
    the Contact's status to Expired at the same time dropped_members_*.xlsx
    flags the Membership for manual cancellation in GlueUp Admin UI. This
    previously only happened once some other field changed and triggered a
    later CHANGED/RENEWAL row for that contact -- Status was never
    explicitly set to Expired at drop time.

v2.2.0  2026-06-22
  - GlueUp Contact Settings confirmed all ICF Global * custom fields,
    including "ICF Global Membership End Date", are Contact-scoped (GlueUp
    has no custom fields on the Membership object). Verified via live test
    import (Tommy Price, member 9066454) with no duplication and a correct
    field update.
  - RENEWAL status no longer generates a separate renewals_*.xlsx worklist
    or requires manual GlueUp Admin UI date edits. build_contact_row() now
    accepts is_renewal and populates "ICF Global Membership End Date"
    directly when True. Start Date / Restart Date / Type remain blank on
    renewal by design, to preserve legacy values in those fields.
  - process_rows() RENEWAL branch: contact import only (no membership
    import, no renewal_rows list). RENEWAL remains a distinct status in
    the comparison report and run log for visibility/audit trail.
  - Removed renewals_*.xlsx from output files, Drive upload list, and the
    import-ready notification email. Email step numbering and wording
    updated accordingly.

v2.1.2  2026-06-15
  - Added SCRIPT_VERSION constant. Notification email footer now references
    the actual running version dynamically instead of hardcoded "v2.0.0".

v2.1.1  2026-06-09
  - build_membership_row: "Membership End Date" now set to 10 years from
    run date (matching GlueUp's new 10-year membership term configuration)
    instead of the ICF Global expiry date. "ICF Global Membership End Date"
    continues to reflect the actual ICF expiry date.

v2.1.0  2026-06-06
  - Removed Make scenario row filter (Status = Active OR expiry > now-365) from
    sync logic; ICF Global API now correctly returns only active members.
  - Sync logic redesigned: NEW members get contact + membership import;
    RENEWED members (expiry date advanced) get contact import + renewals report
    for manual GlueUp UI update; CHANGED members (other fields only) get contact
    import only. Membership import is no longer used for existing members.
  - Added dropped member detection: GlueUp records with an ICF Member ID not
    present in the current ICF Global export are flagged in dropped_members_*.xlsx
    for manual membership cancellation in GlueUp Admin UI.
  - New output file: renewals_*.xlsx — same structure as membership import,
    contains members whose expiration date has advanced; for manual processing.
  - New output file: dropped_members_*.xlsx — replaces the old process_rows
    DROPPED logic; now generated directly from GlueUp index comparison.
  - Comparison report Summary tab updated: shows all 6 statuses (NEW, RENEWAL,
    CHANGED, SAME, SKIPPED, DROPPED) with color coding and Action Required column.
  - Comparison report: added Renewals and Dropped tabs alongside Detail/Changed.
  - Removed: is_active, new_expired, expiry_passed, membership_triggered logic.
  - Notification email updated to reflect new 6-file output and action steps.

v2.0.1  2026-05-15
  - Updated import-ready notification email: steps 2 and 3 now correctly
    reference contact_import_*.xlsx (Contacts → Import) and
    membership_import_*.xlsx (Memberships → Import) separately, and
    reference the "ICF Global GlueUp Sync Operations" doc for full instructions.

v2.0.0  2026-05-15
  - Cloud Run migration: replaced OAuth 2.0 browser flow in _get_drive_service()
    with Application Default Credentials (ADC). On Cloud Run the service account
    supplies credentials automatically; locally use 'gcloud auth application-default login'.
  - GlueUp auth: replaced token-file approach (glueup_token.json) with direct
    HMAC-SHA256 authentication. glueup_auth() now calls the GlueUp session
    endpoint and returns a live token — no pre-generated token file required.
  - Secret Manager: added _load_secrets_from_secret_manager() — when
    GLUEUP_USE_SECRET_MANAGER=1 is set (Cloud Run), GLUEUP_SK is fetched
    from GCP Secret Manager at startup, overriding the hardcoded fallback.
  - Drive download: added download_latest_icf_file_from_drive() — when no
    local activemembers CSV is found, the script downloads the most recent
    one from the Drive Inbound folder (DRIVE_INBOUND_FOLDER_ID).
  - Drive move: after successful upload, the inbound CSV is moved to the
    Processed folder (DRIVE_PROCESSED_FOLDER_ID) so it is not reprocessed.
  - Notifications: added SendGrid email notifications — 'Import Ready' after
    a successful run, 'No Inbound File' warning when the Inbound folder is empty.
  - DRIVE_SCOPES broadened from drive.file to drive for Shared Drive access.
  - Added DRIVE_INBOUND_FOLDER_ID and DRIVE_PROCESSED_FOLDER_ID constants.
  - Added GLUEUP_MD5_PW and GLUEUP_EMAIL constants for headless auth.
  - Removed DRIVE_TOKEN_FILE, DRIVE_CREDENTIALS, TOKEN_FILE references.

v1.19  (previous release — local machine only)
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

import requests as _requests

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip3 install openpyxl --break-system-packages")
    sys.exit(1)

# Google Drive upload — uses Application Default Credentials (ADC)
# On Cloud Run: service account credentials supplied automatically.
# Locally: run 'gcloud auth application-default login' once.
_DRIVE_LIBS_AVAILABLE = False
try:
    from googleapiclient.discovery import build as _gdrive_build
    from googleapiclient.http import MediaFileUpload as _MediaFileUpload
    from google.auth import default as _google_auth_default
    _DRIVE_LIBS_AVAILABLE = True
except ImportError:
    pass  # reported at upload time if Drive upload is attempted

# ─── Configuration ────────────────────────────────────────────────────────────

SCRIPT_VERSION            = "2.3.0"

GLUEUP_BASE_URL           = "https://api-services.glueup.com"
GLUEUP_ORG_ID             = "7912"
GLUEUP_PK                 = "icfwshts"
GLUEUP_SK                 = os.environ.get("GLUEUP_SK", "MF4CAQACEADAzyLnyJdLTnvVextU0XMCAwEAAQIQAKT41snxxRoPfXb0gguT2QIIDoAWHftn8h0CCA1L/8fNanzPAggAksHNF6ZpZQIIARa17LAfBfUCCAP67vZiAQ55")
GLUEUP_MD5_PW             = os.environ.get("GLUEUP_MD5_PW", "70bdb05ade647079069cfebed391758a")
GLUEUP_EMAIL              = os.environ.get("GLUEUP_EMAIL", "technology@icfwashingtonstate.org")
GLUEUP_MEMBERSHIP_TYPE_ID = 37600

SHADOW_EMAIL_DOMAIN       = "members.icfwashingtonstate.org"

# Google Drive folders
DRIVE_INBOUND_FOLDER_ID   = "1j-jfB8MvP7kaIagDO9A3Dy8RxEkFqXXP"   # Data-Transfer > Inbound
DRIVE_PROCESSED_FOLDER_ID = "1HSfnpqX8KeuJrmVsLEP7K5k2HsyUK98E"   # Data-Transfer > Processed
DRIVE_FOLDER_ID           = "1BQ53mlYzkl3N5wz6AiTZm1kDpibempJy"   # Data-Transfer > Sync
DRIVE_SCOPES              = ["https://www.googleapis.com/auth/drive"]

# Notification
NOTIFY_FROM = "technology@icfwashingtonstate.org"
NOTIFY_TO   = "GlueUpNotifiers@icfwashingtonstate.org"

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

# ─── Secret Manager (Cloud Run only) ─────────────────────────────────────────

def _load_secrets_from_secret_manager():
    """Override GLUEUP_SK from GCP Secret Manager when running on Cloud Run."""
    global GLUEUP_SK
    if os.environ.get("GLUEUP_USE_SECRET_MANAGER") != "1":
        return
    try:
        from google.cloud import secretmanager
        project_id = os.environ.get("GOOGLE_CLOUD_PROJECT", "gg-sync-492000")
        client = secretmanager.SecretManagerServiceClient()

        def _get(secret_id):
            name = f"projects/{project_id}/secrets/{secret_id}/versions/latest"
            return client.access_secret_version(request={"name": name}).payload.data.decode().strip()

        print("Loading secrets from GCP Secret Manager...")
        GLUEUP_SK = _get("glueup-private-key")
        print("  ✓ Secrets loaded.")
    except Exception as e:
        print(f"WARNING: Could not load secrets from Secret Manager: {e}. Using hardcoded fallbacks.")

_load_secrets_from_secret_manager()


# ─── Input file discovery ─────────────────────────────────────────────────────

def find_input_file():
    """
    Find the most recently modified file with 'activemembers' in its name
    in the current directory. Returns the path, or raises FileNotFoundError.
    """
    candidates = glob.glob("*activemembers*")
    if not candidates:
        raise FileNotFoundError("No local activemembers file found.")
    candidates.sort(key=os.path.getmtime, reverse=True)
    chosen = candidates[0]
    if len(candidates) > 1:
        log(f"  Found {len(candidates)} activemembers files. Using most recently modified: {chosen}")
        for f in candidates[1:]:
            log(f"    (ignored) {f}")
    return chosen


def _get_drive_service():
    """Return authenticated Drive service using Application Default Credentials."""
    if not _DRIVE_LIBS_AVAILABLE:
        log("ERROR: Google API libraries not installed.")
        return None
    try:
        creds, _ = _google_auth_default(scopes=DRIVE_SCOPES)
        return _gdrive_build("drive", "v3", credentials=creds)
    except Exception as e:
        log(f"ERROR: Could not build Drive service: {e}")
        return None


def download_latest_icf_file_from_drive():
    """
    Find the most recent activemembers_*.csv in the Drive Inbound folder
    and download it to /tmp/. Returns (local_path, drive_file_id).
    """
    log("No local activemembers file found — searching Google Drive Inbound folder...")
    service = _get_drive_service()
    if service is None:
        raise FileNotFoundError("Drive service unavailable — cannot download inbound file.")

    query = (
        f"'{DRIVE_INBOUND_FOLDER_ID}' in parents "
        f"and name contains 'activemembers' "
        f"and mimeType != 'application/vnd.google-apps.folder' "
        f"and trashed = false"
    )
    results = service.files().list(
        q=query,
        fields="files(id, name, createdTime)",
        orderBy="createdTime desc",
        pageSize=5,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    files = results.get("files", [])
    if not files:
        raise FileNotFoundError(
            "No activemembers CSV found in Google Drive Inbound folder.\n"
            f"Folder ID: {DRIVE_INBOUND_FOLDER_ID}\n"
            "Run the Make 'Get Active Members' scenario first."
        )

    latest    = files[0]
    file_id   = latest["id"]
    file_name = latest["name"]
    local_path = f"/tmp/{file_name}"

    log(f"  Found: {file_name} (id: {file_id})")
    log(f"  Downloading to {local_path}...")

    import io
    from googleapiclient.http import MediaIoBaseDownload
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    with open(local_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

    log(f"  ✓ Downloaded: {file_name}")
    return local_path, file_id


def move_inbound_file_to_processed(file_id, file_name):
    """Move the processed CSV from Inbound to Processed folder."""
    log(f"\nMoving {file_name} to Processed folder...")
    service = _get_drive_service()
    if service is None:
        log("  ⚠  Could not move file — Drive service unavailable.")
        return
    try:
        service.files().update(
            fileId=file_id,
            addParents=DRIVE_PROCESSED_FOLDER_ID,
            removeParents=DRIVE_INBOUND_FOLDER_ID,
            fields="id, parents",
            supportsAllDrives=True,
        ).execute()
        log("  ✓ Moved to Data-Transfer > Processed.")
    except Exception as e:
        log(f"  ⚠  Could not move inbound file: {e}")
        log(f"  Move manually — File ID: {file_id}")


# ─── SendGrid Notifications ───────────────────────────────────────────────────

def _get_sendgrid_api_key():
    """Fetch SendGrid API key from Secret Manager."""
    from google.cloud import secretmanager
    project_id = os.environ.get("GOOGLE_CLOUD_PROJECT", "gg-sync-492000")
    client = secretmanager.SecretManagerServiceClient()
    name = f"projects/{project_id}/secrets/sendgrid-api-key/versions/latest"
    return client.access_secret_version(request={"name": name}).payload.data.decode().strip()


def _send_email(subject, body):
    """Send a plain-text email via SendGrid."""
    api_key = _get_sendgrid_api_key()
    payload = {
        "personalizations": [{"to": [{"email": NOTIFY_TO}]}],
        "from": {"email": NOTIFY_FROM, "name": "ICF WA GlueUp Sync"},
        "subject": subject,
        "content": [{"type": "text/plain", "value": body}],
    }
    resp = _requests.post(
        "https://api.sendgrid.com/v3/mail/send",
        json=payload,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        timeout=30,
    )
    resp.raise_for_status()


def send_import_ready_notification(folder_link, contact_count, membership_count,
                                   dropped_count, duplicate_count):
    """Send 'Import Ready' email after a successful sync run."""
    log("\nSending import-ready notification email...")
    date_display = RUN_TS.strftime("%B %d, %Y %I:%M %p")
    subject = f"GlueUp Import Ready — {date_display}"
    body = f"""GlueUp Import Ready — {date_display}

The weekly ICF Global → GlueUp member sync has completed successfully.

SUMMARY
-------
  New members:       {contact_count} contact rows / {membership_count} membership rows
  Dropped members:   {dropped_count} (manual cancellation required)
  Duplicates found:  {duplicate_count}

ACTION REQUIRED
---------------
1. Open the Sync output folder in Google Drive:
   {folder_link}

2. Upload contact_import_*.xlsx via:
   GlueUp Admin UI → Contacts → Import
   (Covers new members' contact fields, renewals' updated End Date,
   and other changed fields — all in one import.)

3. Upload membership_import_*.xlsx via:
   GlueUp Admin UI → Memberships → Import
   (New members only — do NOT use for renewals)

4. Review dropped_members_*.xlsx and cancel the GlueUp membership
   for each listed member via GlueUp Admin UI.

5. Review duplicate_report_*.xlsx and cancel RED-flagged duplicate
   membership records in GlueUp Admin UI.

See the "ICF Global GlueUp Sync Operations" doc for full instructions.

—
Sent automatically by GlobalGlueUpSync v{SCRIPT_VERSION} running on Google Cloud Run.
"""
    try:
        _send_email(subject, body)
        log(f"  ✓ Notification sent to {NOTIFY_TO}.")
    except Exception as e:
        log(f"  ⚠  Notification email failed: {e}")


def send_no_inbound_file_warning():
    """Send warning email when no inbound file is found."""
    log("\nSending no-inbound-file warning email...")
    date_display = RUN_TS.strftime("%B %d, %Y")
    subject = f"⚠ GlueUp Sync Warning — No Inbound File ({date_display})"
    body = f"""GlueUp Sync Warning — No Inbound File Found ({date_display})

The weekly ICF Global → GlueUp member sync ran but found no input file
in the Google Drive Inbound folder.

This likely means the Make 'Get Active Members' scenario did not run
as scheduled. Please check:

1. Make scenario status:
   make.com → Scenarios → ICF Chapter API — Get Active Members

2. Google Drive Inbound folder:
   Data-Transfer > Inbound

If Make did not run, trigger it manually and then re-run the Cloud Run
sync job from the GCP console:
   console.cloud.google.com → Cloud Run → Jobs → glueup-sync → Execute

—
Sent automatically by GlobalGlueUpSync v{SCRIPT_VERSION} running on Google Cloud Run.
"""
    try:
        _send_email(subject, body)
        log(f"  ✓ Warning sent to {NOTIFY_TO}.")
    except Exception as e:
        log(f"  ⚠  Could not send warning email: {e}")

# ─── GlueUp Authentication ────────────────────────────────────────────────────

def make_a_header(method="POST"):
    ts = str(int(time.time() * 1000))
    msg = method + GLUEUP_PK + "1.0" + ts
    d = hmac.new(GLUEUP_SK.encode(), msg.encode(), hashlib.sha256).hexdigest()
    return f"v=1.0;k={GLUEUP_PK};ts={ts};d={d}"

def glueup_auth():
    """
    Authenticate to GlueUp using HMAC-SHA256 and return auth headers dict.
    Generates a fresh session token on every run — no token file required.
    """
    log("Authenticating to GlueUp...")
    ts  = str(int(time.time() * 1000))
    msg = "POST" + GLUEUP_PK + "1.0" + ts
    d   = hmac.new(GLUEUP_SK.encode(), msg.encode(), hashlib.sha256).hexdigest()
    a_header = f"v=1.0;k={GLUEUP_PK};ts={ts};d={d}"

    payload = json.dumps({
        "email":      {"value": GLUEUP_EMAIL},
        "passphrase": {"value": GLUEUP_MD5_PW},
    }).encode()
    req = Request(
        f"{GLUEUP_BASE_URL}/v2/user/session",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "a": a_header,
            "User-Agent": "Mozilla/5.0",
        },
        method="POST",
    )
    try:
        with urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read())
    except HTTPError as e:
        log(f"ERROR: GlueUp authentication failed: HTTP {e.code} {e.reason}")
        sys.exit(1)

    token = data.get("value", {}).get("token")
    if not token:
        log(f"ERROR: No token in GlueUp auth response: {json.dumps(data)}")
        sys.exit(1)

    log("  ✓ GlueUp token obtained.")
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

def build_contact_row(row, glueup_rec, effective_email, has_real_email, is_renewal=False):
    """
    Build one row dict for the Contact import XLSX.
    All 34 columns in CONTACT_FIELDNAMES must be present (the 33 columns
    from sample_contact_import.xlsx, plus "ICF Global Email" added 2026-08-22).
    Blue (ignored) fields are left blank. Active fields are populated.

    is_renewal: when True, "ICF Global Membership End Date" is populated with
    the new ICF Global expiry date. GlueUp confirmed (Contact Settings ->
    Field Settings, 2026-06-22) that all ICF Global * fields, including this
    one, are Contact-scoped custom fields -- GlueUp has no custom fields on
    the Membership object. Contact import safely updates existing contacts
    (no duplication risk), so renewal dates no longer require a separate
    manual GlueUp Admin UI edit. Start Date / Restart Date / Type are left
    blank here intentionally -- only End Date is updated on renewal, per
    decision to preserve legacy values in those three fields.

    Email / ICF Global Email (added 2026-08-22): GlueUp's Contacts Import
    matches/dedupes incoming rows by the "Email" column -- NOT by "ICF
    Global Member ID". Sending ICF Global's email for an already-matched
    contact whose ICF Global email had changed was silently creating a new
    duplicate/orphan Contact instead of updating the real one. To prevent
    this, "Email" (GlueUp's real identifier field) is only ever set from
    ICF Global data at true NEW-member creation (glueup_rec is None); for
    an existing/matched contact, GlueUp's own on-file email is resent
    unchanged, so the sync can never itself cause a duplicate. The new
    "ICF Global Email" custom field always mirrors ICF Global's current
    email regardless, so staff can see when it differs from GlueUp's
    on-file "Email" and decide whether to update the latter manually.
    """
    shadow = not has_real_email
    if glueup_rec and not shadow:
        first, last = extract_glueup_name(glueup_rec)
    else:
        first = row.get(COL_FIRST_NAME, "").strip()
        last  = row.get(COL_LAST_NAME,  "").strip()

    credential = transform_credential(row.get(COL_CREDENTIAL, ""))
    tc_cred    = transform_credential(row.get(COL_TC_CRED, ""))

    # "Email" is GlueUp's matching/identifier field -- never resend ICF
    # Global's email for an already-existing contact (see docstring above).
    if glueup_rec is not None:
        contact_email = extract_glueup_field(glueup_rec, "emailAddress") or effective_email
    else:
        contact_email = effective_email

    return {
        "First Name":                               first,
        "Last Name":                                last,
        "Address":                                  "",          # blue — ignored
        "City":                                     row.get(COL_CITY, "").strip(),
        "State/Province":                           expand_state(row.get(COL_STATE, "")),
        "Postal Code/Zip Code":                     row.get(COL_ZIP, "").strip(),
        "Email":                                    contact_email,
        "ICF Global Email":                         effective_email,
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
        "ICF Global Membership End Date":           row.get(COL_EXPIRY, "").strip() if is_renewal else "",  # populated on renewal only (Contact-scoped field, confirmed 2026-06-22)
        "ICF Global Membership Restart Date":       "",          # blue — ignored (belongs to membership)
        "ICF Global Membership Start Date":         "",          # blue — ignored (belongs to membership)
        "ICF Global Membership Type":               "",          # blue — ignored (belongs to membership)
    }

# Exact column order from sample_contact_import.xlsx, plus "ICF Global Email"
# (added 2026-08-22, custom field internal name icfglobalemail) — 34 columns
# — all must be present.
CONTACT_FIELDNAMES = [
    "First Name", "Last Name", "Address", "City", "State/Province",
    "Postal Code/Zip Code", "Email", "ICF Global Email", "Phone", "Company",
    "Title/Position", "Volunteer Role", "Findable", "Directory Listing Text",
    "Coach Industry", "Coaching Specialization", "Has Email",
    "ICF Chapter Start Date", "ICF Credential", "ICF Credential Award Date",
    "ICF Credential Expire Date", "ICF Global Auto Renewal",
    "ICF Global Member ID", "ICF Global Member Status",
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
        "Membership End Date":                          (RUN_TS + datetime.timedelta(days=3650)).strftime("%m/%d/%Y"),  # 10-year term
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

def write_dropped_members_report(dropped_rows, path):
    """
    Write dropped_members XLSX.
    These are members present in GlueUp with an ICF Member ID that no
    longer appears in the current ICF Global export.
    Action: cancel their membership in GlueUp Admin UI.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dropped Members"

    col_headers = ["ICF Member ID", "GlueUp ID", "First Name", "Last Name",
                   "Email", "Membership End Date", "Action"]
    last_col = openpyxl.utils.get_column_letter(len(col_headers))

    _hdr_cell(ws, 1, 1,
              f"Dropped Members — {RUN_TS.strftime('%Y-%m-%d %H:%M')} "
              f"— {len(dropped_rows)} member(s) require manual cancellation",
              FILL_RED, FONT_WHITE_BOLD)
    ws.merge_cells(f"A1:{last_col}1")

    for col, h in enumerate(col_headers, start=1):
        _hdr_cell(ws, 2, col, h, FILL_HEADER, FONT_WHITE_BOLD)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 24

    for row_num, row in enumerate(dropped_rows, start=3):
        for col, field in enumerate(col_headers, start=1):
            c = ws.cell(row=row_num, column=col, value=row.get(field, ""))
            c.fill = FILL_RED
            c.alignment = Alignment(vertical="top")

    if not dropped_rows:
        ws.cell(row=3, column=1, value="No dropped members this run.")

    wb.save(path)
    log(f"  Written: {path} ({len(dropped_rows)} dropped member(s))")


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
    status_fills = {
        "NEW":     FILL_GREEN,
        "RENEWAL": FILL_YELLOW,
        "CHANGED": PatternFill("solid", fgColor="DDEBF7"),  # light blue
        "DROPPED": FILL_RED,
    }
    status_notes = {
        "NEW":     "Add via contact_import + membership_import",
        "RENEWAL": "Update via contact_import (End Date auto-populated)",
        "CHANGED": "Update via contact_import only",
        "SAME":    "No action required",
        "SKIPPED": "Shadow email, no date data — no action",
        "DROPPED": "Cancel membership in GlueUp Admin UI",
    }
    ws_sum.column_dimensions["C"].width = 52
    _hdr_cell(ws_sum, 3, 3, "Action Required", FILL_GRAY, FONT_BOLD)
    for i, status in enumerate(["NEW", "RENEWAL", "CHANGED", "SAME", "SKIPPED", "DROPPED"], start=4):
        c = ws_sum.cell(row=i, column=1, value=status)
        if status in status_fills:
            c.fill = status_fills[status]
        ws_sum.cell(row=i, column=2, value=counts.get(status, 0))
        ws_sum.cell(row=i, column=3, value=status_notes.get(status, ""))

    # ── Detail / Renewals / Changed / Dropped tabs ──────────────────────────────────────────────
    field_labels = [f[0] for f in COMPARISON_FIELDS]
    headers = (["Status", "ICF Member ID", "Name", "Email"]
               + [f"{lbl} (ICF)" for lbl in field_labels]
               + [f"{lbl} (GlueUp)" for lbl in field_labels])

    for tab_name, filter_fn in [
        ("Detail",   lambda e: True),
        ("Renewals", lambda e: e["status"] == "RENEWAL"),
        ("Changed",  lambda e: e["status"] == "CHANGED"),
        ("Dropped",  lambda e: e["status"] == "DROPPED"),
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
            row_fill = (FILL_GREEN  if status == "NEW"     else
                        FILL_YELLOW if status == "RENEWAL" else
                        PatternFill("solid", fgColor="DDEBF7") if status == "CHANGED" else
                        FILL_RED    if status == "DROPPED" else None)
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
      contact_rows, membership_rows, dropped_rows, comparison_data

    Sync logic:
      NEW     — member in ICF Global, not in GlueUp
                -> contact import + membership import
      RENEWAL — member in both; expiration date has advanced
                -> contact import only, with "ICF Global Membership End Date"
                   populated with the new expiry date (Contact-scoped field;
                   no separate manual GlueUp UI edit needed as of 2026-06-22)
      CHANGED — member in both; other fields differ (no expiry advance)
                -> contact import only
      SAME    — no differences -> no output
      SKIPPED — shadow email + no date data -> no output
      DROPPED — member in GlueUp with ICF Member ID not present in current
                ICF Global export -> dropped_members report (cancel manually)
    """
    contact_rows    = []
    membership_rows = []
    comparison_data = []
    counts = {"total": 0, "skipped": 0, "new": 0, "renewals": 0,
              "changed_other": 0, "same": 0,
              "existing_real": 0, "existing_shadow": 0, "fallback": 0}

    # Track every ICF Member ID seen in this run — used for dropped detection
    icf_ids_seen = set()

    for row in rows:
        counts["total"] += 1
        member_id    = row.get(COL_MEMBER_ID, "").strip()
        raw_email    = row.get(COL_EMAIL, "").strip()
        name_preview = f"{row.get(COL_FIRST_NAME, '')} {row.get(COL_LAST_NAME, '')}".strip()

        if member_id:
            icf_ids_seen.add(member_id)

        # Step 1: Skip check
        skip, skip_msg = should_skip(row)
        if skip:
            log(f"  {skip_msg}")
            counts["skipped"] += 1
            comparison_data.append({
                "member_id": member_id,
                "name":      name_preview,
                "email":     make_shadow_email(member_id),
                "status":    "SKIPPED",
                "diffs":     [],
            })
            continue

        # Step 2: Effective email
        has_real_email  = bool(raw_email)
        effective_email = raw_email.lower() if has_real_email else make_shadow_email(member_id)

        # Step 3: GlueUp lookup — email first, member ID fallback
        glueup_rec = glueup_by_email.get(effective_email.lower())
        if not glueup_rec:
            counts["fallback"] += 1
            glueup_rec = glueup_by_member_id.get(str(member_id))

        # Step 4: Determine sync status
        is_new   = glueup_rec is None
        diffs    = compare_record(row, glueup_rec, effective_email)
        any_diff = any(d["differs"] for d in diffs)

        if is_new:
            status = "NEW"
            counts["new"] += 1
        elif any_diff:
            changed_fields = {d["field"] for d in diffs if d["differs"]}
            if "Membership Expiration Date" in changed_fields:
                # Renewal = expiration date has advanced (not just any difference)
                icf_expiry_str    = row.get(COL_EXPIRY, "").strip()
                glueup_expiry_str = next(
                    (d["glueup_value"] for d in diffs
                     if d["field"] == "Membership Expiration Date"), ""
                )
                is_renewal = False
                try:
                    icf_dt    = datetime.datetime.strptime(icf_expiry_str,    "%m/%d/%Y")
                    glueup_dt = datetime.datetime.strptime(glueup_expiry_str, "%m/%d/%Y")
                    is_renewal = icf_dt > glueup_dt
                except ValueError:
                    pass  # unparseable dates — treat as non-renewal CHANGED
                status = "RENEWAL" if is_renewal else "CHANGED"
            else:
                status = "CHANGED"

            if has_real_email:
                counts["existing_real"] += 1
            else:
                counts["existing_shadow"] += 1
            if status == "RENEWAL":
                counts["renewals"] += 1
            else:
                counts["changed_other"] += 1
        else:
            status = "SAME"
            counts["same"] += 1
            if has_real_email:
                counts["existing_real"] += 1
            else:
                counts["existing_shadow"] += 1

        comparison_data.append({
            "member_id": member_id,
            "name":      f"{row.get(COL_FIRST_NAME, '')} {row.get(COL_LAST_NAME, '')}".strip(),
            "email":     effective_email,
            "status":    status,
            "diffs":     diffs,
        })

        # Step 5: Build output rows
        if status == "NEW":
            contact_rows.append(
                build_contact_row(row, glueup_rec, effective_email, has_real_email)
            )
            membership_rows.append(
                build_membership_row(row, glueup_rec, effective_email, has_real_email)
            )

        elif status == "RENEWAL":
            # Contact import only — End Date populated directly (Contact-scoped
            # field, confirmed 2026-06-22). No membership import, no separate
            # manual worklist required.
            contact_rows.append(
                build_contact_row(row, glueup_rec, effective_email, has_real_email,
                                   is_renewal=True)
            )

        elif status == "CHANGED":
            # Contact import only
            contact_rows.append(
                build_contact_row(row, glueup_rec, effective_email, has_real_email)
            )

    # ── Dropped member detection ──────────────────────────────────────────────
    # Any GlueUp record with an ICF Member ID not seen in this run is dropped.
    # In addition to the dropped_members_*.xlsx manual-cancellation worklist
    # below, also queue a minimal Contact-import row so "ICF Global Member
    # Status" is set to "Expired" on the Contact record at drop time, via the
    # same contact_import_*.xlsx -> Contacts -> Import mechanism used
    # everywhere else (added 2026-08-22 -- previously this only happened
    # incidentally, if some other field later changed and triggered a
    # CHANGED/RENEWAL row for that contact).
    dropped_rows = []
    for rec in glueup_by_member_id.values():
        props  = rec.get("properties", {}) or {}
        icf_id = str(props.get("icfmemberid", "") or "").strip()
        if icf_id and icf_id not in icf_ids_seen:
            first     = rec.get("givenName", "") or ""
            last      = rec.get("familyName", "") or ""
            email_raw = rec.get("emailAddress") or {}
            email     = (email_raw.get("value", "") if isinstance(email_raw, dict)
                         else str(email_raw)).strip()
            glueup_id = str(rec.get("id", ""))
            exp_prop  = extract_glueup_field(rec, "icfglobalmembershipenddate")
            comparison_data.append({
                "member_id": icf_id,
                "name":      f"{first} {last}".strip(),
                "email":     email,
                "status":    "DROPPED",
                "diffs":     [],
            })
            dropped_rows.append({
                "ICF Member ID":       icf_id,
                "GlueUp ID":           glueup_id,
                "First Name":          first,
                "Last Name":           last,
                "Email":               email,
                "Membership End Date": exp_prop,
                "Action":              "Cancel membership in GlueUp Admin UI",
            })
            # Minimal Contact-import row: only set Name/Email (unchanged,
            # for matching)/ICF Global Member ID/Status -- everything else
            # left blank so nothing else on the Contact is touched.
            contact_row = {field: "" for field in CONTACT_FIELDNAMES}
            contact_row["First Name"]                 = first
            contact_row["Last Name"]                  = last
            contact_row["Email"]                      = email
            contact_row["ICF Global Member ID"]       = icf_id
            contact_row["ICF Global Member Status"]   = "Expired"
            contact_rows.append(contact_row)

    log("")
    log("── Run Summary ──────────────────────────────────────────────")
    log(f"  Total rows read:           {counts['total']}")
    log(f"  Skipped (no data):         {counts['skipped']}")
    log(f"  New members:               {counts['new']}")
    log(f"  Renewals (auto via contact import): {counts['renewals']}")
    log(f"  Changed (other fields):    {counts['changed_other']}")
    log(f"  Same (no change):          {counts['same']}")
    log(f"  Dropped from ICF Global:   {len(dropped_rows)}")
    log(f"  Contact rows to import:    {len(contact_rows)}")
    log(f"  Membership rows to import: {len(membership_rows)}")
    log("─────────────────────────────────────────────────────────────")

    return contact_rows, membership_rows, dropped_rows, comparison_data

# ─── Google Drive Upload ──────────────────────────────────────────────────────

def upload_to_drive(file_paths):
    """
    Upload output files to a new dated subfolder in the Sync Output Drive folder.

    Creates:  Sync_YYYYMMDD_HHMM  inside DRIVE_FOLDER_ID
    Uploads:  all files in file_paths into that subfolder

    Returns the folder webViewLink on success, or None on failure.
    """
    log("Uploading output files to Google Drive...")

    service = _get_drive_service()
    if service is None:
        log("  Drive upload skipped — authentication failed (see above).")
        log("  Output files are saved locally.")
        return None

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
        return None

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
    return folder_link


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

    # Resolve input file.
    # Priority: explicit argument > local CWD search > Drive download.
    drive_file_id   = None   # set when file came from Drive; used for move-to-processed
    drive_file_name = None

    if args.input_csv:
        input_path = args.input_csv
        log(f"  Input:    {input_path} (specified)")
    else:
        try:
            input_path = find_input_file()
            log(f"  Input:    {input_path}")
        except FileNotFoundError:
            try:
                input_path, drive_file_id = download_latest_icf_file_from_drive()
                drive_file_name = os.path.basename(input_path)
                log(f"  Input:    {input_path} (downloaded from Drive)")
            except FileNotFoundError as e:
                log(f"\n⚠  {e}")
                if not args.no_drive:
                    send_no_inbound_file_warning()
                log("\nNothing to process. Exiting.")
                sys.exit(0)

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
    contact_rows, membership_rows, dropped_rows, comparison_data = process_rows(
        rows, glueup_by_email, glueup_by_member_id, dry_run=args.dry_run
    )

    if args.dry_run:
        log("Dry run — no files written.")
        return

    contact_path    = f"contact_import_{RUN_TS_STR}.xlsx"
    membership_path = f"membership_import_{RUN_TS_STR}.xlsx"
    dropped_path    = f"dropped_members_{RUN_TS_STR}.xlsx"
    comparison_path = f"comparison_report_{RUN_TS_STR}.xlsx"
    duplicate_path  = f"duplicate_report_{RUN_TS_STR}.xlsx"
    log_path        = f"run_log_{RUN_TS_STR}.txt"

    log("")
    log("Writing output files...")
    write_import_xlsx(contact_rows,    contact_path,    CONTACT_FIELDNAMES)
    write_import_xlsx(membership_rows, membership_path, MEMBERSHIP_FIELDNAMES)
    write_dropped_members_report(dropped_rows, dropped_path)
    write_comparison_report(comparison_data, comparison_path)

    log("Writing duplicate report...")
    # Count duplicate groups for notification
    dupes = {}
    for raw in all_glueup:
        rec    = _unwrap(raw)
        props  = rec.get("properties", {}) or {}
        icf_id = str(props.get("icfmemberid", "") or "").strip()
        if icf_id:
            dupes.setdefault(icf_id, []).append(rec)
    duplicate_count = sum(1 for recs in dupes.values() if len(recs) > 1)
    write_duplicate_report(all_glueup, duplicate_path)

    save_log(log_path)
    log(f"  Written: {log_path}")

    if not args.no_drive:
        log("")
        folder_link = upload_to_drive(
            [contact_path, membership_path, dropped_path,
             comparison_path, duplicate_path, log_path]
        )

        # Move inbound CSV to Processed (Cloud Run path only)
        if drive_file_id:
            move_inbound_file_to_processed(drive_file_id, drive_file_name)

        # Send import-ready notification (Cloud Run path only)
        if drive_file_id and folder_link:
            send_import_ready_notification(
                folder_link      = folder_link,
                contact_count    = len(contact_rows),
                membership_count = len(membership_rows),
                dropped_count    = len(dropped_rows),
                duplicate_count  = duplicate_count,
            )

    log("")
    log("Done.")

if __name__ == "__main__":
    main()