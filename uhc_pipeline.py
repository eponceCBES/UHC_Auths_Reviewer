"""
United Health Care (UHC) Authorizations pipeline.

The source data lives inside an AI Builder / Copilot model response that the
Power Automate flow drops verbatim into the "JSON Payload" column.

What it does, per list item:

  1. PARSE  — read the "JSON Payload" column. That column holds the raw model
     response object; the authorization data we want is the JSON string in its
     `text` field. Parse it into structured fields.

  2. DEDUPE — group every item by authorization number (Title). The oldest item
     (lowest list ID) in each group is canonical; the rest are duplicates.
     By default duplicates are FLAGGED (Notes marker, or Lookup Status =
     "Duplicate" if that choice exists). Pass --delete-duplicates to delete them.

  3. ENRICH — for each canonical item, write the parsed fields back to SharePoint
     (Title, Member Name/DOB, Health Plan ID, dates, Overall Decision, Services,
     Notes, Journal Note) and look the member up against Consumers_formatted.csv
     (by Health Plan ID, then Name+DOB) to fill Client ID + Primary Care Manager
     and set Lookup Status = Matched / Not Found.

Scaling: the default run is INCREMENTAL — it fetches only items whose
LookupStatus is still blank (i.e. the flow created them but they haven't been
enriched). Processed rows are never re-fetched, so per-run cost tracks new
items, not total list size. A server-side Title check then catches any new item
that duplicates one processed in a prior run. Use --all for a full backfill /
audit scan, and index the Title + LookupStatus columns in the list settings
before the list passes ~5,000 rows.

Self-contained: all path resolution and Azure/Graph/consumer-CSV helpers are
defined below, with credential + consumer-CSV paths derived from this script's
own location under "Report Subscriptions" — so it runs on any synced machine.

Required Azure App permissions (admin-consented):
  - Microsoft Graph: Sites.ReadWrite.All  (Application)
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import logging
import os
import re
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import msal
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from cryptography.fernet import Fernet
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent


# ── Credentials / consumer export (resolved relative to this script) ─
def _find_report_subs() -> Path:
    """Locate the 'Report Subscriptions' folder this script lives under, so the
    same code works from any computer / OneDrive mount point. Resolution order:
      1. REPORT_SUBSCRIPTIONS_DIR environment variable, if set.
      2. Walk up the script's path looking for a 'Report Subscriptions' folder.
      3. Fall back to three levels up (…/Report Subscriptions/Special Programs/
         Authorizations/United Health Care/uhc_pipeline.py).
    """
    env = os.environ.get("REPORT_SUBSCRIPTIONS_DIR")
    if env and Path(env).is_dir():
        return Path(env)
    for parent in SCRIPT_DIR.parents:
        if parent.name.lower() == "report subscriptions":
            return parent
    return SCRIPT_DIR.parents[2]


REPORT_SUBS_DIR = _find_report_subs()

ENCRYPTION_DIR = REPORT_SUBS_DIR / "Azure" / "Azure Encryption"
KEY_FILE = ENCRYPTION_DIR / "secret.key"
CREDS_FILE = ENCRYPTION_DIR / "encrypted_creds.bin"

CONSUMERS_CSV = (
    REPORT_SUBS_DIR / "Misc" / "Consumers" / "Consumers_formatted.csv"
)

# Service Suspensions export (Misc). Used to flag, on each matched auth row, any
# service suspension that is currently active for that consumer. The suspensions
# table keys on CONSUMER_UUID (no Client ID), so we bridge it to the auth row's
# Client ID through the Consumers export (which carries both CONSUMER_UUID and
# CLIENT_ID).
SERVICE_SUSPENSIONS_CSV = (
    REPORT_SUBS_DIR / "Misc" / "Service Suspensions" / "Service_Suspensions_formatted.csv"
)
# Suspensions-CSV column names.
SUSP_CONSUMER_UUID_COL = "CONSUMER_UUID"
SUSP_SERVICE_COL = "SERVICE"
SUSP_START_COL = "START_DATE"
SUSP_END_COL = "END_DATE"
# Consumers-CSV column that bridges CONSUMER_UUID -> CLIENT_ID.
CSV_CONSUMER_UUID_COL = "CONSUMER_UUID"
# DISPLAY names of the two SP columns the join populates (resolved to internal
# names at runtime for the write path). The INTERNAL names are also pinned below
# for the report read paths (Excel/calendar). NOTE: SharePoint truncates internal
# names to 32 chars, so "Suspension Start Date" -> "...Start_x0020_Dat" (no "e").
SUSPENDED_SERVICE_COL_TITLE = "Service Suspended"
SUSPENSION_START_COL_TITLE = "Suspension Start Date"
SUSP_SVC_FIELD = "Service_x0020_Suspended"
SUSP_START_FIELD = "Suspension_x0020_Start_x0020_Dat"

# Service Plans export (Misc). Used to compare what UHC authorized against what
# the consumer's CURRENT ACTIVE service plan actually says (hours/week,
# days/week, service present at all). Keys on CONSUMER_UUID like the
# suspensions table, so it bridges to Client ID the same way.
SERVICE_PLANS_CSV = (
    REPORT_SUBS_DIR / "Misc" / "Service Plans" / "Service_Plans_formatted.csv"
)
# Service-Plans CSV column names.
PLAN_CONSUMER_UUID_COL = "CONSUMER_UUID"
PLAN_PROGRAM_COL = "CARE_PROGRAM_NAME"
PLAN_STATUS_COL = "CARE_PLAN_STATUS"
PLAN_SERVICE_COL = "SERVICE"
PLAN_PROVIDER_COL = "PROVIDER"
PLAN_ALLOC_STATUS_COL = "SERVICE_ALLOCATION_STATUS"
PLAN_ALLOC_START_COL = "SERVICE_ALLOCATION_START_DATE"
PLAN_ALLOC_END_COL = "SERVICE_ALLOCATION_END_DATE"
PLAN_UNITS_COL = "UNITS_ALLOCATED"
# The export is one row per allocation SCHEDULE / SUBSERVICE, and every one of
# those rows repeats the same UNITS_ALLOCATED — so the same allocation must be
# counted once or the weekly total comes out as a multiple of the real figure.
PLAN_SCHED_UUID_COL = "SERVICE_ALLOCATION_SCHEDULE_UUID"
PLAN_ALLOC_UUID_COL = "SERVICE_ALLOC_UUID"
PLAN_CARE_PLAN_UUID_COL = "CARE_PLAN_UUID"

# Service Allocation export (Misc). A service allocation stacks MULTIPLE
# schedules over time — WellSky shows them as "Schedule: 06/20/2026 -
# 02/28/2027  82.00 Units Weekly (Su: 16, Mo: 8, ...)" — and only the one whose
# window covers today is in effect. The Service Plans export carries the units
# and per-day hours for every schedule but NOT the schedule's own start/end
# dates, so summing its rows adds up superseded schedules (one consumer's
# Personal Care came out at 56 hrs/week instead of 20.5). This export has those
# dates, keyed by the same schedule UUID, so it's used to keep only the
# schedules in effect today.
SERVICE_ALLOCATION_CSV = (
    REPORT_SUBS_DIR / "Misc" / "Service Allocation" / "Service_Allocation_formatted.csv"
)
ALLOC_SCHED_UUID_COL = "SERVICE_ALLOC_SCHED_UUID"
ALLOC_SCHED_START_COL = "SERVICE_ALLOCATION_SCHEDULE_START_DATE"
ALLOC_SCHED_END_COL = "SERVICE_ALLOCATION_SCHEDULE_END_DATE"
PLAN_ALLOC_TYPE_COL = "ALLOCATION_TYPE"
PLAN_PRIOR_AUTH_COL = "PRIOR_AUTHORIZATION_NO"
PLAN_WEEKDAY_COLS = ("MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY")
PLAN_WEEKEND_COLS = ("SATURDAY", "SUNDAY")
PLAN_DAY_COLS = PLAN_WEEKDAY_COLS + PLAN_WEEKEND_COLS
# Only rows whose CARE_PROGRAM_NAME contains this are treated as UHC-funded.
# Covers "SCO - United 1 (Primary)", "SCO - United X (Auxiliary)",
# "LOC - SCO - United / GAFC" and "One Care - United Health Care 1 (Primary)".
PLAN_UHC_PROGRAM_TOKEN = "United"
# Weekly-hours difference below this is treated as agreement (rounding noise).
PLAN_HOURS_TOLERANCE = 0.25

# DISPLAY names of the two SP columns the plan check populates, plus their
# pinned INTERNAL names for the Excel/calendar read paths (SharePoint truncates
# internal names to 32 chars).
PLAN_DISCREPANCY_COL_TITLE = "Plan Discrepancy"
PLAN_DISCREPANCY_DETAIL_COL_TITLE = "Plan Discrepancy Detail"
PLAN_DISC_FIELD = "Plan_x0020_Discrepancy"
PLAN_DISC_DETAIL_FIELD = "Plan_x0020_Discrepancy_x0020_Det"
# Deterministic journal-note QA flag. Populated in build_field_body by
# validate_note(); a FLAG only (never blocks or edits the note). "OK" when the
# note passes all checks, otherwise a "; "-joined list of human-readable issues.
# DISPLAY name is what you create in the list UI; FIELD is the pinned internal
# name SharePoint derives from that display name. Created via Graph with an
# explicit internal name, so it is "NoteQA" (not the "Note_x0020_QA" the UI
# would have derived from the display name).
NOTE_QA_COL_TITLE = "Note QA"
NOTE_QA_FIELD = "NoteQA"
NOTE_QA_OK = "OK"
# The two non-finding summaries compare_auth_to_plan can return. Shared so the
# run-summary buckets match the written value exactly instead of re-deriving it
# from the string shape.
PLAN_SUMMARY_OK = "OK"
PLAN_SUMMARY_NO_PLAN = "No active plan"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SP_HOSTNAME = "centralboston.sharepoint.com"
SP_SITE_PATH = "/sites/DataManagement"
SP_SITE_URL = f"https://{SP_HOSTNAME}{SP_SITE_PATH}"

CSV_HEALTH_PLAN_COLS = ("ALT_ID_1", "ALT_ID_2")
CSV_CLIENT_ID_COL = "CLIENT_ID"
CSV_CASE_MANAGER_COL = "PRIMARY_CARE_MANAGER"
CSV_FIRST_NAME_COL = "FIRST_NAME"
CSV_LAST_NAME_COL = "LAST_NAME"
CSV_DOB_COL = "DOB"
# Optional columns used for the extra match fallbacks — absent columns are
# simply skipped (the CSV schema can vary between refreshes / machines).
CSV_MEDICAID_COLS = ("MEDICAID_NO", "MEDICAL_POLICY_NO")
CSV_ADDR_COL = "RES_ADDRESS1"
CSV_ZIP_COL = "RES_ZIP"

# ── Configuration ──────────────────────────────────────────────────
SP_LIST_ID = "9a466e43-6e94-458f-a611-ecf372914315"  # United Health Care Authorizations
SP_LIST_TITLE = "United Health Care Authorizations"   # used to build item display-form (PDF) links

JSON_PAYLOAD_FIELD = "JSONPayload"

# Inner-JSON key  ->  SP internal column name (plain text / number columns).
TEXT_FIELD_MAP = {
    "member_name": "MemberName",
    "health_plan_id": "HealthPlanID",
    "medicaid_id": "MedicaidID",
}
# Nested address sub-key  ->  SP internal column name. The payload's "address"
# object is { street_address, city, state, zip_code }.
ADDRESS_FIELD_MAP = {
    "street_address": "StreetAddress",
    "city": "City",
    "state": "State",
    "zip_code": "ZipCode",
}
CHANGE_TYPE_FIELD = "ChangeType"
# Display normalization for the Change Type column. The column must read
# "Initiate" — that is the word the UHC service authorization itself uses
# ("Request Type: HMK  Initiate"), so the list matches the source document.
# "New" is the old extraction wording and "Initiation" the first attempt at
# renaming it; both normalize to "Initiate". Keyed on the lowercased value, so
# the payload itself is left alone and only the column is renamed.
CHANGE_TYPE_DISPLAY = {"new": "Initiate", "initiation": "Initiate"}
# Inner-JSON key  ->  SP date column (dateOnly).
DATE_FIELD_MAP = {
    "member_dob": "MemberDOB",
    "review_date": "ReviewDate",
    "auth_period_start": "AuthPeriodStart",
    "auth_period_end": "AuthPeriodEnd",
}

OVERALL_DECISION_FIELD = "OverallDecision"
DECISION_MAP = {
    "approved": "Approved",
    "fully_approved": "Approved",
    "partially_approved": "Partially Approved",
    "partial": "Partially Approved",
    "partially approved": "Partially Approved",
    "denied": "Denied",
    "denial": "Denied",
    "rejected": "Denied",
}

LOOKUP_STATUS_FIELD = "LookupStatus"
PRIMARY_CARE_MANAGER_FIELD = "PrimaryCareManager"
DUPLICATE_STATUS = "Duplicate"
DUP_MARKER = "[DUPLICATE]"

LOG_FILE = SCRIPT_DIR / "uhc_run.log"

# Persisted last-run marker. Lets an incremental run pull only what arrived
# since the previous run and gate on the list's modified date WITHOUT opening
# any item payload (i.e. without touching the PII inside JSONPayload).
STATE_FILE = SCRIPT_DIR / "uhc_last_run.json"

# Fixed output name for the interactive calendar (overwritten every run, like
# the xlsx). Sits next to this script.
_HTML_CAL_NAME = "uhc_authorizations_calendar.html"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("uhc_logs")


# ── Last-run state (incremental gating) ────────────────────────────
def load_last_run() -> str | None:
    """ISO-8601 UTC timestamp of the previous successful run, or None if this is
    the first run / the marker is missing or corrupt."""
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        ts = (data or {}).get("last_run")
        return ts.strip() if isinstance(ts, str) and ts.strip() else None
    except (FileNotFoundError, json.JSONDecodeError, OSError, AttributeError):
        return None


def save_last_run(ts: str) -> None:
    try:
        STATE_FILE.write_text(json.dumps({"last_run": ts}), encoding="utf-8")
        log.info("[state] saved last_run=%s", ts)
    except OSError as e:
        log.warning("[state] could not write %s: %s", STATE_FILE, e)


def _utc_now_iso() -> str:
    """Current UTC time as an ISO-8601 'Z' string (matches Graph timestamps)."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ── Vendored: credentials / tokens / sessions ──────────────────────
def load_azure_creds() -> dict:
    return json.loads(Fernet(KEY_FILE.read_bytes()).decrypt(CREDS_FILE.read_bytes()))


def _msal_app(creds: dict) -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        client_id=creds["client_id"],
        authority=f"https://login.microsoftonline.com/{creds['tenant_id']}",
        client_credential=creds["client_secret"],
    )


def get_token(app: msal.ConfidentialClientApplication, scope: str) -> str:
    result = app.acquire_token_for_client(scopes=[scope])
    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed for {scope}: {result}")
    return result["access_token"]


# Connect/read timeouts (seconds). Without these a stalled socket blocks the
# whole script FOREVER — which is what caused the hourly task to wedge: a hung
# call never returns, the process never exits, and the next run kills it. A read
# timeout instead raises after 60s so the run fails fast and the next run retries.
_CONNECT_TIMEOUT = 15
_READ_TIMEOUT = 60
_DEFAULT_TIMEOUT = (_CONNECT_TIMEOUT, _READ_TIMEOUT)


class _TimeoutSession(requests.Session):
    """requests.Session that applies a default timeout to every request unless
    the caller passes its own. This is what guarantees no call can hang forever."""
    def request(self, *args, **kwargs):
        kwargs.setdefault("timeout", _DEFAULT_TIMEOUT)
        return super().request(*args, **kwargs)


def _new_session() -> requests.Session:
    """A session with a hard timeout on every call + automatic retry/backoff on
    transient failures (connection drops, 429/5xx throttling)."""
    s = _TimeoutSession()
    retry = Retry(
        total=4, connect=4, read=4, status=4,
        backoff_factor=2,  # 0s, 2s, 4s, 8s between retries
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "POST", "PATCH", "DELETE"]),
        respect_retry_after_header=True,
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


def graph_session(token: str) -> requests.Session:
    s = _new_session()
    s.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    })
    return s


def sp_session(token: str) -> requests.Session:
    s = _new_session()
    s.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
    })
    return s


def resolve_site_id(g: requests.Session) -> str:
    r = g.get(f"{GRAPH_BASE}/sites/{SP_HOSTNAME}:{SP_SITE_PATH}")
    r.raise_for_status()
    return r.json()["id"]


# ── Vendored: date + name parsing ──────────────────────────────────
def _norm_dob(s: str | None) -> str | None:
    """Normalize a date string to YYYY-MM-DD; returns None if unparseable."""
    if not s:
        return None
    s = str(s).strip()
    if not s:
        return None
    s = s.split("T")[0].split(" ")[0]
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y",
                "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


_RICH_TEXT_TAG = re.compile(r"<[^>]+>")


def _plain_text(v) -> str:
    """Flatten a SharePoint rich-text value to plain text. Enhanced-rich-text
    columns return their value wrapped in a <div class="ExternalClass...">…</div>
    (with a fresh GUID every read) and HTML-encode entities (& -> &amp;). Strip
    the tags and decode entities so the value round-trips and compares/display
    cleanly. Safe on plain strings (returns them stripped)."""
    if not v:
        return ""
    s = _RICH_TEXT_TAG.sub("", str(v))
    return html.unescape(s).strip()


def _norm_id(v) -> str | None:
    """Normalize an ID (Medicaid #, etc.) for exact comparison: strip a trailing
    '.0' float artifact, drop spaces/dashes, uppercase. Returns None if empty."""
    s = str(v or "").strip()
    if not s:
        return None
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"[\s\-]", "", s).upper()
    return s or None


# Common street-word abbreviations so "123 Main Street" and "123 Main St"
# produce the same key. Mapped to a single canonical short form.
_STREET_ABBR = {
    "street": "st", "str": "st",
    "avenue": "ave", "av": "ave",
    "road": "rd", "drive": "dr", "lane": "ln", "court": "ct",
    "place": "pl", "boulevard": "blvd", "blvd": "blvd",
    "terrace": "ter", "circle": "cir", "highway": "hwy", "parkway": "pkwy",
    "square": "sq", "trail": "trl",
    "north": "n", "south": "s", "east": "e", "west": "w",
    "northeast": "ne", "northwest": "nw", "southeast": "se", "southwest": "sw",
    "first": "1st", "second": "2nd", "third": "3rd", "fourth": "4th",
    "fifth": "5th",
}
# Tokens that start a unit/apartment designator — everything from here on in a
# street string is dropped so "10 Main St Apt 3" matches "10 Main St".
_UNIT_TOKENS = {"apt", "apartment", "unit", "ste", "suite", "fl", "floor",
                "rm", "room", "#"}


def _norm_street(s: str | None) -> str | None:
    """Normalize a street line to a comparable token string, or None."""
    if not s:
        return None
    s = str(s).lower()
    s = re.sub(r"[.,]", " ", s)
    s = s.replace("#", " # ")
    tokens: list[str] = []
    for tok in s.split():
        if tok in _UNIT_TOKENS:
            break  # drop apt/unit and anything after it
        tokens.append(_STREET_ABBR.get(tok, tok))
    out = " ".join(tokens).strip()
    return out or None


def _zip5(v) -> str | None:
    m = re.search(r"\d{5}", str(v or ""))
    return m.group(0) if m else None


def _addr_key(street: str | None, zip_code) -> str | None:
    """A street+zip5 key for address matching, or None if either part missing."""
    st = _norm_street(street)
    z = _zip5(zip_code)
    if st and z:
        return f"{st}|{z}"
    return None


_SUFFIXES = {"jr", "sr", "ii", "iii", "iv"}


def _name_candidates(s: str | None) -> list[tuple[str, str]]:
    """All plausible (last_lower, first_lower) keys for a member name.
    Handles 'LAST, FIRST', 'FIRST LAST', middle initials, multi-word last
    names, and trailing JR/SR/II/III/IV suffixes."""
    if not s:
        return []
    s = s.strip()
    out: list[tuple[str, str]] = []

    def _add(last: str, first: str) -> None:
        last, first = last.strip().lower(), first.strip().lower()
        if last and first and (last, first) not in out:
            out.append((last, first))

    if "," in s:
        last_raw, first_part = (p.strip() for p in s.split(",", 1))
        if not last_raw or not first_part:
            return []
        first_words = first_part.split()
        if not first_words:
            return []
        first = first_words[0]
        first_full = " ".join(first_words)
        last_clean = re.sub(r"\s+(jr|sr|ii|iii|iv)\.?$", "",
                            last_raw, flags=re.IGNORECASE).strip()
        _add(last_raw, first)
        if last_clean and last_clean.lower() != last_raw.lower():
            _add(last_clean, first)
        if first_full != first:
            _add(last_raw, first_full)
            if last_clean and last_clean.lower() != last_raw.lower():
                _add(last_clean, first_full)
    else:
        parts = s.split()
        if len(parts) < 2:
            return []
        if parts[-1].lower().rstrip(".") in _SUFFIXES and len(parts) >= 3:
            parts = parts[:-1]
        first = parts[0]
        for k in range(1, min(len(parts), 4)):
            _add(" ".join(parts[-k:]), first)
        if len(parts) >= 3:
            for k in range(1, min(len(parts) - 1, 4)):
                _add(" ".join(parts[-k:]), " ".join(parts[:-k]))
    return out


# ── Vendored: consumer index ───────────────────────────────────────
# Each entry's payload is (client_id, case_manager). Address rows additionally
# carry (last_lower, dob_iso) so an address hit can be corroborated.
_CONSUMER_INDEX_CACHE: dict | None = None


def load_consumer_index() -> dict:
    """Build every match index we can from Consumers_formatted.csv:
       - by_hpid     : { health_plan_id     : (client_id, cm) }
       - by_medicaid : { medicaid_norm      : (client_id, cm) }
       - by_namedob  : { (last,first,dob)   : (client_id, cm) }
       - by_addr     : { street|zip5        : [ (client_id, cm, last, dob) ] }

    HPID, name and DOB columns are required; Medicaid and address columns are
    optional (skipped if the CSV refresh doesn't include them). Cached for the
    process lifetime."""
    global _CONSUMER_INDEX_CACHE
    if _CONSUMER_INDEX_CACHE is not None:
        return _CONSUMER_INDEX_CACHE
    if not CONSUMERS_CSV.exists():
        raise FileNotFoundError(f"Consumers CSV not found at {CONSUMERS_CSV}")

    by_hpid: dict[str, tuple[str, str]] = {}
    by_medicaid: dict[str, tuple[str, str]] = {}
    by_namedob: dict[tuple[str, str, str], tuple[str, str]] = {}
    by_addr: dict[str, list[tuple[str, str, str, str]]] = {}
    client_by_consumer_uuid: dict[str, str] = {}  # CONSUMER_UUID -> CLIENT_ID

    required = (
        *CSV_HEALTH_PLAN_COLS, CSV_CLIENT_ID_COL, CSV_CASE_MANAGER_COL,
        CSV_FIRST_NAME_COL, CSV_LAST_NAME_COL, CSV_DOB_COL,
    )
    with CONSUMERS_CSV.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        cols = set(reader.fieldnames or [])
        missing = [c for c in required if c not in cols]
        if missing:
            raise RuntimeError(
                f"CSV missing expected columns: {missing}. "
                f"Available: {reader.fieldnames}"
            )
        medicaid_cols = [c for c in CSV_MEDICAID_COLS if c in cols]
        has_addr = CSV_ADDR_COL in cols and CSV_ZIP_COL in cols
        has_consumer_uuid = CSV_CONSUMER_UUID_COL in cols

        for row in reader:
            client_id = (row.get(CSV_CLIENT_ID_COL) or "").strip()
            if client_id.endswith(".0"):
                client_id = client_id[:-2]
            cm = (row.get(CSV_CASE_MANAGER_COL) or "").strip()
            if not client_id:
                continue
            pay = (client_id, cm)

            if has_consumer_uuid:
                cuuid = (row.get(CSV_CONSUMER_UUID_COL) or "").strip()
                if cuuid:
                    client_by_consumer_uuid[cuuid] = client_id

            for col in CSV_HEALTH_PLAN_COLS:
                v = (row.get(col) or "").strip()
                if v:
                    by_hpid[v] = pay

            for col in medicaid_cols:
                mid = _norm_id(row.get(col))
                if mid:
                    by_medicaid.setdefault(mid, pay)

            first = (row.get(CSV_FIRST_NAME_COL) or "").strip().lower()
            last = (row.get(CSV_LAST_NAME_COL) or "").strip().lower()
            dob = _norm_dob(row.get(CSV_DOB_COL))
            if first and last and dob:
                by_namedob[(last, first, dob)] = pay
                first_parts = first.split()
                if len(first_parts) >= 2 and first_parts[1]:
                    contracted = f"{first_parts[0]} {first_parts[1][0]}"
                    by_namedob.setdefault((last, contracted, dob), pay)
                    by_namedob.setdefault((last, first_parts[0], dob), pay)

            if has_addr:
                akey = _addr_key(row.get(CSV_ADDR_COL), row.get(CSV_ZIP_COL))
                if akey:
                    by_addr.setdefault(akey, []).append(
                        (client_id, cm, last, dob or "")
                    )

    log.info("Loaded consumer index: %d HPID, %d Medicaid, %d (name,DOB), "
             "%d address key(s), %d CONSUMER_UUID",
             len(by_hpid), len(by_medicaid), len(by_namedob), len(by_addr),
             len(client_by_consumer_uuid))
    _CONSUMER_INDEX_CACHE = {
        "by_hpid": by_hpid,
        "by_medicaid": by_medicaid,
        "by_namedob": by_namedob,
        "by_addr": by_addr,
        "client_by_consumer_uuid": client_by_consumer_uuid,
    }
    return _CONSUMER_INDEX_CACHE


# ── Active service-suspension index ────────────────────────────────
_SUSPENSION_INDEX_CACHE: dict | None = None


def _is_active_suspension(start_iso: str | None, end_iso: str | None,
                          today_iso: str) -> bool:
    """A suspension is 'active' when its window covers today: it has started
    (START_DATE <= today, or no start date recorded) AND has not yet ended
    (END_DATE blank, or END_DATE >= today). Tweak this one function to change
    what counts as active (e.g. require END_DATE blank only)."""
    if start_iso and start_iso > today_iso:
        return False  # not started yet
    if end_iso and end_iso < today_iso:
        return False  # already ended
    return True


def load_suspension_index() -> dict[str, tuple[str, str | None]]:
    """Map CLIENT_ID -> (service_names, suspension_start_iso) for consumers with
    an active service suspension. Built from Service_Suspensions_formatted.csv,
    bridged to Client ID via the consumer index's CONSUMER_UUID map.

    When a consumer has several active suspensions, the most recently started one
    wins for the date, and all distinct suspended-service names for that consumer
    are joined with '; '. Cached for the process lifetime. Returns {} (with a
    warning) if the suspensions CSV isn't synced — the rest of the run proceeds
    normally, just without suspension flags."""
    global _SUSPENSION_INDEX_CACHE
    if _SUSPENSION_INDEX_CACHE is not None:
        return _SUSPENSION_INDEX_CACHE

    if not SERVICE_SUSPENSIONS_CSV.exists():
        log.warning("[suspend] suspensions CSV not found at %s — suspension "
                    "flags will be skipped this run.", SERVICE_SUSPENSIONS_CSV)
        _SUSPENSION_INDEX_CACHE = {}
        return _SUSPENSION_INDEX_CACHE

    uuid_to_client = load_consumer_index().get("client_by_consumer_uuid") or {}
    if not uuid_to_client:
        log.warning("[suspend] Consumers export has no CONSUMER_UUID column — "
                    "cannot bridge suspensions to Client ID; skipping.")
        _SUSPENSION_INDEX_CACHE = {}
        return _SUSPENSION_INDEX_CACHE

    today_iso = datetime.now().strftime("%Y-%m-%d")
    # client_id -> {services: set, latest_start: iso}
    acc: dict[str, dict] = {}
    active = total = unbridged = 0
    with SERVICE_SUSPENSIONS_CSV.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        cols = set(reader.fieldnames or [])
        missing = [c for c in (SUSP_CONSUMER_UUID_COL, SUSP_SERVICE_COL,
                               SUSP_START_COL, SUSP_END_COL) if c not in cols]
        if missing:
            log.warning("[suspend] suspensions CSV missing column(s) %s — "
                        "skipping. Available: %s", missing, reader.fieldnames)
            _SUSPENSION_INDEX_CACHE = {}
            return _SUSPENSION_INDEX_CACHE

        for row in reader:
            total += 1
            start_iso = _norm_dob(row.get(SUSP_START_COL))
            end_iso = _norm_dob(row.get(SUSP_END_COL))
            if not _is_active_suspension(start_iso, end_iso, today_iso):
                continue
            active += 1
            cuuid = (row.get(SUSP_CONSUMER_UUID_COL) or "").strip()
            client_id = uuid_to_client.get(cuuid)
            if not client_id:
                unbridged += 1
                continue
            service = (row.get(SUSP_SERVICE_COL) or "").strip()
            entry = acc.setdefault(client_id, {"services": set(), "start": None})
            if service:
                entry["services"].add(service)
            if start_iso and (entry["start"] is None or start_iso > entry["start"]):
                entry["start"] = start_iso

    idx = {
        cid: ("; ".join(sorted(v["services"])), v["start"])
        for cid, v in acc.items()
    }
    log.info("[suspend] %d active suspension row(s) of %d total -> %d consumer(s) "
             "with an active suspension (%d row(s) had no Client ID bridge)",
             active, total, len(idx), unbridged)
    _SUSPENSION_INDEX_CACHE = idx
    return _SUSPENSION_INDEX_CACHE


# ── Active service-plan index + auth-vs-plan discrepancy check ─────
# Compares what UHC authorized against what the consumer's CURRENT ACTIVE
# service plan says. Grounded on the Misc/Service Plans export:
#   • UNITS_ALLOCATED equals the sum of the MONDAY..SUNDAY cells on every active
#     UHC row that has day cells, so units is the weekly total and the day cells
#     carry the weekday/weekend split.
#   • Weekend and night hours are SEPARATE allocation rows ("Personal Care -
#     Weekends", "HDM Meal Weekend or Holiday ..."), so everything is aggregated
#     per service FAMILY across rows before comparing.
_PLAN_INDEX_CACHE: dict | None = None

# WellSky SERVICE / UHC service name  ->  canonical family. Ordered: the first
# matching rule wins, so put the more specific tokens first.
_FAMILY_RULES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("hdm",            ("hdm", "home delivered meal", "meal")),
    ("pers",           ("pers", "personal emergency response",
                        "medication dispensing")),
    ("adh",            ("adult day health", "adult day")),
    ("personal_care",  ("personal care", "pca")),
    ("homemaker",      ("homemaker", "hmk")),
    ("companion",      ("companion",)),
    ("laundry",        ("laundry",)),
    ("chore",          ("chore",)),
    ("transportation", ("transportation", "chair car")),
    ("home_health_aide", ("home health aide", "supportive home care aide")),
    ("grocery",        ("grocery",)),
    ("cds",            ("consumer directed",)),
    ("gafc",           ("gafc", "group adult foster care", "adult foster care")),
    ("eaa",            ("environmental accessibility",)),
)

# Families whose weekly quantity is comparable between auth and plan, and the
# unit the auth expresses it in. Families not listed here are presence-only
# (monthly/per-trip services where a weekly number is meaningless).
_FAMILY_UNITS: dict[str, str] = {
    "personal_care": "hrs",
    "homemaker": "hrs",
    "companion": "hrs",
    "home_health_aide": "hrs",
    "laundry": "hrs",
    "cds": "hrs",
    "hdm": "meals",
    "adh": "days",
}

# UNITS_ALLOCATED (and the MONDAY..SUNDAY cells, which sum to it) are counted in
# 15-MINUTE units for time-based services — so 4 units = 1 hour. Meal services
# count actual meals and per-diem services count days, both 1:1. Multiply the
# plan side by this before comparing it to the authorization, which is always
# expressed in whole hours / meals / days.
_UNIT_TO_AUTH_FACTOR: dict[str, float] = {
    "hrs": 0.25,
    "meals": 1.0,
    "days": 1.0,
}

# Families that can stand in for another on the PLAN side. Personal Care is
# frequently delivered as Consumer Directed Services — the PDF says "Personal
# Care" but the plan carries it under CDS — so a CDS allocation satisfies a PC
# authorization rather than reading as "authorized but not on the plan".
_FAMILY_EQUIVALENTS: dict[str, tuple[str, ...]] = {
    "personal_care": ("cds",),
    "cds": ("personal_care",),
}

# Auth service lines that are one-off install / set-up FEES, not ongoing
# services. Whether these ever reach the service plan depends on the vendor —
# a MedScope set-up fee won't appear, a Lifeline one usually will — so they are
# never compared.
_SETUP_FEE_TOKENS = ("install", "set up", "setup", "set-up", "installation")

# change_type values that mean "this service is being STARTED". The plan is
# expected to be empty (or missing the service) until the initiation auth is
# keyed in, so nothing about that gap is a discrepancy. Case-manager feedback,
# 2026-08-05: "No services in place and initiation auth received; this is not
# considered a discrepancy." Both the old ("New") and current ("Initiate")
# wording are listed so historical rows behave the same.
_INITIATION_CHANGE_TYPES = {"new", "initiate", "initiation"}


def _fnum(v) -> float | None:
    s = str(v or "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _service_family(name: str) -> str | None:
    """Canonical family for a service name from EITHER vocabulary (WellSky's
    'HDM Meal Lunch Weekday Hot' or UHC's 'Home Delivered Meals').

    Tokens match on WORD BOUNDARIES, not as substrings — 'pers' is a substring
    of 'personal care', so plain `in` would file every Personal Care allocation
    under the PERS device family."""
    low = (name or "").strip().lower()
    if not low:
        return None

    # Consumer Directed Services carries the consumer's actual Personal Care
    # hours, but the CDS group also contains admin extras — Holiday, Earned
    # Time, Skills Training, Intake and Orientation — which are tiny one-off
    # allocations, not part of the PC schedule. Split those off so they can be
    # ignored instead of polluting the hours comparison.
    if "consumer directed" in low or "consumer directed svcs" in low:
        if any(t in low for t in ("holiday", "earned time", "skills training",
                                  "intake and orientation")):
            return "cds_ancillary"
        return "cds"
    for fam, tokens in _FAMILY_RULES:
        for t in tokens:
            # Trailing 's?' so a token matches its plural too ('meal' must hit
            # UHC's 'Home Delivered Meals', not just WellSky's 'HDM Meal ...').
            if re.search(rf"\b{re.escape(t)}s?\b", low):
                return fam
    return None


def _is_weekend_service(name: str) -> bool:
    """WellSky splits weekend coverage into its own allocation row."""
    low = (name or "").lower()
    return "weekend" in low or "holiday" in low


def _parse_qty(text: str | None, weekly_only: bool = False) -> float | None:
    """Leading number out of '12.75 hrs/week', '5 days/week', '7 meals/week',
    '283 units (per diem)'. Returns None when there's no number.

    With weekly_only=True, returns a number ONLY when the text expresses a
    per-WEEK rate. UHC also puts whole-auth totals in the same field ('696
    units', '283 units (per diem)'), and comparing those to a weekly plan figure
    produces nonsense like 'auth says 696 hrs/week'."""
    s = (text or "").strip()
    if not s:
        return None
    if weekly_only:
        low = s.lower()
        if not re.search(r"(/\s*(wk|week)|per\s+week|weekly)", low):
            return None
    num = ""
    for ch in s:
        if ch.isdigit() or (ch == "." and "." not in num):
            num += ch
        elif num:
            break
        elif ch in " \t":
            continue
        else:
            break
    return _fnum(num)


def _is_plan_row_active(row: dict, today_iso: str) -> bool:
    """Active = the care plan is Active, the allocation is Active, and the
    allocation window covers today."""
    if (row.get(PLAN_STATUS_COL) or "").strip() != "Active":
        return False
    if (row.get(PLAN_ALLOC_STATUS_COL) or "").strip() != "Active":
        return False
    start = _norm_dob(row.get(PLAN_ALLOC_START_COL))
    end = _norm_dob(row.get(PLAN_ALLOC_END_COL))
    if start and start > today_iso:
        return False
    if end and end < today_iso:
        return False
    return True


def load_active_schedule_ids() -> set[str] | None:
    """Schedule UUIDs whose OWN window covers today, from the Service Allocation
    export. Returns None (meaning 'no filter available') if the export isn't
    synced or lacks the columns — the caller then falls back to the allocation
    window, which over-counts superseded schedules but still runs."""
    if not SERVICE_ALLOCATION_CSV.exists():
        log.warning("[plan] service allocation CSV not found at %s — cannot "
                    "filter to the schedule in effect today; weekly totals will "
                    "include superseded schedules.", SERVICE_ALLOCATION_CSV)
        return None

    today_iso = datetime.now().strftime("%Y-%m-%d")
    active: set[str] = set()
    total = 0
    with SERVICE_ALLOCATION_CSV.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        cols = set(reader.fieldnames or [])
        missing = [c for c in (ALLOC_SCHED_UUID_COL, ALLOC_SCHED_START_COL,
                               ALLOC_SCHED_END_COL) if c not in cols]
        if missing:
            log.warning("[plan] service allocation CSV missing column(s) %s — "
                        "cannot filter to the current schedule.", missing)
            return None
        for row in reader:
            total += 1
            start = _norm_dob(row.get(ALLOC_SCHED_START_COL))
            end = _norm_dob(row.get(ALLOC_SCHED_END_COL))
            if start and start > today_iso:
                continue
            if end and end < today_iso:
                continue
            uuid = (row.get(ALLOC_SCHED_UUID_COL) or "").strip()
            if uuid:
                active.add(uuid)

    log.info("[plan] %d of %d schedule(s) are in effect today", len(active), total)
    return active


def load_service_plan_index(
        only_clients: set[str] | None = None) -> dict[str, dict[str, dict]]:
    """Map CLIENT_ID -> { family: {units, weekday, weekend, days, services,
    providers, uhc_funded, end_dates, auth_nos} } for allocations that are
    active TODAY, aggregated per service family.

    Pass `only_clients` (the Client IDs that actually have authorizations) to
    skip every other consumer while streaming — the export covers the whole
    agency, and only a small slice of it is ever compared. The file is a flat
    CSV with no index, so it still has to be read once either way.

    Rows from every care program are kept (a consumer's Personal Care may sit
    under Home Care rather than the SCO program, and calling that 'missing'
    would be a false positive). Cached for the process lifetime. Returns {} with
    a warning if the export isn't synced — the rest of the run proceeds without
    the check."""
    global _PLAN_INDEX_CACHE
    if _PLAN_INDEX_CACHE is not None:
        return _PLAN_INDEX_CACHE

    if not SERVICE_PLANS_CSV.exists():
        log.warning("[plan] service plans CSV not found at %s — the auth-vs-plan "
                    "check will be skipped this run.", SERVICE_PLANS_CSV)
        _PLAN_INDEX_CACHE = {}
        return _PLAN_INDEX_CACHE

    uuid_to_client = load_consumer_index().get("client_by_consumer_uuid") or {}
    if not uuid_to_client:
        log.warning("[plan] Consumers export has no CONSUMER_UUID column — "
                    "cannot bridge service plans to Client ID; skipping.")
        _PLAN_INDEX_CACHE = {}
        return _PLAN_INDEX_CACHE

    today_iso = datetime.now().strftime("%Y-%m-%d")
    # Schedules in effect today. A service allocation stacks several schedules
    # over its life and only one is current — without this every past schedule
    # gets summed in.
    active_scheds = load_active_schedule_ids()
    superseded = 0
    idx: dict[str, dict[str, dict]] = {}
    # Diagnostic: a consumer carrying more than one ACTIVE care plan would have
    # their allocations summed across both, inflating the weekly totals.
    care_plans_by_client: dict[str, set[str]] = {}
    total = active = unbridged = unmapped = duplicates = 0

    with SERVICE_PLANS_CSV.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        cols = set(reader.fieldnames or [])
        required = (PLAN_CONSUMER_UUID_COL, PLAN_SERVICE_COL, PLAN_STATUS_COL,
                    PLAN_ALLOC_STATUS_COL, PLAN_UNITS_COL)
        missing = [c for c in required if c not in cols]
        if missing:
            log.warning("[plan] service plans CSV missing column(s) %s — "
                        "skipping the auth-vs-plan check.", missing)
            _PLAN_INDEX_CACHE = {}
            return _PLAN_INDEX_CACHE
        day_cols = [c for c in PLAN_DAY_COLS if c in cols]

        for row in reader:
            total += 1
            if not _is_plan_row_active(row, today_iso):
                continue
            active += 1
            client_id = uuid_to_client.get(
                (row.get(PLAN_CONSUMER_UUID_COL) or "").strip()
            )
            if not client_id:
                unbridged += 1
                continue
            if only_clients is not None and client_id not in only_clients:
                continue
            cp_uuid = (row.get(PLAN_CARE_PLAN_UUID_COL) or "").strip()
            if cp_uuid:
                care_plans_by_client.setdefault(client_id, set()).add(cp_uuid)
            # Drop schedules that have been superseded — only the one whose own
            # window covers today describes the service the consumer gets now.
            sched_uuid = (row.get(PLAN_SCHED_UUID_COL) or "").strip()
            if active_scheds is not None and sched_uuid and sched_uuid not in active_scheds:
                superseded += 1
                continue

            service = (row.get(PLAN_SERVICE_COL) or "").strip()
            fam = _service_family(service)
            if not fam:
                unmapped += 1
                continue

            bucket = idx.setdefault(client_id, {}).setdefault(fam, {
                "units": 0.0, "weekday": 0.0, "weekend": 0.0,
                "days": set(), "services": set(), "providers": set(),
                "uhc_funded": False, "end_dates": set(), "auth_nos": set(),
                "seen": set(),
            })
            # Count each allocation schedule once — the same one repeats across
            # rows (one per subservice), each carrying the same UNITS_ALLOCATED.
            dedup_key = ((row.get(PLAN_SCHED_UUID_COL) or "").strip()
                         or (row.get(PLAN_ALLOC_UUID_COL) or "").strip())
            if dedup_key:
                if dedup_key in bucket["seen"]:
                    duplicates += 1
                    continue
                bucket["seen"].add(dedup_key)
            units = _fnum(row.get(PLAN_UNITS_COL)) or 0.0
            bucket["units"] += units
            bucket["services"].add(service)
            provider = (row.get(PLAN_PROVIDER_COL) or "").strip()
            if provider:
                bucket["providers"].add(provider)
            if PLAN_UHC_PROGRAM_TOKEN in (row.get(PLAN_PROGRAM_COL) or ""):
                bucket["uhc_funded"] = True
            end = _norm_dob(row.get(PLAN_ALLOC_END_COL))
            if end:
                bucket["end_dates"].add(end)
            auth_no = (row.get(PLAN_PRIOR_AUTH_COL) or "").strip()
            if auth_no:
                bucket["auth_nos"].add(auth_no)

            # Day cells carry the split. A "- Weekends"/"Holiday" row with no day
            # cells still counts entirely as weekend time.
            any_day = False
            for c in day_cols:
                v = _fnum(row.get(c))
                if v is None:
                    continue
                any_day = True
                bucket["days"].add(c)
                if c in PLAN_WEEKEND_COLS:
                    bucket["weekend"] += v
                else:
                    bucket["weekday"] += v
            if not any_day and units:
                if _is_weekend_service(service):
                    bucket["weekend"] += units
                else:
                    bucket["weekday"] += units

    log.info("[plan] %d active allocation row(s) of %d total -> %d consumer(s) "
             "(%d row(s) had no Client ID bridge, %d service name(s) unmapped, "
             "%d duplicate schedule row(s) collapsed, %d superseded schedule "
             "row(s) dropped)",
             active, total, len(idx), unbridged, unmapped, duplicates, superseded)
    multi = {c: len(u) for c, u in care_plans_by_client.items() if len(u) > 1}
    if multi:
        log.warning("[plan] %d of %d consumer(s) have MORE THAN ONE active care "
                    "plan (max %d) — their weekly totals are summed across all "
                    "of them, which inflates the comparison.",
                    len(multi), len(care_plans_by_client), max(multi.values()))
    _PLAN_INDEX_CACHE = idx
    return _PLAN_INDEX_CACHE


def _fmt_qty(v: float) -> str:
    return f"{v:g}"


def compare_auth_to_plan(parsed: dict,
                         plan: dict[str, dict] | None) -> tuple[str, str]:
    """Compare one parsed authorization against that consumer's active plan.

    Returns (summary, detail):
      summary — short flag for the list view, e.g. "3 discrepancies" / "OK" /
                "No active plan".
      detail  — one line per finding, or "" when everything agrees.

    Only APPROVED service lines are compared; denied/terminated lines are
    ignored entirely.

    Program feedback narrowed this to TWO findings, and nothing else counts as
    a discrepancy:
      1. weekly HOURS (or meals) mismatch per service family,
      2. DAY-SCHEDULE mismatch (days/week) for per-diem services.
    Explicitly dropped: a consumer with NO active service plan at all (program
    feedback 2026-08-11), a service authorized but absent from an otherwise
    active plan, a denied/terminated line still active on the plan, and
    plan-allocation end dates that fall short of the auth period."""
    services = parsed.get("services")
    if not isinstance(services, list):
        services = []
    approved = [s for s in services
                if isinstance(s, dict) and s.get("approved") is not False]

    # What the auth is DOING drives how much of it is comparable at all:
    #   • increase — the auth is what the plan will BECOME. A plan sitting
    #                BELOW the authorized amount is the expected mid-flight
    #                state, not a finding (CM feedback 2026-08-05: consumer
    #                on HM 2.5 hrs/wk, auth received to increase to 8).
    #                A plan sitting ABOVE it is still a real discrepancy.
    change = (str(parsed.get("change_type") or "")).strip().lower()
    is_increase = change == "increase"

    if not plan:
        # No active service plan is NOT a discrepancy (program feedback
        # 2026-08-11). With nothing to compare against there are no hours or
        # day-schedule differences to report, so the row is clean.
        return PLAN_SUMMARY_OK if approved else "", ""

    findings: list[str] = []

    # NOTE: deliberately NOT flagging a denied/terminated auth line whose service
    # is still active on the plan. Program feedback 2026-08-06: the only findings
    # that matter are the weekly hours, the day schedule, and having no active
    # service plan at all. Nothing else is a discrepancy.

    # Aggregate the auth side per family too — UHC also splits a service across
    # lines (HDM weekday/weekend meals, PC vs PC medical appointment).
    auth_fams: dict[str, dict] = {}
    for s in approved:
        name = (s.get("name") or "").strip()
        combined = f"{name} {(s.get('subcategory') or '')}"
        # One-off install / set-up fee lines are never on the plan reliably.
        if any(t in combined.lower() for t in _SETUP_FEE_TOKENS):
            continue
        fam = _service_family(combined)
        if not fam or fam == "cds_ancillary":
            continue
        a = auth_fams.setdefault(fam, {
            "names": set(), "total": 0.0, "weekday": 0.0, "weekend": 0.0,
            "has_total": False, "has_split": False,
            "lines": 0, "lines_with_total": 0, "max_line_days": 0.0,
        })
        a["lines"] += 1
        a["names"].add(name or fam)
        total = _parse_qty(s.get("units_or_frequency"), weekly_only=True)
        wd = _parse_qty(s.get("weekday_hours"), weekly_only=True)
        we = _parse_qty(s.get("weekend_hours"), weekly_only=True)
        if total is not None:
            a["total"] += total
            a["has_total"] = True
            a["lines_with_total"] += 1
        if wd is not None:
            a["weekday"] += wd
            a["has_split"] = True
        if we is not None:
            a["weekend"] += we
            a["has_split"] = True
        # Per-diem families (ADH) are the one place summing across lines is
        # WRONG. UHC bills the same attendance three ways — per diem, per 15
        # min, per half day — and the extraction routinely emits one line each,
        # every one of them carrying the SAME "5 days/week" from the
        # notification notes. Summing produced "auth says 20 day(s)/week".
        # The lines describe one schedule, so take the largest, not the sum.
        a["max_line_days"] = max(a["max_line_days"], (wd or 0.0) + (we or 0.0))

    for fam, a in sorted(auth_fams.items()):
        label = " / ".join(sorted(a["names"])) or fam
        p = plan.get(fam)
        if not p:
            # Fall back to an equivalent family before calling it missing — PC
            # is routinely carried on the plan as Consumer Directed Services.
            for alt in _FAMILY_EQUIVALENTS.get(fam, ()):
                if plan.get(alt):
                    p = plan[alt]
                    break
        if not p:
            # The consumer HAS an active plan, this one service just isn't on it.
            # Program feedback 2026-08-06: not a discrepancy — only a consumer
            # with NO active plan at all is flagged (handled above). Nothing to
            # compare hours or days against, so move on.
            continue

        unit = _FAMILY_UNITS.get(fam)
        if unit:
            # UNITS_ALLOCATED is in 15-minute units for time-based services, so
            # scale the plan side into the auth's unit before comparing.
            factor = _UNIT_TO_AUTH_FACTOR.get(unit, 1.0)
            plan_total = p["units"] * factor

            # Per-diem families (ADH) express units_or_frequency as a whole-auth
            # unit count ("283 units (per diem)"), which is NOT a weekly figure —
            # comparing it to the plan would be meaningless. For those, only the
            # days/week check below applies.
            if unit == "days":
                auth_days = a["max_line_days"] if a["has_split"] else None
                plan_days = len(p["days"]) if p["days"] else 0
                if (auth_days and plan_days
                        and abs(auth_days - plan_days) >= 1
                        and not (is_increase and plan_days < auth_days)):
                    findings.append(
                        f"{label}: auth says {_fmt_qty(auth_days)} day(s)/week, "
                        f"plan is scheduled {plan_days} day(s)/week."
                    )
            else:
                # Weekly total. Only comparable when EVERY line in the family
                # yielded a weekly figure — otherwise the sum is partial (a PC
                # auth whose main line has no weekly total but whose medical-
                # appointment line does would "total" 0.75 hrs/week) and the
                # comparison is meaningless. Fall back to the weekday+weekend
                # split, which is the more reliable of the two.
                if a["has_total"] and a["lines_with_total"] == a["lines"]:
                    auth_total = a["total"]
                elif a["has_split"]:
                    auth_total = a["weekday"] + a["weekend"]
                else:
                    auth_total = None
                if (auth_total is not None
                        and abs(auth_total - plan_total) > PLAN_HOURS_TOLERANCE
                        and not (is_increase and plan_total < auth_total)):
                    findings.append(
                        f"{label}: auth says {_fmt_qty(auth_total)} {unit}/week, "
                        f"plan has {_fmt_qty(plan_total)} {unit}/week "
                        f"(off by {_fmt_qty(abs(auth_total - plan_total))})."
                    )
                # NOTE: the weekday/weekend split is deliberately NOT compared.
                # When the weekly total matches, which days the hours land on is
                # scheduling, not a discrepancy — an auth reading "3 hrs
                # weekday" against a plan that schedules the same 3 hrs on a
                # Saturday is correct, not a finding. The same applies to meals
                # (weekend meals are delivered on a weekday frozen drop) and to
                # auths whose own split contradicts their own total (PC totals
                # net out medical-escort hours while the weekday figure does
                # not). The weekly total is the contract; the day is not.

        # NOTE: the plan allocation's end date vs the auth's end date is
        # deliberately NOT compared. Program feedback 2026-08-06: end-date drift
        # is not a discrepancy — only hours and the day schedule are.

    # NOTE: deliberately NOT flagging services that are on the plan but absent
    # from this authorization. Each auth document covers only the service(s) it
    # renews — the consumer's other services come from their own authorizations,
    # so "PC auth doesn't mention Homemaker" is normal, not a discrepancy. A dry
    # run over the live list showed this one check producing 1,565 of 2,753
    # findings, all noise.

    if not findings:
        return PLAN_SUMMARY_OK, ""
    n = len(findings)
    summary = f"{n} discrepanc{'y' if n == 1 else 'ies'}"
    return summary, "\n".join(f"• {f}" for f in findings)


def normalize_change_types(g: requests.Session, site_id: str,
                           items: list[dict],
                           update_in_place: bool = True) -> None:
    """Rewrite any stale Change Type wording on the list to its current form.

    CHANGE_TYPE_DISPLAY is applied when a row is first enriched, so a rename
    lands on new rows only and every older row keeps whatever word was current
    the day it was written — the list ends up carrying three spellings of the
    same thing ('New', 'Initiation', 'Initiate'). Rather than a one-off backfill
    script somebody has to remember to run, this pass runs on EVERY pipeline run
    and repairs whatever it finds. Writes only rows whose value actually
    changes, so it costs nothing once the list is clean."""
    col = CHANGE_TYPE_FIELD
    stale = [it for it in items
             if ((it.get("fields") or {}).get(col) or "").strip().lower()
             in CHANGE_TYPE_DISPLAY]
    if not stale:
        log.info("[changetype] all rows already on current wording")
        return

    wrote = errors = 0
    for it in stale:
        f = it.get("fields") or {}
        cur = (f.get(col) or "").strip()
        want = CHANGE_TYPE_DISPLAY[cur.lower()]
        if want == cur:
            continue
        try:
            patch_fields(g, site_id, it["id"], {col: want})
            if update_in_place:
                f[col] = want
            wrote += 1
        except Exception as e:
            errors += 1
            log.warning("[changetype] PATCH failed for item %s: %s", it["id"], e)
    log.info("[changetype] normalized %d row(s) to current wording (errors=%d)",
             wrote, errors)


def sync_plan_discrepancies(g: requests.Session, site_id: str,
                            items: list[dict],
                            update_in_place: bool = True,
                            dry_run: bool = False) -> None:
    """Compare every matched auth row against the consumer's active service plan
    and write the result to the 'Plan Discrepancy' / 'Plan Discrepancy Detail'
    columns — clearing them when the row no longer has findings. Plans change
    daily and independently of the auths, so this runs over the whole list each
    run and writes only the rows whose value actually changed.

    Skips entirely (with a warning) if the columns aren't in the list yet or the
    service plans / Consumers data can't be loaded."""
    disc_col = resolve_column(g, site_id, PLAN_DISCREPANCY_COL_TITLE)
    detail_col = resolve_column(g, site_id, PLAN_DISCREPANCY_DETAIL_COL_TITLE)
    if not disc_col or not detail_col:
        log.warning("[plan] column(s) not found in list (%r->%s, %r->%s) — add "
                    "them in the list UI; skipping the auth-vs-plan check.",
                    PLAN_DISCREPANCY_COL_TITLE, disc_col,
                    PLAN_DISCREPANCY_DETAIL_COL_TITLE, detail_col)
        return
    log.info("[plan] columns resolved: %r->%s, %r->%s",
             PLAN_DISCREPANCY_COL_TITLE, disc_col,
             PLAN_DISCREPANCY_DETAIL_COL_TITLE, detail_col)
    # The Excel/calendar read paths use the PINNED internal names. SharePoint
    # derives internal names from whatever the column was originally called, so
    # if they don't match, the report columns would silently come out blank.
    for resolved, pinned, title in (
        (disc_col, PLAN_DISC_FIELD, PLAN_DISCREPANCY_COL_TITLE),
        (detail_col, PLAN_DISC_DETAIL_FIELD, PLAN_DISCREPANCY_DETAIL_COL_TITLE),
    ):
        if resolved != pinned:
            log.warning("[plan] column %r resolves to internal name %r but the "
                        "report path is pinned to %r — update that constant or "
                        "the Excel column will be blank.", title, resolved, pinned)

    if dry_run:
        log.info("[plan] DRY RUN — computing findings only, nothing is written.")

    # Only the consumers who actually have an authorization on this list — the
    # export covers the whole agency, so this drops most of it before it's held
    # in memory or compared.
    want_clients = {
        (it.get("fields", {}) or {}).get("ClientID", "").strip()
        for it in items if not _xls_is_duplicate(it)
    }
    want_clients.discard("")
    log.info("[plan] %d consumer(s) with authorizations to look up", len(want_clients))

    idx = load_service_plan_index(only_clients=want_clients)
    if not idx:
        log.warning("[plan] empty service-plan index — skipping (existing flags "
                    "left untouched rather than mass-cleared).")
        return

    # Two different questions, so two different tallies:
    #   totals  — the state of the WHOLE list after this run. Every compared row
    #             lands in exactly one bucket, whether or not it changed.
    #   written — how many rows this run actually PATCHed.
    # These used to be conflated: the counters sat after the `if not body`
    # guard, so a run reporting "flagged=53" meant "53 rows whose flag MOVED",
    # not "53 rows have a discrepancy" — with ~960 unchanged rows invisible.
    totals = {"flagged": 0, "ok": 0, "no_plan": 0, "n/a": 0}
    written = cleared = skipped = errors = 0

    def bucket(summary: str) -> str:
        """Classify a COMPUTED summary by exact value. 'No active plan' is no
        longer emitted (2026-08-11 — not a discrepancy), so that bucket now
        always reads 0; it stays only so the counter line keeps its shape."""
        if not summary:
            return "n/a"                    # nothing approved to compare
        if summary == PLAN_SUMMARY_NO_PLAN:
            return "no_plan"
        if summary == PLAN_SUMMARY_OK:
            return "ok"
        return "flagged"                    # "N discrepanc(y|ies)"

    for it in items:
        f = it.get("fields", {}) or {}
        if _xls_is_duplicate(it):
            skipped += 1
            continue
        client_id = (f.get("ClientID") or "").strip()
        if not client_id:
            skipped += 1  # unmatched rows can't be tied to a consumer
            continue
        parsed = parse_payload(f.get(JSON_PAYLOAD_FIELD))
        if not parsed:
            skipped += 1
            continue

        want_disc, want_detail = compare_auth_to_plan(parsed, idx.get(client_id))
        # Count BEFORE the change check — an unchanged row is still a row with
        # (or without) a discrepancy, and it belongs in the totals.
        totals[bucket(want_disc)] += 1

        cur_disc = _plain_text(f.get(disc_col))
        cur_detail = _plain_text(f.get(detail_col))

        body: dict = {}
        if want_disc != cur_disc:
            body[disc_col] = want_disc
        if want_detail != cur_detail:
            body[detail_col] = want_detail
        if not body:
            continue

        if dry_run:
            # Auth number only — never log member names or plan details.
            log.info("[plan][dry-run] auth %s (Client %s): %s%s",
                     (f.get("Title") or "?").strip(), client_id,
                     want_disc or "(cleared)",
                     "\n" + want_detail if want_detail else "")
            written += 1
            if not want_disc:
                cleared += 1
            continue

        try:
            patch_fields(g, site_id, it["id"], body)
            if update_in_place:
                for k, v in body.items():
                    f[k] = v
            written += 1
            if not want_disc:
                cleared += 1
        except Exception as e:
            log.warning("[plan] PATCH failed for item %s: %s", it["id"], e)
            errors += 1

    log.info("[plan] %sdone. LIST TOTALS: flagged=%d ok=%d no-plan=%d n/a=%d "
             "(compared=%d, skipped=%d) | THIS RUN: %s=%d (cleared=%d) errors=%d",
             "DRY RUN " if dry_run else "",
             totals["flagged"], totals["ok"], totals["no_plan"], totals["n/a"],
             sum(totals.values()), skipped,
             "would-write" if dry_run else "wrote", written, cleared, errors)


# ── SharePoint helpers (list-scoped) ───────────────────────────────
# Filtering/ordering on these columns needs them flagged "Indexed" in the list
# settings (Title and LookupStatus). Until ~5,000 rows, non-indexed equality
# filters still work with this Prefer header; past that they may start failing,
# so index those two columns in the list UI before the list gets large.
_NONINDEXED_PREFER = {"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"}


def list_items(g: requests.Session, site_id: str,
               filter_: str | None = None) -> list[dict]:
    """Page through list items. With `filter_` (an OData $filter on fields/…),
    only matching items are pulled — used to fetch just unprocessed rows so the
    per-run cost tracks new items, not total list size."""
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{SP_LIST_ID}/items?expand=fields&$top=100"
    if filter_:
        url += f"&$filter={quote(filter_)}"
    headers = _NONINDEXED_PREFER if filter_ else None
    out: list[dict] = []
    while url:
        r = g.get(url, headers=headers)
        r.raise_for_status()
        data = r.json()
        out.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return out


def list_last_modified(g: requests.Session, site_id: str) -> str | None:
    """The list's own lastModifiedDateTime — a single metadata read on the list
    resource. Used to decide whether anything changed since the previous run
    WITHOUT fetching or opening any item (no payload / PII is touched). Returns
    an ISO-8601 'Z' string, or None if the read fails."""
    r = g.get(
        f"{GRAPH_BASE}/sites/{site_id}/lists/{SP_LIST_ID}"
        f"?$select=lastModifiedDateTime"
    )
    if not r.ok:
        log.warning("[gate] could not read list lastModifiedDateTime: %s %s",
                    r.status_code, r.text[:150])
        return None
    return r.json().get("lastModifiedDateTime")


def title_exists(g: requests.Session, site_id: str, auth: str,
                 exclude_id: str | int | None = None) -> bool | None:
    """True if a list item OTHER THAN exclude_id has Title == auth, False if not,
    None if the check could not be performed (caller should not block on None).
    Used to catch a new item that duplicates one processed in a prior run.

    exclude_id matters whenever the row already carries its authorization number
    in Title at creation time, which is how the Tufts ingest writes rows. Without
    it the guard finds the very item it is checking and demotes it as a duplicate
    of itself: on 2026-08-17 that made all 8 Tufts rows duplicates, 0 canonical,
    and every hourly run produced an empty workbook and an empty calendar while
    still exiting 0."""
    if not auth:
        return False
    safe = auth.replace("'", "''")
    flt = quote(f"fields/Title eq '{safe}'")
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/lists/{SP_LIST_ID}/items"
        f"?$select=id&$filter={flt}&$top={2 if exclude_id is not None else 1}"
    )
    r = g.get(url, headers=_NONINDEXED_PREFER)
    if not r.ok:
        log.warning("[guard] Title-exists check failed for %r: %s %s",
                    auth, r.status_code, r.text[:150])
        return None
    ids = [str(v.get("id")) for v in r.json().get("value", [])]
    if exclude_id is not None:
        ids = [i for i in ids if i != str(exclude_id)]
    return bool(ids)


_EXISTING_COLUMNS: set[str] | None = None


def existing_columns(g: requests.Session, site_id: str) -> set[str]:
    """Internal names of the list's columns, cached. Used to skip writing to
    columns that haven't been added to the list yet (the app registration
    can't create columns, so newer payload fields like address/change_type are
    only written once someone adds the column in the list UI)."""
    global _EXISTING_COLUMNS
    if _EXISTING_COLUMNS is not None:
        return _EXISTING_COLUMNS
    cols: set[str] = set()
    r = g.get(f"{GRAPH_BASE}/sites/{site_id}/lists/{SP_LIST_ID}/columns?$select=name")
    if r.ok:
        cols = {c.get("name") for c in r.json().get("value", []) if c.get("name")}
    else:
        log.warning("[schema] could not read columns: %s %s",
                    r.status_code, r.text[:150])
    _EXISTING_COLUMNS = cols
    return cols


_COLUMN_NAME_BY_LABEL: dict[str, str] | None = None


def resolve_column(g: requests.Session, site_id: str, label: str) -> str | None:
    """Resolve a column's INTERNAL name from either its display name or its
    internal name (case-insensitive). Lets the pipeline target newly-added
    columns by their friendly title (e.g. 'Service Suspended') without having to
    know the _x0020_-encoded internal name the UI generated. Returns None if no
    column matches (caller logs + skips)."""
    global _COLUMN_NAME_BY_LABEL
    if _COLUMN_NAME_BY_LABEL is None:
        _COLUMN_NAME_BY_LABEL = {}
        r = g.get(f"{GRAPH_BASE}/sites/{site_id}/lists/{SP_LIST_ID}"
                  f"/columns?$select=name,displayName")
        if r.ok:
            for c in r.json().get("value", []):
                internal = c.get("name")
                if not internal:
                    continue
                _COLUMN_NAME_BY_LABEL[internal.lower()] = internal
                disp = c.get("displayName")
                if disp:
                    _COLUMN_NAME_BY_LABEL.setdefault(disp.lower(), internal)
        else:
            log.warning("[schema] could not read columns for label resolution: "
                        "%s %s", r.status_code, r.text[:150])
    return _COLUMN_NAME_BY_LABEL.get(label.lower())


def patch_fields(g: requests.Session, site_id: str, item_id: str,
                 fields: dict) -> None:
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/lists/{SP_LIST_ID}/items/{item_id}/fields"
    )
    r = g.patch(url, json=fields)
    if not r.ok:
        raise RuntimeError(
            f"PATCH fields failed for item {item_id}: {r.status_code} {r.text}"
        )


def delete_item(g: requests.Session, site_id: str, item_id: str) -> None:
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{SP_LIST_ID}/items/{item_id}"
    r = g.delete(url)
    if not r.ok and r.status_code != 404:
        raise RuntimeError(
            f"DELETE failed for item {item_id}: {r.status_code} {r.text}"
        )


_DUP_CHOICE_ENSURED = False


def _duplicate_choice_present(g: requests.Session, site_id: str) -> bool:
    """Read the Lookup Status column via Graph (reads work fine) and report
    whether 'Duplicate' is already an allowed choice."""
    r = g.get(
        f"{GRAPH_BASE}/sites/{site_id}/lists/{SP_LIST_ID}/columns"
        f"?$filter=name eq '{LOOKUP_STATUS_FIELD}'"
    )
    if not r.ok:
        return False
    vals = r.json().get("value", [])
    if not vals:
        return False
    return DUPLICATE_STATUS in ((vals[0].get("choice") or {}).get("choices") or [])


def ensure_duplicate_choice(g: requests.Session, sp: requests.Session,
                            site_id: str) -> bool:
    """Return True if 'Duplicate' is an allowed Lookup Status choice — adding it
    first if possible. Reading/writing choice VALUES works via Graph; only
    editing the choice SCHEMA is blocked, so we try SharePoint REST for that.
    If both the existing-check and the add fail, the caller falls back to a
    Notes marker."""
    global _DUP_CHOICE_ENSURED
    if _DUP_CHOICE_ENSURED:
        return True
    if _duplicate_choice_present(g, site_id):
        _DUP_CHOICE_ENSURED = True
        return True
    fields_url = (
        f"{SP_SITE_URL}/_api/web/lists(guid'{SP_LIST_ID}')"
        f"/fields?$filter=InternalName eq '{LOOKUP_STATUS_FIELD}'"
    )
    r = sp.get(fields_url)
    if not r.ok:
        log.info("[dup] cannot auto-add 'Duplicate' choice (SP REST: %s) - "
                 "flagging duplicates in Notes instead", r.status_code)
        return False
    results = (r.json().get("d") or {}).get("results") or []
    if not results:
        log.warning("[dup] %s field not found via SP REST", LOOKUP_STATUS_FIELD)
        return False
    field = results[0]
    choices = list((field.get("Choices") or {}).get("results") or [])
    if DUPLICATE_STATUS in choices:
        _DUP_CHOICE_ENSURED = True
        return True
    choices.append(DUPLICATE_STATUS)
    field_id = field["Id"]
    update_url = (
        f"{SP_SITE_URL}/_api/web/lists(guid'{SP_LIST_ID}')/fields(guid'{field_id}')"
    )
    body = {
        "__metadata": {"type": "SP.FieldChoice"},
        "Choices": {"results": choices},
    }
    pr = sp.post(update_url, json=body,
                 headers={"X-HTTP-Method": "MERGE", "IF-MATCH": "*"})
    if not pr.ok:
        log.warning("[dup] SP REST choice update failed: %s %s",
                    pr.status_code, pr.text[:200])
        return False
    log.info("[dup] added '%s' choice to %s column",
             DUPLICATE_STATUS, LOOKUP_STATUS_FIELD)
    _DUP_CHOICE_ENSURED = True
    return True


# ── Payload parsing ────────────────────────────────────────────────
def parse_payload(raw: str | None) -> dict | None:
    """Extract the inner authorization JSON from the raw 'JSON Payload' value.

    The column holds the model-response object; the auth data is the JSON
    string in its `text` field (sometimes wrapped in a ```json code fence).
    Falls back to treating the whole payload as the auth object if there's no
    `text` field.
    """
    if not raw:
        return None
    try:
        outer = json.loads(raw) if isinstance(raw, str) else raw
    except (json.JSONDecodeError, TypeError):
        return None

    text = outer.get("text") if isinstance(outer, dict) else None
    if not text:
        # Maybe the payload already IS the auth object.
        # Tufts eFax forms carry a reference_number where UHC carries an
        # authorization_number. Requiring only the UHC spelling threw away 15
        # good extractions on the Tufts list -- they parsed fine, were never
        # matched to a consumer, and landed in Needs Review.
        return outer if isinstance(outer, dict) and (
            "authorization_number" in outer
            or "reference_number" in outer) else None

    t = text.strip()
    if t.startswith("```"):
        t = t.strip("`")
        if t[:4].lower() == "json":
            t = t[4:]
        t = t.strip()
    try:
        inner = json.loads(t)
    except json.JSONDecodeError:
        log.warning("[parse] inner text is not valid JSON: %r", t[:200])
        return None
    return inner if isinstance(inner, dict) else None


def _iso_date(v) -> str | None:
    # Noon UTC, NOT midnight. These are Date-Only columns, but SharePoint
    # stores the instant in UTC and the list UI renders it in the site's
    # regional time zone (Eastern, UTC-4/-5). Midnight UTC therefore displays
    # as the PREVIOUS calendar day: a DOB of 1971-01-01 shows up in the list as
    # 12/31/1970. Noon UTC lands on the same calendar day in every zone from
    # UTC-11 to UTC+11, so what the fax says is what the list shows.
    norm = _norm_dob(v)  # -> 'YYYY-MM-DD' or None
    return f"{norm}T12:00:00Z" if norm else None


def render_services(services) -> str:
    """Human-readable multiline summary of the services array."""
    if not isinstance(services, list) or not services:
        return ""
    lines: list[str] = []
    for s in services:
        if not isinstance(s, dict):
            continue
        name = (s.get("name") or "").strip() or "Service"
        sub = (s.get("subcategory") or "").strip()
        if sub:
            name = f"{name} ({sub})"
        approved = s.get("approved")
        status = ("Approved" if approved is True
                  else "Denied" if approved is False else "—")
        freq = (s.get("units_or_frequency") or "").strip()
        start = _norm_dob(s.get("start_date"))
        end = _norm_dob(s.get("end_date"))
        bits = [f"{name}: {status}"]
        if freq:
            bits.append(freq)
        if start or end:
            bits.append(f"{start or '?'} → {end or '?'}")
        line = " — ".join(bits)
        denial = (s.get("denial_reason") or "").strip()
        if denial:
            line += f"\n    Reason: {denial}"
        lines.append(line)
    return "\n".join(lines)


def auth_number_of(item: dict, parsed: dict | None) -> str:
    """Authorization number for dedup/Title — existing Title wins, else parsed."""
    f = item.get("fields") or {}
    existing = (f.get("Title") or "").strip()
    if existing:
        return existing
    if parsed:
        return (str(parsed.get("authorization_number") or "")).strip()
    return ""


# ── Build the PATCH body from a parsed payload ─────────────────────
# ── Journal-note QA (deterministic double-check) ───────────────────────
# Mirrors the bridge's _SVC_ABBR (service full name -> subject abbreviation).
# The note must LEAD each service with its abbreviation, so the full name
# appearing without its abbrev — or an "of <full name>" / "<full name> service"
# construction — is the #1 authoring error. Only services whose abbrev actually
# DIFFERS from the full name are listed; PERS / Laundry / Transportation abbrev
# to themselves so no word-order error is possible for them.
_NOTE_ABBR_FULLNAMES: tuple[tuple[str, str], ...] = (
    ("home delivered meals", "HDM"),
    ("adult day health", "ADH"),
    ("consumer directed services", "CDC"),
    ("consumer directed care", "CDC"),
    ("personal care", "PC"),
    ("homemaker", "HM"),
    ("chore", "HCH"),
)
# Canonical service family -> subject abbreviation, for checking the full name
# of a parsed["services"] item against the note. Keys match _service_family().
_NOTE_FAMILY_ABBR: dict[str, tuple[str, str]] = {
    # family: (full name as it reads in a note, abbreviation)
    "hdm":           ("Home Delivered Meals", "HDM"),
    "adh":           ("Adult Day Health", "ADH"),
    "personal_care": ("Personal Care", "PC"),
    "homemaker":     ("Homemaker", "HM"),
    "chore":         ("Chore", "HCH"),
    "cds":           ("Consumer Directed Services", "CDC"),
}
# Families that legitimately count in whole UNITS in the note: ADH is per-diem
# (day counts) and transportation is per-trip. Every other compared family is
# hrs/wk (or meals/wk for HDM), so "unit(s)" in the note is an error there.
_NOTE_UNIT_OK_FAMILIES = frozenset({"adh", "transportation"})


def validate_note(parsed: dict) -> list[str]:
    """Deterministic double-check of a journal note against the structured data.

    Returns a list of human-readable issue strings; an empty list means the note
    passed every check. This is a FLAG generator only — it never edits the note.
    """
    if not isinstance(parsed, dict):
        return []
    note = (parsed.get("journal_note") or "").strip()
    if not note:
        return []
    low = note.lower()
    services = parsed.get("services")
    if not isinstance(services, list):
        services = []
    issues: list[str] = []

    def _add(msg: str) -> None:
        if msg not in issues:
            issues.append(msg)

    # (a) Service abbreviation / word order — the note must lead with the abbrev.
    for full, abbr in _NOTE_ABBR_FULLNAMES:
        full_re = re.escape(full)
        # "of <full name>" (e.g. "... increase of Homemaker") — wrong word order.
        if re.search(rf"\bof\s+{full_re}\b", low):
            _add(f"note says 'of {full}' — lead with abbreviation '{abbr}'")
        # Trailing "service"/"services" after the full name (e.g. "Chore services").
        if re.search(rf"\b{full_re}\s+services?\b", low):
            _add(f"note has '{full} service(s)' — use abbreviation '{abbr}'")
    # Full name of an authorized service present without its abbreviation.
    for svc in services:
        if not isinstance(svc, dict):
            continue
        fam = _service_family(svc.get("name"))
        pair = _NOTE_FAMILY_ABBR.get(fam or "")
        if not pair:
            continue
        full, abbr = pair
        full_low = full.lower()
        if re.search(rf"\b{re.escape(full_low)}\b", low) and \
                not re.search(rf"\b{re.escape(abbr.lower())}\b", low):
            _add(f"note uses full name '{full}' without its abbreviation '{abbr}'")

    # (b) Units instead of hours/meals — only ADH (days) and transportation
    # (trips) legitimately count in units; everything else must be hrs/wk (or
    # meals/wk for HDM).
    if re.search(r"\bunits?\b", low):
        for svc in services:
            if not isinstance(svc, dict):
                continue
            fam = _service_family(svc.get("name"))
            if not fam or fam in _NOTE_UNIT_OK_FAMILIES:
                continue
            unit = _FAMILY_UNITS.get(fam)
            if unit in ("hrs", "meals"):
                _add("note uses 'unit(s)' for an hour/meal-based service — "
                     "express it in hrs/wk (or meals/wk)")
                break

    # (c) One-time auth but the note never says "one time".
    one_time = False
    for svc in services:
        if not isinstance(svc, dict):
            continue
        uf = (svc.get("units_or_frequency") or "").lower()
        if "one-time" in uf or "one time" in uf:
            one_time = True
            break
    ct = (parsed.get("change_type") or "").lower()
    if "one-time" in ct or "one time" in ct:
        one_time = True
    if one_time and not ("one time" in low or "one-time" in low):
        _add("auth is one-time but the note never states 'one time'")

    return issues


def _is_literal_extract(parsed: dict) -> bool:
    """True when `parsed` came from the EXTRACT-ONLY prompt (mini copies
    literally; Claude decides). Old-prompt payloads carry `journal_note` text
    and services with `units_or_frequency`; new ones carry
    `notification_notes_verbatim` and services with `amount_verbatim`."""
    if not isinstance(parsed, dict):
        return False
    if "notification_notes_verbatim" in parsed:
        return True
    svc = parsed.get("services")
    return bool(isinstance(svc, list) and svc and isinstance(svc[0], dict)
                and "amount_verbatim" in svc[0])


def _compose_module():
    """compose.py lives next to this file (shipped together). Imported lazily so
    the legacy path never needs it and a missing file only affects new-schema rows."""
    import importlib.util
    p = Path(__file__).resolve().parent / "compose.py"
    spec = importlib.util.spec_from_file_location("compose", p)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def build_field_body(parsed: dict) -> dict:
    """SharePoint field body for one authorization. Schema-adaptive:
      * legacy payload (old AI Builder prompt)  -> _build_field_body_legacy, unchanged
      * literal extract (EXTRACT-ONLY prompt)   -> identifiers/dates as before,
        Services rendered literally, Notes = notification notes verbatim, and
        ChangeType / JournalNote / CarePlanComments DECIDED BY CLAUDE via
        compose.py (PHI-guarded: identifiers never reach the prompt).
    On a Claude failure those three are left EMPTY so the hourly review_notes.py
    picks the row up next run -- nothing half-written."""
    if not _is_literal_extract(parsed):
        return _build_field_body_legacy(parsed)

    body: dict = {}
    auth_no = (str(parsed.get("authorization_number") or "")).strip()
    if auth_no:
        body["Title"] = auth_no
    for key, col in TEXT_FIELD_MAP.items():
        v = parsed.get(key)
        if v not in (None, ""):
            body[col] = str(v).strip()
    for key, col in DATE_FIELD_MAP.items():
        iso = _iso_date(parsed.get(key))
        if iso:
            body[col] = iso
    decision = DECISION_MAP.get((str(parsed.get("overall_decision_verbatim") or "")
                                 .strip().lower().replace(" ", "_")))
    if decision:
        body[OVERALL_DECISION_FIELD] = decision
    addr = parsed.get("address")
    if isinstance(addr, dict):
        # new schema uses street/city/state/zip; old used street_address/zip_code
        alias = {"street_address": "street", "zip_code": "zip"}
        for key, col in ADDRESS_FIELD_MAP.items():
            v = addr.get(key)
            if v in (None, ""):
                v = addr.get(alias.get(key, key))
            if v not in (None, ""):
                body[col] = str(v).strip()

    cm = _compose_module()
    services = cm.render_services(parsed)
    if services:
        body["Services"] = services
    notes = (parsed.get("notification_notes_verbatim") or "").strip()
    if notes:
        body["Notes"] = notes

    d = cm.compose(parsed)
    logging.info("compose: sent_fields=%s redactions=%s error=%s",
                 d.get("sent_fields"), d.get("redactions"), d.get("error") or "none")
    if not d.get("error"):
        ct = (d.get("change_type") or "").strip()
        if ct:
            body[CHANGE_TYPE_FIELD] = CHANGE_TYPE_DISPLAY.get(ct.lower(), ct)
        if d.get("journal_note"):
            body["JournalNote"] = d["journal_note"]
        if d.get("summary"):
            body["CarePlanComments"] = d["summary"]
    return body


def _build_field_body_legacy(parsed: dict) -> dict:
    body: dict = {}
    auth_no = (str(parsed.get("authorization_number") or "")).strip()
    if auth_no:
        body["Title"] = auth_no
    for key, col in TEXT_FIELD_MAP.items():
        v = parsed.get(key)
        if v not in (None, ""):
            body[col] = str(v).strip()
    for key, col in DATE_FIELD_MAP.items():
        iso = _iso_date(parsed.get(key))
        if iso:
            body[col] = iso
    decision = DECISION_MAP.get((str(parsed.get("overall_decision") or "")
                                 .strip().lower()))
    if decision:
        body[OVERALL_DECISION_FIELD] = decision
    addr = parsed.get("address")
    if isinstance(addr, dict):
        for key, col in ADDRESS_FIELD_MAP.items():
            v = addr.get(key)
            if v not in (None, ""):
                body[col] = str(v).strip()
    change_type = (str(parsed.get("change_type") or "")).strip()
    if change_type and change_type.lower() != "null":
        body[CHANGE_TYPE_FIELD] = CHANGE_TYPE_DISPLAY.get(
            change_type.lower(), change_type)
    services = render_services(parsed.get("services"))
    if services:
        body["Services"] = services
    notes = (parsed.get("notes") or "").strip()
    if notes:
        body["Notes"] = notes
    journal = (parsed.get("journal_note") or "").strip()
    if journal:
        body["JournalNote"] = journal
        # Deterministic QA double-check of the note vs the structured data. FLAG
        # ONLY — never edits or blocks the note. Populates the "Note QA" column.
        issues = validate_note(parsed)
        body[NOTE_QA_FIELD] = "; ".join(issues) if issues else NOTE_QA_OK
    return body


# ── Consumer lookup ────────────────────────────────────────────────
def cm_name_uhc(raw: str | None) -> str:
    """UHC keeps the role tag on the case-manager name. Take the first manager
    if several are listed, but preserve the parenthetical role — e.g.:

      'Enrique Ponce (GSSC)'              -> 'Enrique Ponce (GSSC)'
      'Ponce, Enrique (CM); Doe, Jane'    -> 'Ponce, Enrique (CM)'
    """
    if not raw:
        return ""
    head = re.split(r"[;|]", raw, maxsplit=1)[0]
    return re.sub(r"\s+", " ", head).strip()


def _addr_match(parsed: dict, by_addr: dict) -> tuple[str, str] | None:
    """Address fallback: match on street+zip, but only accept it when a resident
    at that address is corroborated by DOB or last name — and only when that
    resolves to a SINGLE client. Returns (client_id, cm) or None. Never guesses
    between multiple residents."""
    addr = parsed.get("address")
    if not isinstance(addr, dict):
        return None
    akey = _addr_key(addr.get("street_address"), addr.get("zip_code"))
    if not akey:
        return None
    candidates = by_addr.get(akey)
    if not candidates:
        return None
    dob_key = _norm_dob(parsed.get("member_dob"))
    last_names = {last for last, _ in _name_candidates(parsed.get("member_name"))}
    corroborated = [
        c for c in candidates
        if (dob_key and c[3] == dob_key) or (c[2] and c[2] in last_names)
    ]
    if not corroborated:
        return None
    unique_clients = {c[0] for c in corroborated}
    if len(unique_clients) != 1:
        log.info("[enrich] address %s matched %d different consumers "
                 "(ambiguous) — skipping address fallback", akey,
                 len(unique_clients))
        return None
    c = corroborated[0]
    return c[0], c[1]


def lookup_consumer(parsed: dict) -> tuple[str, str, str] | None:
    """Resolve a consumer using every available signal, most-reliable first.
    Returns (client_id, cm_name, method) or None.

    Order: Health Plan ID → Medicaid ID → Name+DOB → Address (corroborated)."""
    idx = load_consumer_index()
    by_hpid = idx["by_hpid"]
    by_medicaid = idx["by_medicaid"]
    by_namedob = idx["by_namedob"]
    by_addr = idx["by_addr"]

    # 1. Health Plan ID (exact).
    hpid = (str(parsed.get("health_plan_id") or "")).strip()
    if hpid and hpid in by_hpid:
        cid, cm = by_hpid[hpid]
        return cid, cm_name_uhc(cm), "HealthPlanID"

    # 2. Medicaid ID (exact).
    mid = _norm_id(parsed.get("medicaid_id"))
    if mid and mid in by_medicaid:
        cid, cm = by_medicaid[mid]
        return cid, cm_name_uhc(cm), "Medicaid"

    # 3. Name + DOB.
    dob_key = _norm_dob(parsed.get("member_dob"))
    if dob_key:
        for last, first in _name_candidates(parsed.get("member_name")):
            hit = by_namedob.get((last, first, dob_key))
            if hit:
                cid, cm = hit
                return cid, cm_name_uhc(cm), "Name+DOB"

    # 4. Address (street+zip), corroborated by DOB or last name.
    addr_hit = _addr_match(parsed, by_addr)
    if addr_hit:
        cid, cm = addr_hit
        return cid, cm_name_uhc(cm), "Address"

    return None


# ── Main run ───────────────────────────────────────────────────────
def run(g: requests.Session, sp: requests.Session, site_id: str,
        items: list[dict], delete_duplicates: bool, reprocess: bool,
        limit: int | None, guard_existing: bool = False) -> None:
    # 0. Know which columns exist so we don't PATCH ones that aren't there yet.
    cols = existing_columns(g, site_id)
    optional_targets = (
        set(ADDRESS_FIELD_MAP.values()) | {CHANGE_TYPE_FIELD, "MedicaidID"}
    )
    missing_cols = sorted(c for c in optional_targets if c not in cols)
    if missing_cols:
        log.warning("[schema] list is missing column(s) %s — those payload "
                    "values will be SKIPPED until you add them in the list UI",
                    missing_cols)

    # 1. Parse every item once.
    parsed_by_id: dict[str, dict | None] = {}
    for it in items:
        f = it.get("fields") or {}
        parsed_by_id[it["id"]] = parse_payload(f.get(JSON_PAYLOAD_FIELD))

    # 2. Group by authorization number → canonical (oldest) vs duplicates.
    groups: dict[str, list[dict]] = {}
    no_auth: list[dict] = []
    for it in items:
        auth = auth_number_of(it, parsed_by_id[it["id"]])
        if not auth:
            no_auth.append(it)
            continue
        groups.setdefault(auth, []).append(it)

    if no_auth:
        log.info("[scan] %d item(s) have no authorization number (unparseable "
                 "payload / blank) — skipped", len(no_auth))

    canonical: list[dict] = []
    duplicates: list[tuple[dict, str, str]] = []  # (item, canonical_id, auth)
    for auth, group in groups.items():
        group.sort(key=lambda it: int(it["id"]))
        canonical.append(group[0])
        for dup in group[1:]:
            duplicates.append((dup, group[0]["id"], auth))

    # 2b. Cross-batch guard: when we fetched only unprocessed items, a canonical
    # here may still duplicate an item processed in a PRIOR run (its twin isn't
    # in this batch). Check Title existence server-side and demote any hit.
    if guard_existing and canonical:
        survivors: list[dict] = []
        cross = 0
        for it in canonical:
            auth = auth_number_of(it, parsed_by_id[it["id"]])
            exists = (title_exists(g, site_id, auth, exclude_id=it["id"])
                      if auth else False)
            if exists is True:
                duplicates.append((it, "(prior run)", auth))
                cross += 1
            else:  # False or None (check failed) -> keep, don't lose data
                survivors.append(it)
        if cross:
            log.info("[scan] %d canonical item(s) already exist from a prior "
                     "run -> treated as duplicates", cross)
        canonical = survivors

    log.info("[scan] %d unique auth number(s), %d canonical, %d duplicate(s)",
             len(groups), len(canonical), len(duplicates))

    # 3. Handle duplicates first.
    handle_duplicates(g, sp, site_id, duplicates, delete_duplicates)

    # 4. Enrich canonical items that still need it.
    def needs_processing(it: dict) -> bool:
        if reprocess:
            return True
        status = ((it.get("fields") or {}).get(LOOKUP_STATUS_FIELD) or "").strip()
        return status in ("", "Pending")

    # No-auth items can't be deduped (no auth number to key on), but they still
    # deserve full enrichment — write their descriptive fields via
    # build_field_body + attempt a consumer match. Previously they were dropped
    # entirely, which left MemberName/HealthPlanID/MemberDOB blank on any auth
    # that arrived without an authorization_number (it only ever got a ClientID
    # later from the rematch pass, which writes ID/CM only).
    enrichable = canonical + no_auth
    todo = [it for it in enrichable
            if parsed_by_id[it["id"]] and needs_processing(it)]
    if limit is not None and limit >= 0:
        todo = todo[:limit]
    log.info("[enrich] %d item(s) to process (%d canonical + %d no-auth)",
             len(todo), len(canonical), len(no_auth))

    matched = not_found = errors = 0
    for it in todo:
        item_id = it["id"]
        parsed = parsed_by_id[item_id]
        body = build_field_body(parsed)
        title = body.get("Title") or item_id

        match = None
        try:
            match = lookup_consumer(parsed)
        except Exception as e:
            log.exception("[enrich] consumer lookup failed for %s: %s", title, e)

        match_method = None
        if match:
            client_id, cm_name, match_method = match
            body["ClientID"] = client_id
            if cm_name:
                body[PRIMARY_CARE_MANAGER_FIELD] = cm_name
            body[LOOKUP_STATUS_FIELD] = "Matched"
        else:
            body[LOOKUP_STATUS_FIELD] = "Not Found"

        # Drop any field whose column doesn't exist yet (e.g. address /
        # change_type before they're added in the list UI) so the PATCH for the
        # fields that DO exist still succeeds.
        body = {k: v for k, v in body.items() if k in cols}

        try:
            patch_fields(g, site_id, item_id, body)
            if match:
                log.info("[enrich] OK %s -> ClientID=%s CM=%s (via %s)",
                         title, body["ClientID"],
                         body.get(PRIMARY_CARE_MANAGER_FIELD, "n/a"), match_method)
                matched += 1
            else:
                log.info("[enrich] %s -> Not Found (HPID=%r name=%r DOB=%r)",
                         title, parsed.get("health_plan_id"),
                         parsed.get("member_name"), parsed.get("member_dob"))
                not_found += 1
        except Exception as e:
            log.exception("[enrich] PATCH failed for %s: %s", title, e)
            errors += 1

    log.info("[enrich] done. matched=%d not_found=%d errors=%d",
             matched, not_found, errors)


def handle_duplicates(g: requests.Session, sp: requests.Session, site_id: str,
                      duplicates: list[tuple[dict, str, str]],
                      delete: bool) -> None:
    if not duplicates:
        return
    if delete:
        deleted = errors = 0
        for dup, keep_id, auth in duplicates:
            try:
                delete_item(g, site_id, dup["id"])
                log.info("[dup] deleted item %s (dup of %s, auth=%s)",
                         dup["id"], keep_id, auth)
                deleted += 1
            except Exception as e:
                log.warning("[dup] delete failed for %s: %s", dup["id"], e)
                errors += 1
        log.info("[dup] done. deleted=%d errors=%d", deleted, errors)
        return

    # Flag mode. The tenant blocks programmatic choice-column edits (Graph)
    # and app-only SharePoint REST, so we can't add a "Duplicate" Lookup Status
    # choice. Instead we stamp a marker into the (otherwise blank) Notes column
    # of each duplicate and leave it un-enriched. Re-run --delete-duplicates to
    # remove them, or add a "Duplicate" choice in the list settings UI.
    use_status = ensure_duplicate_choice(g, sp, site_id)
    flagged = errors = 0
    for dup, keep_id, auth in duplicates:
        f = dup.get("fields") or {}
        marker = f"{DUP_MARKER} of auth {auth} (canonical item {keep_id})"
        if use_status:
            body = {LOOKUP_STATUS_FIELD: DUPLICATE_STATUS}
            if (f.get(LOOKUP_STATUS_FIELD) or "").strip() == DUPLICATE_STATUS:
                continue
        else:
            if DUP_MARKER in (f.get("Notes") or ""):
                continue
            body = {"Notes": marker}
        try:
            patch_fields(g, site_id, dup["id"], body)
            log.info("[dup] flagged item %s as duplicate (dup of %s, auth=%s)",
                     dup["id"], keep_id, auth)
            flagged += 1
        except Exception as e:
            log.warning("[dup] flag failed for %s: %s", dup["id"], e)
            errors += 1
    log.info("[dup] done. flagged=%d errors=%d", flagged, errors)


# ── Re-match previously-unmatched items ────────────────────────────
def rematch_unmatched(g: requests.Session, site_id: str, items: list[dict],
                      cols: set[str], update_in_place: bool = False) -> None:
    """Re-run the consumer lookup on EVERY item still missing a Client ID —
    whether its LookupStatus is 'Not Found' or never set at all. This is what
    "populates the whole map": pass the full item list and any row that can be
    matched (now that the cascade has Medicaid + corroborated-Address signals)
    gets its Client ID / Case Manager filled in.

    Rows that already have a Client ID, duplicates, and rows with no parseable
    payload are skipped. Rows that still don't resolve are left exactly as-is.
    When `update_in_place` is True the matched values are also written back onto
    the in-memory item dict, so a report built from this same list reflects the
    fresh matches without a re-fetch."""
    fixed = still = skipped = errors = 0
    for it in items:
        f = it.get("fields") or {}
        # Guard: only touch rows that genuinely lack a client id (and CM).
        if (f.get("ClientID") or "").strip():
            skipped += 1
            continue
        if _xls_is_duplicate(it):
            skipped += 1
            continue
        parsed = parse_payload(f.get(JSON_PAYLOAD_FIELD))
        if not parsed:
            skipped += 1
            continue

        title = (f.get("Title") or "").strip() or it["id"]
        try:
            match = lookup_consumer(parsed)
        except Exception as e:
            log.exception("[rematch] lookup failed for %s: %s", title, e)
            errors += 1
            continue
        if not match:
            still += 1
            continue

        client_id, cm_name, method = match
        body = {"ClientID": client_id, LOOKUP_STATUS_FIELD: "Matched"}
        if cm_name:
            body[PRIMARY_CARE_MANAGER_FIELD] = cm_name
        body = {k: v for k, v in body.items() if k in cols}
        try:
            patch_fields(g, site_id, it["id"], body)
            if update_in_place:
                f["ClientID"] = client_id
                f[LOOKUP_STATUS_FIELD] = "Matched"
                if cm_name:
                    f[PRIMARY_CARE_MANAGER_FIELD] = cm_name
            log.info("[rematch] OK %s -> ClientID=%s CM=%s (via %s)",
                     title, client_id, cm_name or "n/a", method)
            fixed += 1
        except Exception as e:
            log.exception("[rematch] PATCH failed for %s: %s", title, e)
            errors += 1

    log.info("[rematch] done. newly_matched=%d still_unmatched=%d skipped=%d "
             "errors=%d", fixed, still, skipped, errors)


# ── Active service-suspension sync ─────────────────────────────────
def sync_suspensions(g: requests.Session, site_id: str, items: list[dict],
                     update_in_place: bool = True) -> None:
    """For every matched auth row (has a Client ID), set the 'Service Suspended'
    and 'Suspension Start Date' columns to the consumer's currently-active
    suspension — and CLEAR them when there's no longer an active suspension. The
    suspensions data changes independently of the auths (daily refresh), so this
    runs over the whole list each run and writes only the rows whose value
    actually changed, keeping churn minimal.

    Skips entirely (with a warning) if the two columns aren't in the list yet or
    the suspensions/Consumers data can't be loaded."""
    svc_col = resolve_column(g, site_id, SUSPENDED_SERVICE_COL_TITLE)
    date_col = resolve_column(g, site_id, SUSPENSION_START_COL_TITLE)
    if not svc_col or not date_col:
        log.warning("[suspend] column(s) not found in list (%r->%s, %r->%s) — "
                    "add them in the list UI; skipping suspension sync.",
                    SUSPENDED_SERVICE_COL_TITLE, svc_col,
                    SUSPENSION_START_COL_TITLE, date_col)
        return
    log.info("[suspend] columns resolved: %r->%s, %r->%s",
             SUSPENDED_SERVICE_COL_TITLE, svc_col,
             SUSPENSION_START_COL_TITLE, date_col)

    susp = load_suspension_index()  # {} is valid (clears stale flags)

    updated = cleared = skipped = errors = 0
    for it in items:
        f = it.get("fields", {}) or {}
        if _xls_is_duplicate(it):
            skipped += 1
            continue
        client_id = (f.get("ClientID") or "").strip()
        if not client_id:
            skipped += 1  # unmatched rows can't be tied to a consumer
            continue

        hit = susp.get(client_id)
        want_svc = hit[0] if hit else ""
        want_start = hit[1] if hit else None  # iso 'YYYY-MM-DD' or None

        cur_svc = _plain_text(f.get(svc_col))  # column is rich text; flatten to compare
        # The stored date comes back as 'YYYY-MM-DDZ' / 'YYYY-MM-DDT00:00:00Z';
        # take the date portion so the compare is idempotent (no needless re-writes).
        cur_raw = f.get(date_col)
        cur_start = _norm_dob(str(cur_raw)[:10]) if cur_raw else None  # iso or None

        body: dict = {}
        if want_svc != cur_svc:
            body[svc_col] = want_svc  # "" clears the text column
        if want_start != cur_start:
            body[date_col] = _iso_date(want_start)  # None clears the date column
        if not body:
            continue

        try:
            patch_fields(g, site_id, it["id"], body)
            if update_in_place:
                if svc_col in body:
                    f[svc_col] = body[svc_col]
                if date_col in body:
                    f[date_col] = body[date_col]
            if want_svc:
                updated += 1
            else:
                cleared += 1
        except Exception as e:
            log.warning("[suspend] PATCH failed for item %s: %s", it["id"], e)
            errors += 1

    log.info("[suspend] done. set/updated=%d cleared=%d skipped=%d errors=%d",
             updated, cleared, skipped, errors)


# ── Excel workbook (per-CM tabs) — vendored so this stays a single file ─
# Mirrors the SWH report: one tab per Primary Care Manager with all their
# auths, an "All Authorizations" tab covering the whole list, and a "Needs
# Review" tab for items still missing a Client ID. The File column links to the
# SharePoint file. The output filename is FIXED (no date) so each daily run
# overwrites it in place — "the same name so it gets updated every day".
_XLS_COLUMNS: list[tuple[str, str, str]] = [
    ("Title", "Auth #", "text"),
    ("MemberName", "Member", "text"),
    ("MemberDOB", "DOB", "date"),
    ("__address__", "Address", "address"),
    ("HealthPlanID", "Health Plan ID", "text"),
    ("MedicaidID", "Medicaid ID", "text"),
    ("ClientID", "Client ID", "text"),
    ("PrimaryCareManager", "Case Manager", "text"),
    ("Services", "Services", "text"),
    ("ChangeType", "Change Type", "text"),
    (SUSP_SVC_FIELD, "Service Suspended", "richtext"),
    (SUSP_START_FIELD, "Suspension Start", "date"),
    (PLAN_DISC_FIELD, "Plan Discrepancy", "richtext"),
    (PLAN_DISC_DETAIL_FIELD, "Plan Discrepancy Detail", "richtext"),
    ("AuthPeriodStart", "Auth Start", "date"),
    ("AuthPeriodEnd", "Auth End", "date"),
    ("OverallDecision", "Decision", "text"),
    ("ReviewDate", "Review Date", "date"),
    ("Notes", "Notes", "text"),
    ("JournalNote", "Journal Note", "text"),
    ("__pdf__", "PDF", "itemlink"),
]
_XLS_FILE_COL_IDX = next(i for i, (k, _, _) in enumerate(_XLS_COLUMNS, 1)
                         if k == "__pdf__")
_XLS_CM_FIELD = "PrimaryCareManager"
_XLS_NEEDS_REVIEW_TAB = "Needs Review"
_XLS_INVALID_SHEET_CHARS = ":\\/?*[]"
_XLS_HEADER_FONT = Font(bold=True, color="FFFFFFFF")
_XLS_HEADER_FILL = PatternFill("solid", fgColor="FF1F4E78")
_XLS_LINK_FONT = Font(color="FF0563C1", underline="single")

# Fixed output name (overwritten daily). Sits next to this script.
_XLS_OUT_NAME = "uhc_authorizations.xlsx"


def _xls_out_path() -> Path:
    return SCRIPT_DIR / _XLS_OUT_NAME


def _xls_safe_sheet_name(name: str, used: set[str]) -> str:
    """Return an Excel-safe, unique sheet name <= 31 chars."""
    cleaned = name or "Unknown"
    for ch in _XLS_INVALID_SHEET_CHARS:
        cleaned = cleaned.replace(ch, "-")
    cleaned = cleaned[:31].strip() or "Unknown"
    out = cleaned
    i = 2
    while out in used:
        suffix = f" ({i})"
        out = (cleaned[: 31 - len(suffix)]).strip() + suffix
        i += 1
    used.add(out)
    return out


def _dispform_url(item_id) -> str:
    """Classic item display-form URL. Opens the list item (including its PDF
    attachment as a download link). Used instead of a direct file URL because
    list-item attachments aren't reachable with this script's app-only token."""
    return (f"{SP_SITE_URL}/Lists/{quote(SP_LIST_TITLE)}"
            f"/DispForm.aspx?ID={item_id}")


def _xls_get_url(field_value) -> str:
    if isinstance(field_value, dict):
        return (field_value.get("Url") or field_value.get("url") or "").strip()
    if isinstance(field_value, str):
        return field_value.strip()
    return ""


def _xls_fmt_date(s) -> str:
    if not s:
        return ""
    s = str(s)
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).strftime("%m/%d/%Y")
    except Exception:
        return s.split("T")[0]


def _xls_address(f: dict) -> str:
    """One-line 'street, city, state zip' from the separate address columns."""
    street = (f.get("StreetAddress") or "").strip()
    city = (f.get("City") or "").strip()
    state = (f.get("State") or "").strip()
    zip_code = (f.get("ZipCode") or "").strip()
    locality = ", ".join(p for p in (city, state) if p)
    if zip_code:
        locality = (locality + " " + zip_code).strip()
    return ", ".join(p for p in (street, locality) if p)


def _xls_write_rows(ws, items: list[dict]) -> None:
    headers = [h for _, h, _ in _XLS_COLUMNS]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.font = _XLS_HEADER_FONT
        c.fill = _XLS_HEADER_FILL
        c.alignment = Alignment(vertical="center")

    for it in items:
        f = it.get("fields", {}) or {}
        row: list[str] = []
        for key, _, kind in _XLS_COLUMNS:
            if kind == "url":
                row.append(_xls_get_url(f.get(key)))
            elif kind == "date":
                row.append(_xls_fmt_date(f.get(key)))
            elif kind == "address":
                row.append(_xls_address(f))
            elif kind == "richtext":
                row.append(_plain_text(f.get(key)))
            elif kind == "itemlink":
                row.append("Open PDF" if f.get("Attachments") else "")
            else:
                v = f.get(key)
                row.append("" if v is None else str(v))
        ws.append(row)

        if f.get("Attachments"):
            cell = ws.cell(row=ws.max_row, column=_XLS_FILE_COL_IDX)
            cell.hyperlink = _dispform_url(it["id"])
            cell.font = _XLS_LINK_FONT

    for col_idx in range(1, len(headers) + 1):
        max_len = len(headers[col_idx - 1])
        for row in ws.iter_rows(
            min_col=col_idx, max_col=col_idx, min_row=2, values_only=True
        ):
            v = "" if row[0] is None else str(row[0])
            if len(v) > max_len:
                max_len = len(v)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    ws.freeze_panes = "A2"


def _xls_is_duplicate(it: dict) -> bool:
    """A flagged duplicate / blank row we don't want in the report. Duplicates
    are never enriched, so their SP fields stay empty (only JSONPayload + a
    Notes marker are set) — they'd otherwise show as blank rows. We recognize
    them by the Lookup Status = 'Duplicate' choice or the '[DUPLICATE]' Notes
    marker the pipeline stamps on them."""
    f = it.get("fields", {}) or {}
    if (f.get(LOOKUP_STATUS_FIELD) or "").strip() == DUPLICATE_STATUS:
        return True
    if DUP_MARKER in (f.get("Notes") or ""):
        return True
    return False


def _xls_dedupe(items: list[dict]) -> list[dict]:
    """Drop flagged duplicates, blank/no-auth rows, and collapse any remaining
    same-Title items to the canonical (lowest ID) — so the report has exactly
    one clean row per authorization."""
    seen: dict[str, dict] = {}
    out: list[dict] = []
    for it in items:
        if _xls_is_duplicate(it):
            continue
        f = it.get("fields", {}) or {}
        title = (f.get("Title") or "").strip()
        if not title:
            continue  # no auth number -> blank row, skip
        prev = seen.get(title)
        if prev is None:
            seen[title] = it
            out.append(it)
        elif int(it["id"]) < int(prev["id"]):
            # Keep the canonical (oldest) row for this auth number.
            out[out.index(prev)] = it
            seen[title] = it
    return out


def build_workbook(items: list[dict], out_path: Path | None = None) -> Path:
    """Build the per-CM workbook from an already-fetched items list."""
    if out_path is None:
        out_path = _xls_out_path()
    before = len(items)
    items = _xls_dedupe(items)
    if before != len(items):
        log.info("[excel] dropped %d duplicate/blank row(s); %d remain",
                 before - len(items), len(items))
    by_cm: dict[str, list[dict]] = {}
    needs_review: list[dict] = []
    for it in items:
        f = it.get("fields", {}) or {}
        if not (f.get("ClientID") or "").strip():
            needs_review.append(it)
            continue
        cm = (f.get(_XLS_CM_FIELD) or "").strip() or "Unassigned"
        by_cm.setdefault(cm, []).append(it)

    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()

    all_sheet_name = _xls_safe_sheet_name("All Authorizations", used)
    ws_all = wb.create_sheet(all_sheet_name)
    _xls_write_rows(ws_all, items)
    log.info("[excel] sheet %r: %d row(s)", all_sheet_name, len(items))

    for cm in sorted(by_cm, key=str.lower):
        sheet_name = _xls_safe_sheet_name(cm, used)
        ws = wb.create_sheet(sheet_name)
        _xls_write_rows(ws, by_cm[cm])
        log.info("[excel] sheet %r: %d row(s)", sheet_name, len(by_cm[cm]))

    if needs_review:
        sheet_name = _xls_safe_sheet_name(_XLS_NEEDS_REVIEW_TAB, used)
        ws = wb.create_sheet(sheet_name)
        _xls_write_rows(ws, needs_review)
        log.info("[excel] sheet %r: %d row(s)", sheet_name, len(needs_review))

    wb.save(out_path)
    log.info("[excel] wrote %s", out_path)
    return out_path


# ── Interactive calendar (by item creation date) — vendored ────────
# A single self-contained HTML file: a month-grid calendar keyed on the date
# each authorization was *received* (the list item's createdDateTime). Each day
# with auths shows a count; clicking it opens a panel listing every
# authorization received that day. No external assets, no network — the data is
# embedded as JSON so the file works offline from OneDrive.
_CAL_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>UHC Authorizations — Calendar</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Geist:wght@300..700&family=Geist+Mono:wght@400..600&display=swap" rel="stylesheet">
<style>
:root {
  /* CBES blue theme — same Bamboo layout, cool palette */
  --bg: #eef2f8;
  --surface: #ffffff;
  --surface-2: #f3f6fb;
  --border: #d3dcea;
  --border-soft: #e4ebf5;
  --text: #12233f;
  --text-2: #4d5d76;
  --text-3: #95a3b8;
  --accent: #1c75bc;
  --accent-hover: #10306e;
  --accent-soft: rgba(28, 117, 188, 0.10);

  /* Decision colors — CBES-aligned (green / gold / red / slate) */
  --dec-approved: #03b581;
  --dec-partial: #d99211;
  --dec-denied: #c0334b;
  --dec-other: #5b6b86;

  /* Change-type urgency flags (time-critical, separate from decision) */
  --flag-rr: #7c5cff;    /* Records Request — appeal records due fast */
  --flag-term: #e8590c;  /* Termination — service being stopped */
  --flag-susp: #1098ad;  /* Suspension — consumer has an active service suspension */
  --flag-disc: #c0299a;  /* Plan Discrepancy — auth disagrees with the active service plan */

  --shadow-sm: 0 1px 2px rgba(16, 48, 110, 0.05);
  --shadow-md: 0 4px 20px -6px rgba(16, 48, 110, 0.10);
  --ease: cubic-bezier(0.16, 1, 0.3, 1);
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html, body { height: 100%; }
body {
  font-family: 'Geist', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  font-size: 14px;
  line-height: 1.5;
  color: var(--text);
  background: var(--bg);
  -webkit-font-smoothing: antialiased;
  letter-spacing: -0.011em;
}
button { font-family: inherit; font-size: inherit; cursor: pointer; }
a { color: inherit; }

.app {
  max-width: 1400px;
  margin: 0 auto;
  padding: 32px 40px 64px;
  min-height: 100dvh;
  display: flex;
  flex-direction: column;
}

/* ===== Header ===== */
.header {
  display: grid;
  grid-template-columns: 1fr auto 1fr;
  gap: 48px;
  align-items: center;
  padding-bottom: 28px;
  border-bottom: 1px solid var(--border);
  margin-bottom: 28px;
}
.brand { display: flex; flex-direction: column; gap: 2px; }
.brand-eyebrow {
  font-family: 'Geist Mono', monospace;
  font-size: 11px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--text-3);
}
.brand-title {
  font-size: 22px;
  font-weight: 600;
  letter-spacing: -0.025em;
}

.view-toggle {
  display: inline-flex;
  padding: 3px;
  background: var(--surface-2);
  border-radius: 11px;
  border: 1px solid var(--border);
}
.view-toggle button {
  border: none;
  background: transparent;
  padding: 7px 18px;
  font-size: 13px;
  font-weight: 500;
  color: var(--text-2);
  border-radius: 8px;
  transition: all 0.2s var(--ease);
}
.view-toggle button[aria-selected="true"] {
  background: var(--surface);
  color: var(--text);
  box-shadow: var(--shadow-sm);
}
.view-toggle button:hover:not([aria-selected="true"]) { color: var(--text); }
.view-toggle button:active { transform: translateY(0.5px); }

.stats {
  justify-self: end;
  display: flex;
  flex-direction: column;
  align-items: flex-end;
  gap: 2px;
}
.stats-value {
  font-family: 'Geist Mono', monospace;
  font-size: 24px;
  font-weight: 500;
  letter-spacing: -0.03em;
  line-height: 1;
}
.stats-label {
  font-family: 'Geist Mono', monospace;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--text-3);
}

/* ===== Nav row ===== */
.nav-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 20px;
  gap: 24px;
  flex-wrap: wrap;
}
.period-nav { display: flex; align-items: center; gap: 14px; }
.icon-btn {
  width: 32px;
  height: 32px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: 1px solid var(--border);
  background: var(--surface);
  border-radius: 8px;
  color: var(--text-2);
  transition: all 0.2s var(--ease);
}
.icon-btn:hover { background: var(--surface-2); color: var(--text); }
.icon-btn:active { transform: scale(0.94); }
.period-label {
  font-size: 16px;
  font-weight: 600;
  letter-spacing: -0.015em;
  min-width: 220px;
  text-align: center;
}

.legend {
  display: flex;
  align-items: center;
  gap: 18px;
  font-size: 12px;
  color: var(--text-2);
  flex-wrap: wrap;
}
.legend-item { display: inline-flex; align-items: center; gap: 6px; }
.legend-dot { width: 7px; height: 7px; border-radius: 50%; }

/* ===== Month view ===== */
.calendar {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 20px;
  overflow: hidden;
  box-shadow: var(--shadow-md);
}
.weekdays {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  border-bottom: 1px solid var(--border);
  background: var(--surface-2);
}
.weekday {
  padding: 12px 16px;
  font-family: 'Geist Mono', monospace;
  font-size: 10px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--text-3);
}
.days-grid { display: grid; grid-template-columns: repeat(7, 1fr); }
.day-cell {
  aspect-ratio: 1 / 0.85;
  border-right: 1px solid var(--border);
  border-bottom: 1px solid var(--border);
  padding: 10px 12px 10px;
  display: flex;
  flex-direction: column;
  gap: 4px;
  position: relative;
  background: var(--surface);
  transition: background 0.2s var(--ease), transform 0.15s var(--ease);
  user-select: none;
  animation: fadeUp 0.4s var(--ease) both;
  animation-delay: calc(var(--i, 0) * 6ms);
}
.day-cell:nth-child(7n) { border-right: none; }
.day-cell.other-month { background: var(--bg); }
.day-cell.other-month .day-num { color: var(--text-3); }
.day-cell.has-data { cursor: pointer; }
.day-cell.has-data:hover { background: var(--accent-soft); }
.day-cell.has-data:active { transform: scale(0.985); }
.day-cell.today .day-num {
  color: var(--surface);
  background: var(--accent);
  width: 26px;
  height: 26px;
  border-radius: 50%;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-weight: 600;
  font-size: 13px;
}
.day-num {
  font-family: 'Geist Mono', monospace;
  font-size: 15px;
  font-weight: 600;
  color: var(--text);
  line-height: 1;
  letter-spacing: -0.01em;
}
.visits-block {
  margin-top: auto;
  display: flex;
  flex-direction: column;
  gap: 2px;
  line-height: 1;
}
.visits-count {
  font-family: 'Geist', sans-serif;
  font-size: 17px;
  font-weight: 600;
  letter-spacing: -0.02em;
  color: var(--text);
}
.visits-label {
  font-family: 'Geist Mono', monospace;
  font-size: 9px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.14em;
  color: var(--text-3);
}
.other-month .visits-count { color: var(--text-3); }
.status-bars {
  display: flex;
  gap: 2px;
  height: 3px;
  border-radius: 2px;
  overflow: hidden;
  margin-top: 6px;
}
.status-bar { height: 100%; min-width: 4px; }

/* ===== Week view ===== */
.week-grid {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  gap: 10px;
}
.week-day {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 14px;
  min-height: 360px;
  padding: 16px;
  display: flex;
  flex-direction: column;
  gap: 12px;
  transition: all 0.2s var(--ease);
  animation: fadeUp 0.4s var(--ease) both;
  animation-delay: calc(var(--i, 0) * 30ms);
}
.week-day.has-data { cursor: pointer; }
.week-day.has-data:hover {
  border-color: var(--accent);
  box-shadow: var(--shadow-md);
}
.week-day.has-data:active { transform: translateY(0.5px); }
.week-day-header { display: flex; flex-direction: column; gap: 2px; }
.week-day-name {
  font-family: 'Geist Mono', monospace;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--text-3);
}
.week-day-num {
  font-size: 22px;
  font-weight: 600;
  letter-spacing: -0.025em;
  line-height: 1;
}
.week-day-total {
  font-family: 'Geist Mono', monospace;
  font-size: 11px;
  color: var(--text-2);
  margin-top: 2px;
}
.week-stats {
  display: flex;
  flex-direction: column;
  gap: 6px;
  padding: 10px 0;
  border-top: 1px solid var(--border-soft);
  border-bottom: 1px solid var(--border-soft);
}
.week-stat-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  font-size: 12px;
}
.week-stat-label {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: var(--text-2);
}
.week-stat-count {
  font-family: 'Geist Mono', monospace;
  font-weight: 500;
  color: var(--text);
}
.week-patients {
  font-size: 12px;
  color: var(--text-2);
  display: flex;
  flex-direction: column;
  gap: 3px;
  overflow: hidden;
  flex: 1;
}
.week-patients .more { color: var(--text-3); font-family: 'Geist Mono', monospace; font-size: 11px; }
.week-empty {
  font-family: 'Geist Mono', monospace;
  font-size: 11px;
  color: var(--text-3);
  text-transform: uppercase;
  letter-spacing: 0.1em;
}

/* ===== Day view ===== */
.day-view {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 20px;
  overflow: hidden;
  box-shadow: var(--shadow-md);
}
.day-view-header {
  padding: 28px 32px 24px;
  border-bottom: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 24px;
  flex-wrap: wrap;
}
.day-view-title {
  font-size: 24px;
  font-weight: 600;
  letter-spacing: -0.025em;
  line-height: 1.2;
}
.day-view-subtitle {
  font-size: 13px;
  color: var(--text-2);
  margin-top: 4px;
  font-family: 'Geist Mono', monospace;
}
.day-badges { display: flex; gap: 6px; flex-wrap: wrap; }
.badge {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  padding: 6px 12px;
  border-radius: 100px;
  font-size: 12px;
  background: var(--surface-2);
  border: 1px solid var(--border);
  cursor: pointer;
  transition: all 0.18s var(--ease);
  user-select: none;
}
.badge:hover {
  background: var(--surface);
  border-color: var(--text-3);
}
.badge:active { transform: scale(0.97); }
.badge.active {
  background: var(--accent-soft);
  border-color: var(--accent);
  color: var(--accent-hover);
}
.badge.dimmed { opacity: 0.4; }
.badge.dimmed:hover { opacity: 1; }
.badge-dot { width: 6px; height: 6px; border-radius: 50%; }
.badge strong {
  font-family: 'Geist Mono', monospace;
  font-weight: 600;
  margin-left: 2px;
}

/* Urgency flags (Records Request / Termination / Suspension) — change-type + active-suspension driven */
.day-flags {
  position: absolute;
  top: 6px;
  right: 6px;
  display: flex;
  flex-direction: column;
  gap: 3px;
  align-items: flex-end;
  z-index: 1;
  pointer-events: none;
}
.day-flag {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 2px 6px;
  border-radius: 100px;
  font-family: 'Geist Mono', monospace;
  font-size: 9px;
  font-weight: 600;
  letter-spacing: 0.03em;
  color: #fff;
  line-height: 1.45;
  box-shadow: var(--shadow-sm);
}
.flag-pill {
  display: inline-flex;
  align-items: center;
  padding: 3px 9px;
  border-radius: 100px;
  font-size: 11px;
  font-weight: 600;
  white-space: nowrap;
  color: var(--fc, var(--text-2));
  background: color-mix(in srgb, var(--fc, var(--text-3)) 13%, transparent);
  border: 1px solid color-mix(in srgb, var(--fc, var(--text-3)) 36%, transparent);
}
.badge.flag-badge.active {
  background: color-mix(in srgb, var(--fc, var(--accent)) 14%, transparent);
  border-color: var(--fc, var(--accent));
  color: var(--text);
}
.week-flags { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px; }

.filter-clear {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 6px 10px 6px 12px;
  border-radius: 100px;
  font-size: 11px;
  font-family: 'Geist Mono', monospace;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  background: transparent;
  border: 1px dashed var(--border);
  color: var(--text-2);
  cursor: pointer;
  transition: all 0.18s var(--ease);
}
.filter-clear:hover {
  color: var(--text);
  border-color: var(--text-3);
}
.filter-clear svg { width: 11px; height: 11px; }

/* ===== Global search ===== */
.search-wrap { position: relative; display: flex; align-items: center; gap: 8px; }
.search-box {
  font-family: inherit;
  font-size: 13px;
  color: var(--text);
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 8px 30px 8px 32px;
  min-width: 260px;
  transition: border-color 0.15s var(--ease);
  background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%234d5d76' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='11' cy='11' r='8'/><line x1='21' y1='21' x2='16.65' y2='16.65'/></svg>");
  background-repeat: no-repeat;
  background-position: left 10px center;
  background-size: 14px;
}
.search-box:focus { outline: 2px solid var(--accent); outline-offset: 1px; border-color: var(--accent); }
.search-box::placeholder { color: var(--text-3); }
.search-clear-x {
  position: absolute; right: 7px; top: 50%; transform: translateY(-50%);
  border: none; background: none; color: var(--text-3); cursor: pointer;
  display: none; padding: 3px; line-height: 0; border-radius: 6px;
}
.search-clear-x:hover { color: var(--text); background: var(--surface-2); }
.search-clear-x svg { width: 13px; height: 13px; }
.search-view-head {
  display: flex; align-items: baseline; gap: 10px;
  padding: 4px 4px 16px; flex-wrap: wrap;
}
.search-view-title { font-size: 18px; font-weight: 600; letter-spacing: -0.02em; }
.search-view-sub { font-family: 'Geist Mono', monospace; font-size: 11px; color: var(--text-3); }
.search-date-head {
  font-family: 'Geist Mono', monospace; font-size: 11px; font-weight: 600;
  text-transform: uppercase; letter-spacing: 0.08em; color: var(--text-2);
  padding: 16px 4px 6px;
}
.pivot-table td.svc {
  white-space: pre-line;
  font-size: 12px;
  color: var(--text-2);
  line-height: 1.45;
  min-width: 200px;
  max-width: 340px;
}

.filter-bar {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 14px 32px;
  border-bottom: 1px solid var(--border-soft);
  background: var(--surface-2);
  flex-wrap: wrap;
}
.filter-bar-label {
  font-family: 'Geist Mono', monospace;
  font-size: 10px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--text-3);
}
.filter-select {
  font-family: inherit;
  font-size: 13px;
  color: var(--text);
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 6px 30px 6px 12px;
  cursor: pointer;
  appearance: none;
  -webkit-appearance: none;
  background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%234d5d76' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><polyline points='6 9 12 15 18 9'/></svg>");
  background-repeat: no-repeat;
  background-position: right 10px center;
  background-size: 12px;
  transition: border-color 0.15s var(--ease);
  min-width: 240px;
  max-width: 360px;
}
.filter-select:hover { border-color: var(--text-3); }
.filter-select:focus { outline: 2px solid var(--accent); outline-offset: 1px; border-color: var(--accent); }

/* CM grouping inside the day view */
.cm-section + .cm-section { border-top: 8px solid var(--surface-2); }
.cm-section-head {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 14px 32px 10px;
}
.cm-section-name {
  font-size: 14px;
  font-weight: 600;
  letter-spacing: -0.01em;
  color: var(--text);
}
.cm-section-head.unmatched .cm-section-name { color: var(--dec-denied); }
.cm-section-count {
  font-family: 'Geist Mono', monospace;
  font-size: 11px;
  font-weight: 600;
  color: var(--text-2);
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: 100px;
  padding: 2px 10px;
}
.cm-section-head.unmatched .cm-section-count { color: var(--dec-denied); border-color: rgba(136,19,55,0.25); }

.pivot-wrap { overflow-x: auto; }
.pivot-table { width: 100%; border-collapse: collapse; }
.pivot-table th {
  text-align: left;
  padding: 10px 16px;
  font-family: 'Geist Mono', monospace;
  font-size: 10px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--text-3);
  background: var(--surface-2);
  border-bottom: 1px solid var(--border);
  border-top: 1px solid var(--border-soft);
  white-space: nowrap;
}
.pivot-table th.num, .pivot-table td.num { text-align: center; }
.pivot-table td {
  padding: 11px 16px;
  border-bottom: 1px solid var(--border-soft);
  font-size: 13px;
  vertical-align: middle;
}
.pivot-table tbody tr { transition: background 0.12s var(--ease); cursor: pointer; }
.pivot-table tbody tr:hover td { background: var(--accent-soft); }
.pivot-table .patient-id {
  font-family: 'Geist Mono', monospace;
  font-size: 12px;
  color: var(--text-2);
  white-space: nowrap;
}
.dec-pill {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 3px 9px;
  border-radius: 100px;
  font-size: 11.5px;
  font-weight: 500;
  background: var(--surface-2);
  border: 1px solid var(--border);
  white-space: nowrap;
}
.dec-dot { width: 6px; height: 6px; border-radius: 50%; }

/* ===== Modal ===== */
.modal-backdrop {
  position: fixed;
  inset: 0;
  background: rgba(28, 25, 23, 0.45);
  backdrop-filter: blur(4px);
  -webkit-backdrop-filter: blur(4px);
  display: none;
  align-items: center;
  justify-content: center;
  z-index: 1000;
  padding: 24px;
  animation: modalFade 0.18s var(--ease);
}
.modal-backdrop.open { display: flex; }
@keyframes modalFade { from { opacity: 0; } to { opacity: 1; } }
.modal {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 20px;
  box-shadow: 0 24px 60px -12px rgba(28, 25, 23, 0.25);
  width: 100%;
  max-width: 820px;
  max-height: 88vh;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  animation: modalPop 0.22s var(--ease);
}
@keyframes modalPop { from { opacity: 0; transform: translateY(8px) scale(0.98); } to { opacity: 1; transform: translateY(0) scale(1); } }
.modal-header {
  padding: 22px 28px 18px;
  border-bottom: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 16px;
}
.modal-eyebrow {
  font-family: 'Geist Mono', monospace;
  font-size: 10px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--text-3);
  margin-bottom: 4px;
}
.modal-title {
  font-size: 20px;
  font-weight: 600;
  letter-spacing: -0.02em;
  line-height: 1.25;
}
.modal-subtitle {
  font-size: 12px;
  color: var(--text-2);
  margin-top: 4px;
  font-family: 'Geist Mono', monospace;
}
.modal-close {
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: 10px;
  width: 34px;
  height: 34px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  color: var(--text-2);
  transition: all 0.15s var(--ease);
  flex-shrink: 0;
}
.modal-close:hover { background: var(--surface); color: var(--text); border-color: var(--text-3); }
.modal-close svg { width: 14px; height: 14px; }
.modal-body {
  padding: 20px 28px 24px;
  overflow-y: auto;
  flex: 1;
}
.modal-event-head {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 14px;
  flex-wrap: wrap;
}
.modal-event-head .status-pill {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  padding: 4px 10px;
  border-radius: 100px;
  font-size: 11px;
  font-family: 'Geist Mono', monospace;
  font-weight: 500;
  background: var(--surface-2);
  border: 1px solid var(--border);
}
.modal-event-head .status-pill .status-dot { width: 6px; height: 6px; border-radius: 50%; }
.modal-event-head .modal-event-meta {
  font-family: 'Geist Mono', monospace;
  font-size: 11px;
  color: var(--text-3);
}
.modal-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px 24px;
}
.modal-field { display: flex; flex-direction: column; gap: 3px; min-width: 0; }
.modal-field .label {
  font-family: 'Geist Mono', monospace;
  font-size: 10px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.1em;
  color: var(--text-3);
}
.modal-field .value {
  font-size: 13px;
  color: var(--text);
  word-break: break-word;
  white-space: pre-wrap;
}
.modal-field.full { grid-column: 1 / -1; }
.pdf-link {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  margin-top: 16px;
  font-size: 12.5px;
  font-weight: 600;
  color: var(--accent-hover);
  text-decoration: none;
}
.pdf-link:hover { text-decoration: underline; }
@media (max-width: 640px) {
  .modal-grid { grid-template-columns: 1fr; }
  .modal-header, .modal-body { padding-left: 20px; padding-right: 20px; }
}

/* ===== Empty states ===== */
.empty {
  padding: 80px 40px;
  text-align: center;
}
.empty-title {
  font-size: 17px;
  font-weight: 500;
  color: var(--text);
  margin-bottom: 6px;
  letter-spacing: -0.015em;
}
.empty-hint { font-size: 13px; color: var(--text-2); }

/* ===== Footer ===== */
.footer {
  margin-top: 40px;
  padding-top: 20px;
  border-top: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  font-family: 'Geist Mono', monospace;
  font-size: 10px;
  color: var(--text-3);
  text-transform: uppercase;
  letter-spacing: 0.12em;
  flex-wrap: wrap;
  gap: 12px;
}

/* ===== Transitions ===== */
.view-section { animation: fadeIn 0.25s var(--ease); }
@keyframes fadeIn { from { opacity: 0; transform: translateY(2px); } to { opacity: 1; transform: translateY(0); } }
@keyframes fadeUp { from { opacity: 0; transform: translateY(3px); } to { opacity: 1; transform: translateY(0); } }

/* ===== Responsive ===== */
@media (max-width: 1024px) {
  .week-grid { grid-template-columns: repeat(2, 1fr); }
}
@media (max-width: 768px) {
  .app { padding: 20px 16px 48px; }
  .header { grid-template-columns: 1fr; gap: 16px; padding-bottom: 20px; margin-bottom: 20px; }
  .stats { justify-self: start; align-items: flex-start; }
  .view-toggle { justify-self: start; }
  .nav-row { flex-direction: column; align-items: flex-start; }
  .period-label { min-width: auto; text-align: left; }
  .week-grid { grid-template-columns: 1fr; }
  .day-cell { padding: 6px 8px; aspect-ratio: 1 / 1; }
  .day-num { font-size: 11px; }
  .weekday { padding: 8px 6px; font-size: 9px; }
  .cm-section-head, .day-view-header, .filter-bar { padding-left: 18px; padding-right: 18px; }
}

/* ===== CBES blue emphasis ===== */
body { background: #e7eef8; }
.app { position: relative; }
.header { border-bottom-color: #cdd9ec; }
.brand-title { color: var(--accent-hover); }
.stats-value { color: var(--accent); }
.period-label { color: var(--accent-hover); }
.view-toggle button[aria-selected="true"] { color: var(--accent-hover); }
.calendar { border-color: #c6d5ea; }
.weekdays { background: linear-gradient(135deg, var(--accent), var(--accent-hover)); border-bottom: none; }
.weekday { color: rgba(255, 255, 255, 0.92); }
.day-view-header { background: linear-gradient(135deg, var(--accent), var(--accent-hover)); border-bottom: none; }
.day-view-title { color: #ffffff; }
.day-view-subtitle { color: rgba(255, 255, 255, 0.85); }
.modal-header { background: linear-gradient(135deg, var(--accent-hover), var(--accent)); border-bottom: none; }
.modal-eyebrow { color: rgba(255, 255, 255, 0.78); }
.modal-title { color: #ffffff; }
.modal-subtitle { color: rgba(255, 255, 255, 0.85); }
.modal-close { background: rgba(255, 255, 255, 0.16); border-color: rgba(255, 255, 255, 0.32); color: #ffffff; }
.modal-close:hover { background: rgba(255, 255, 255, 0.30); color: #ffffff; border-color: rgba(255, 255, 255, 0.55); }
.footer { border-top-color: #cdd9ec; }
</style>
</head>
<body>
<div class="app">
  <header class="header">
    <div class="brand">
      <div class="brand-eyebrow">United Health Care</div>
      <div class="brand-title">Authorizations</div>
    </div>
    <div class="view-toggle" role="tablist">
      <button data-view="month" aria-selected="true">Month</button>
      <button data-view="week" aria-selected="false">Week</button>
      <button data-view="day" aria-selected="false">Day</button>
    </div>
    <div class="stats">
      <span class="stats-value" id="stats-count">0</span>
      <span class="stats-label" id="stats-label">authorizations</span>
    </div>
  </header>

  <div class="nav-row">
    <div class="period-nav">
      <button class="icon-btn" id="prev-btn" aria-label="Previous period">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg>
      </button>
      <div class="period-label" id="period-label">—</div>
      <button class="icon-btn" id="next-btn" aria-label="Next period">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg>
      </button>
    </div>
    <div class="legend" id="legend"></div>
    <div class="search-wrap">
      <input type="search" id="global-search" class="search-box" autocomplete="off" spellcheck="false"
             placeholder="Search name, client ID, auth #, Medicaid #…" aria-label="Search authorizations">
      <button class="search-clear-x" id="search-clear-x" aria-label="Clear search">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button>
    </div>
  </div>

  <main style="flex: 1;">
    <section id="month-view" class="view-section">
      <div class="calendar">
        <div class="weekdays">
          <div class="weekday">Sun</div><div class="weekday">Mon</div><div class="weekday">Tue</div><div class="weekday">Wed</div><div class="weekday">Thu</div><div class="weekday">Fri</div><div class="weekday">Sat</div>
        </div>
        <div class="days-grid" id="days-grid"></div>
      </div>
    </section>

    <section id="week-view" class="view-section" hidden>
      <div class="week-grid" id="week-grid"></div>
    </section>

    <section id="day-view" class="view-section" hidden>
      <div class="day-view">
        <div class="day-view-header">
          <div>
            <div class="day-view-title" id="day-title">—</div>
            <div class="day-view-subtitle" id="day-subtitle">—</div>
          </div>
          <div class="day-badges" id="day-badges"></div>
        </div>
        <div class="filter-bar" id="filter-bar" hidden>
          <span class="filter-bar-label">Case Manager</span>
          <select class="filter-select" id="filter-practice" aria-label="Filter by Case Manager">
            <option value="">All</option>
          </select>
        </div>
        <div id="day-table-container"></div>
      </div>
    </section>

    <section id="search-view" class="view-section" hidden>
      <div class="day-view">
        <div class="search-view-head">
          <span class="search-view-title" id="search-title">Search results</span>
          <span class="search-view-sub" id="search-sub">—</span>
        </div>
        <div id="search-results-container"></div>
      </div>
    </section>
  </main>

  <footer class="footer">
    <div>Source · United Health Care Authorizations</div>
    <div id="generated-at">—</div>
  </footer>
</div>

<div class="modal-backdrop" id="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="modal-title">
  <div class="modal" role="document">
    <div class="modal-header">
      <div>
        <div class="modal-eyebrow" id="modal-eyebrow">Authorization Detail</div>
        <div class="modal-title" id="modal-title">—</div>
        <div class="modal-subtitle" id="modal-subtitle">—</div>
      </div>
      <button class="modal-close" id="modal-close" aria-label="Close">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button>
    </div>
    <div class="modal-body" id="modal-body"></div>
  </div>
</div>

<script>
const RECORDS = __DATA__;
const GENERATED = "__GENERATED__";

// ---- Decision buckets / colors ----
function decBucket(dec){
  const d = (dec || "").toLowerCase();
  if (d.includes("partial")) return "Partially Approved";
  if (d.includes("approv")) return "Approved";
  if (d.includes("deni") || d.includes("reject")) return "Denied";
  return "Other";
}
const DEC_ORDER = ["Approved", "Partially Approved", "Denied", "Other"];
const DEC_COLOR = {
  "Approved": "var(--dec-approved)",
  "Partially Approved": "var(--dec-partial)",
  "Denied": "var(--dec-denied)",
  "Other": "var(--dec-other)"
};
const UN = "__UN__";
function cmKeyOf(r){ return (r.cm || "").trim() || UN; }
function cmLabel(k){ return k === UN ? "Unmatched — no case manager" : k; }

// ---- Change-type urgency flags (time-critical docs that aren't normal auths) ----
const FLAG_DEFS = {
  "Records Request": { short: "REC REQ", color: "var(--flag-rr)" },
  "Termination":     { short: "TERM",    color: "var(--flag-term)" },
  "Suspension":      { short: "SUSP",    color: "var(--flag-susp)" },
};
const FLAG_ORDER = Object.keys(FLAG_DEFS);
// A row is flagged by its change-type if that's a known urgency flag; otherwise,
// if the consumer has an active service suspension (joined from the Service
// Suspensions table), surface it as a Suspension flag.
function flagKeyOf(r){
  const c = (r.change || "").trim();
  if (FLAG_DEFS[c]) return c;
  if ((r.suspended || "").trim()) return "Suspension";
  return null;
}
function flagCountsFor(recs){ const o = {}; for (const r of recs){ const k = flagKeyOf(r); if (k) o[k] = (o[k] || 0) + 1; } return o; }
// Plan-discrepancy flag. Deliberately NOT part of FLAG_DEFS: flagKeyOf returns a
// single change-type flag, and a terminated auth can also disagree with the
// plan — folding it in there would hide one of the two. Detail is '' when the
// auth agrees with the plan, so it doubles as the has-a-finding test.
function hasPlanDisc(r){ return !!(r.plan_disc_detail || '').trim(); }
function discPill(r){
  if (!hasPlanDisc(r)) return '';
  return ` <span class="flag-pill" style="--fc:var(--flag-disc)" title="${escapeHtml(r.plan_disc || 'Plan discrepancy')}">&#9888; DISCREPANCY</span>`;
}
function changeCell(r){
  const c = (r.change || "").trim();
  if (!c) return '<span style="color:var(--text-3)">—</span>';
  const f = FLAG_DEFS[c];
  return f ? `<span class="flag-pill" style="--fc:${f.color}">${escapeHtml(c)}</span>` : escapeHtml(c);
}

// ---- Index records by received date ----
const byDate = {};
for (const r of RECORDS) { (byDate[r.date] = byDate[r.date] || []).push(r); }

// Decision buckets actually present (for the legend).
const PRESENT_DECISIONS = DEC_ORDER.filter(s => RECORDS.some(r => decBucket(r.decision) === s));

const state = { view: 'month', cursor: todayCursor(), selectedDate: null, filterCM: '', filterDecision: null, filterFlag: null, filterDisc: false, query: '' };

function todayCursor(){ const d = new Date(); return { year: d.getFullYear(), month: d.getMonth(), day: d.getDate() }; }
function isoDate(y, m, d){ return `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`; }
function formatMonthYear(y, m){ return new Date(y, m, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' }); }
function formatLongDate(y, m, d){ return new Date(y, m, d).toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }); }
function addDaysCursor(cursor, n){ const d = new Date(cursor.year, cursor.month, cursor.day); d.setDate(d.getDate() + n); return { year: d.getFullYear(), month: d.getMonth(), day: d.getDate() }; }
function sundayOf(cursor){ const d = new Date(cursor.year, cursor.month, cursor.day); d.setDate(d.getDate() - d.getDay()); return d; }
function escapeHtml(s){ return String(s ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":"&#39;" }[c])); }
function setStat(n, label){ document.getElementById('stats-count').textContent = Number(n).toLocaleString(); document.getElementById('stats-label').textContent = label; }

function decCountsFor(recs){ const o = {}; for (const r of recs) { const b = decBucket(r.decision); o[b] = (o[b] || 0) + 1; } return o; }
function decPill(dec){
  const raw = (dec || "").trim();
  const b = decBucket(dec);
  if (!raw) return '<span style="color:var(--text-3)">—</span>';
  return `<span class="dec-pill"><span class="dec-dot" style="background:${DEC_COLOR[b]}"></span>${escapeHtml(b)}</span>`;
}

function render(){
  // An active search takes over the main area.
  if (state.query) { renderSearch(); return; }
  document.getElementById('search-view').hidden = true;
  document.querySelectorAll('.view-toggle button').forEach(btn => {
    btn.setAttribute('aria-selected', btn.dataset.view === state.view ? 'true' : 'false');
  });
  document.getElementById('month-view').hidden = state.view !== 'month';
  document.getElementById('week-view').hidden = state.view !== 'week';
  document.getElementById('day-view').hidden = state.view !== 'day';
  renderLegend();
  renderPeriodLabel();
  if (state.view === 'month') renderMonth();
  else if (state.view === 'week') renderWeek();
  else if (state.view === 'day') renderDay();
}

function renderSearch(){
  // Hide the calendar views; show the results panel.
  document.getElementById('month-view').hidden = true;
  document.getElementById('week-view').hidden = true;
  document.getElementById('day-view').hidden = true;
  document.getElementById('search-view').hidden = false;
  document.querySelectorAll('.view-toggle button').forEach(btn => btn.setAttribute('aria-selected', 'false'));

  const q = state.query.toLowerCase();
  const terms = q.split(/\s+/).filter(Boolean);
  const hit = r => {
    const hay = `${r.member || ''} ${r.client || ''} ${r.auth || ''} ${r.medicaid || ''} ${r.cm || ''}`.toLowerCase();
    return terms.every(t => hay.includes(t));
  };
  const matches = RECORDS.filter(hit);
  setStat(matches.length, matches.length === 1 ? 'match' : 'matches');
  document.getElementById('search-sub').textContent =
    `"${state.query}" · ${matches.length} of ${RECORDS.length} authorization${RECORDS.length === 1 ? '' : 's'}`;

  const container = document.getElementById('search-results-container');
  if (!matches.length) {
    container.innerHTML = `<div class="empty"><div class="empty-title">No matches</div>
      <div class="empty-hint">Nothing matched "${escapeHtml(state.query)}". Try a name, client ID, auth #, or Medicaid #.</div></div>`;
    return;
  }

  // Group results by date, newest first.
  const byDay = {};
  for (const r of matches) { (byDay[r.date] = byDay[r.date] || []).push(r); }
  const days = Object.keys(byDay).sort().reverse();

  const flat = [];
  let html = '';
  for (const day of days) {
    const [yy, mm, dd] = day.split('-').map(Number);
    const list = byDay[day].slice().sort((a, b) => (a.member || '').localeCompare(b.member || ''));
    const rows = list.map(r => {
      const idx = flat.length; flat.push(r);
      return `<tr data-idx="${idx}" tabindex="0">
        <td class="patient-id">${r.auth ? escapeHtml(r.auth) : '—'}</td>
        <td>${escapeHtml(r.member || '(no name)')}${discPill(r)}</td>
        <td class="patient-id">${r.client ? escapeHtml(r.client) : '—'}</td>
        <td class="patient-id">${r.medicaid ? escapeHtml(r.medicaid) : '—'}</td>
        <td>${escapeHtml(cmLabel(cmKeyOf(r)))}</td>
        <td class="svc">${r.services ? escapeHtml(r.services) : '—'}</td>
        <td>${decPill(r.decision)}</td>
      </tr>`;
    }).join('');
    html += `
      <div class="cm-section">
        <div class="search-date-head">${formatLongDate(yy, mm - 1, dd)} · ${list.length}</div>
        <div class="pivot-wrap">
          <table class="pivot-table">
            <thead><tr>
              <th>Auth #</th><th>Client Name</th><th>Client ID</th><th>Medicaid #</th><th>Case Manager</th><th>Services</th><th>Decision</th>
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>`;
  }
  container.innerHTML = html;
  container.querySelectorAll('tbody tr').forEach(tr => {
    const open = () => { const r = flat[Number(tr.dataset.idx)]; if (r) openAuthModal(r); };
    tr.addEventListener('click', open);
    tr.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
  });
}

function renderLegend(){
  document.getElementById('legend').innerHTML = PRESENT_DECISIONS.map(s => `
    <span class="legend-item"><span class="legend-dot" style="background: ${DEC_COLOR[s]}"></span>${s}</span>
  `).join('');
}

function renderPeriodLabel(){
  const el = document.getElementById('period-label');
  if (state.view === 'day' && state.selectedDate) {
    const [y, m, d] = state.selectedDate.split('-').map(Number);
    el.textContent = formatLongDate(y, m - 1, d);
  } else if (state.view === 'week') {
    const sun = sundayOf(state.cursor);
    const sat = new Date(sun); sat.setDate(sat.getDate() + 6);
    el.textContent = `${sun.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} – ${sat.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}`;
  } else {
    el.textContent = formatMonthYear(state.cursor.year, state.cursor.month);
  }
}

function renderMonth(){
  const { year, month } = state.cursor;
  const firstDayOfWeek = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const prevMonthLast = new Date(year, month, 0).getDate();

  const cells = [];
  for (let i = firstDayOfWeek - 1; i >= 0; i--) {
    cells.push({ year: month === 0 ? year - 1 : year, month: month === 0 ? 11 : month - 1, day: prevMonthLast - i, otherMonth: true });
  }
  for (let d = 1; d <= daysInMonth; d++) cells.push({ year, month, day: d, otherMonth: false });
  while (cells.length < 42) {
    const nextDay = cells.length - firstDayOfWeek - daysInMonth + 1;
    cells.push({ year: month === 11 ? year + 1 : year, month: month === 11 ? 0 : month + 1, day: nextDay, otherMonth: true });
  }

  const td = todayCursor();
  let total = 0;
  const grid = document.getElementById('days-grid');
  grid.innerHTML = cells.map((c, i) => {
    const key = isoDate(c.year, c.month, c.day);
    const recs = (!c.otherMonth && byDate[key]) ? byDate[key] : null;
    const isToday = c.year === td.year && c.month === td.month && c.day === td.day;
    const hasData = !!recs;
    if (hasData) total += recs.length;
    const counts = hasData ? decCountsFor(recs) : {};
    const bars = hasData
      ? DEC_ORDER.filter(s => (counts[s] || 0) > 0)
          .map(s => `<div class="status-bar" style="background:${DEC_COLOR[s]}; flex:${counts[s]}" title="${s}: ${counts[s]}"></div>`).join('')
      : '';
    const flags = hasData ? flagCountsFor(recs) : {};
    const flagMarks = FLAG_ORDER.filter(k => flags[k])
      .map(k => `<span class="day-flag" style="background:${FLAG_DEFS[k].color}" title="${k}: ${flags[k]}">${FLAG_DEFS[k].short} ${flags[k]}</span>`).join('');
    return `
      <div class="day-cell ${c.otherMonth ? 'other-month' : ''} ${isToday ? 'today' : ''} ${hasData ? 'has-data' : ''}" data-date="${key}" style="--i:${i}">
        <div class="day-num">${c.day}</div>
        ${flagMarks ? `<div class="day-flags">${flagMarks}</div>` : ''}
        ${hasData ? `
          <div class="visits-block">
            <div class="visits-count">${recs.length}</div>
            <div class="visits-label">Auth${recs.length === 1 ? '' : 's'}</div>
          </div>` : ''}
        ${bars ? `<div class="status-bars">${bars}</div>` : ''}
      </div>`;
  }).join('');

  setStat(total, 'auths this month');
  grid.querySelectorAll('.day-cell.has-data').forEach(cell => {
    cell.addEventListener('click', () => drillToDay(cell.dataset.date));
  });
}

function renderWeek(){
  const sun = sundayOf(state.cursor);
  const grid = document.getElementById('week-grid');
  let total = 0;
  grid.innerHTML = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(sun); d.setDate(d.getDate() + i);
    const key = isoDate(d.getFullYear(), d.getMonth(), d.getDate());
    const recs = byDate[key] || null;
    if (recs) total += recs.length;
    const name = d.toLocaleDateString('en-US', { weekday: 'short' });
    const counts = recs ? decCountsFor(recs) : {};
    const rows = recs
      ? DEC_ORDER.filter(s => (counts[s] || 0) > 0).map(s => `
          <div class="week-stat-row">
            <span class="week-stat-label"><span class="legend-dot" style="background:${DEC_COLOR[s]}"></span>${s}</span>
            <span class="week-stat-count">${counts[s]}</span>
          </div>`).join('')
      : '';
    const members = recs && recs.length
      ? recs.slice(0, 6).map(r => `<div>${escapeHtml(r.member || '(no name)')}</div>`).join('') +
        (recs.length > 6 ? `<div class="more">+${recs.length - 6} more</div>` : '')
      : '<div class="week-empty">No auths</div>';
    const cmCount = recs ? new Set(recs.map(cmKeyOf)).size : 0;
    const wflags = recs ? flagCountsFor(recs) : {};
    const wflagMarks = FLAG_ORDER.filter(k => wflags[k])
      .map(k => `<span class="day-flag" style="background:${FLAG_DEFS[k].color}" title="${k}: ${wflags[k]}">${FLAG_DEFS[k].short} ${wflags[k]}</span>`).join('');
    return `
      <div class="week-day ${recs ? 'has-data' : ''}" data-date="${key}" style="--i:${i}">
        <div class="week-day-header">
          <div class="week-day-name">${name}</div>
          <div class="week-day-num">${d.getDate()}</div>
          ${recs ? `<div class="week-day-total">${recs.length} auth${recs.length === 1 ? '' : 's'} · ${cmCount} CM${cmCount === 1 ? '' : 's'}</div>` : ''}
          ${wflagMarks ? `<div class="week-flags">${wflagMarks}</div>` : ''}
        </div>
        ${rows ? `<div class="week-stats">${rows}</div>` : ''}
        <div class="week-patients">${members}</div>
      </div>`;
  }).join('');

  setStat(total, 'auths this week');
  grid.querySelectorAll('.week-day.has-data').forEach(el => {
    el.addEventListener('click', () => drillToDay(el.dataset.date));
  });
}

function renderDay(){
  const dateStr = state.selectedDate || isoDate(state.cursor.year, state.cursor.month, state.cursor.day);
  const all = byDate[dateStr] || [];
  const [y, m, d] = dateStr.split('-').map(Number);
  document.getElementById('day-title').textContent = formatLongDate(y, m - 1, d);

  if (!all.length) {
    document.getElementById('day-subtitle').textContent = 'No authorizations received';
    document.getElementById('day-badges').innerHTML = '';
    document.getElementById('filter-bar').hidden = true;
    document.getElementById('day-table-container').innerHTML = `
      <div class="empty">
        <div class="empty-title">Nothing to show here</div>
        <div class="empty-hint">No UHC authorizations were received on this day.</div>
      </div>`;
    setStat(0, 'auths on this day');
    return;
  }

  // Case-manager dropdown (Unmatched listed last).
  const cmKeys = Array.from(new Set(all.map(cmKeyOf)))
    .sort((a, b) => a === UN ? 1 : b === UN ? -1 : a.localeCompare(b));
  if (!cmKeys.includes(state.filterCM)) state.filterCM = '';
  const selectEl = document.getElementById('filter-practice');
  selectEl.innerHTML = `<option value="">All case managers (${cmKeys.length})</option>` +
    cmKeys.map(k => `<option value="${escapeHtml(k)}" ${state.filterCM === k ? 'selected' : ''}>${escapeHtml(cmLabel(k))}</option>`).join('');
  document.getElementById('filter-bar').hidden = false;

  // Decision filter badges.
  const counts = decCountsFor(all);
  const badges = DEC_ORDER.filter(s => (counts[s] || 0) > 0).map(s => {
    const active = state.filterDecision === s, dim = state.filterDecision && !active;
    return `<span class="badge ${active ? 'active' : ''} ${dim ? 'dimmed' : ''}" data-dec="${escapeHtml(s)}" role="button" tabindex="0" aria-pressed="${active}">
      <span class="badge-dot" style="background:${DEC_COLOR[s]}"></span>${s} <strong>${counts[s]}</strong></span>`;
  }).join('');
  // Urgency-flag filter chips (Records Request / Termination / Suspension) — time-critical.
  const fcounts = flagCountsFor(all);
  const flagBadges = FLAG_ORDER.filter(k => fcounts[k]).map(k => {
    const active = state.filterFlag === k, dim = state.filterFlag && !active;
    return `<span class="badge flag-badge ${active ? 'active' : ''} ${dim ? 'dimmed' : ''}" style="--fc:${FLAG_DEFS[k].color}" data-flag="${escapeHtml(k)}" role="button" tabindex="0" aria-pressed="${active}">
      <span class="badge-dot" style="background:${FLAG_DEFS[k].color}"></span>${k} <strong>${fcounts[k]}</strong></span>`;
  }).join('');
  // Plan-discrepancy chip — its own filter, independent of the urgency flags.
  const dcount = all.filter(hasPlanDisc).length;
  const discBadge = dcount ? (() => {
    const active = state.filterDisc;
    return `<span class="badge flag-badge ${active ? 'active' : ''}" style="--fc:var(--flag-disc)" data-disc="1" role="button" tabindex="0" aria-pressed="${active}">
      <span class="badge-dot" style="background:var(--flag-disc)"></span>Plan Discrepancy <strong>${dcount}</strong></span>`;
  })() : '';
  const clearBtn = (state.filterDecision || state.filterFlag || state.filterDisc) ? `
    <button class="filter-clear" id="clear-filter" aria-label="Clear filter">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      Clear
    </button>` : '';
  document.getElementById('day-badges').innerHTML = badges + flagBadges + discBadge + clearBtn;

  // Apply filters.
  let recs = all.slice();
  if (state.filterCM) recs = recs.filter(r => cmKeyOf(r) === state.filterCM);
  if (state.filterDecision) recs = recs.filter(r => decBucket(r.decision) === state.filterDecision);
  if (state.filterFlag) recs = recs.filter(r => flagKeyOf(r) === state.filterFlag);
  if (state.filterDisc) recs = recs.filter(hasPlanDisc);

  const cmsAll = new Set(all.map(cmKeyOf)).size;
  document.getElementById('day-subtitle').textContent = (state.filterCM || state.filterDecision || state.filterFlag || state.filterDisc)
    ? `Showing ${recs.length} of ${all.length} auth${all.length === 1 ? '' : 's'}`
    : `${all.length} auth${all.length === 1 ? '' : 's'} · ${cmsAll} case manager${cmsAll === 1 ? '' : 's'}`;

  // Group filtered records by CM and build per-CM tables.
  const groups = {};
  for (const r of recs) { const k = cmKeyOf(r); (groups[k] = groups[k] || []).push(r); }
  const names = Object.keys(groups).sort((a, b) => a === UN ? 1 : b === UN ? -1 : a.localeCompare(b));

  const flat = [];
  let html = '';
  if (!recs.length) {
    html = `<div class="empty"><div class="empty-title">No matches</div><div class="empty-hint">Try a different filter or clear it.</div></div>`;
  } else {
    for (const k of names) {
      const list = groups[k].slice().sort((a, b) => (a.member || '').localeCompare(b.member || ''));
      const isUn = k === UN;
      const rows = list.map(r => {
        const idx = flat.length; flat.push(r);
        return `<tr data-idx="${idx}" tabindex="0">
          <td class="patient-id">${r.auth ? escapeHtml(r.auth) : '—'}</td>
          <td>${escapeHtml(r.member || '(no name)')}${discPill(r)}</td>
          <td class="patient-id">${r.client ? escapeHtml(r.client) : '—'}</td>
          <td class="patient-id">${r.medicaid ? escapeHtml(r.medicaid) : '—'}</td>
          <td>${changeCell(r)}</td>
          <td class="svc">${r.services ? escapeHtml(r.services) : '—'}</td>
          <td>${decPill(r.decision)}</td>
        </tr>`;
      }).join('');
      html += `
        <div class="cm-section">
          <div class="cm-section-head ${isUn ? 'unmatched' : ''}">
            <span class="cm-section-name">${escapeHtml(cmLabel(k))}</span>
            <span class="cm-section-count">${list.length}</span>
          </div>
          <div class="pivot-wrap">
            <table class="pivot-table">
              <thead><tr>
                <th>Auth #</th><th>Client Name</th><th>Client ID</th><th>Medicaid #</th><th>Change Type</th><th>Services</th><th>Decision</th>
              </tr></thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </div>`;
    }
  }
  document.getElementById('day-table-container').innerHTML = html;

  document.querySelectorAll('#day-table-container tbody tr').forEach(tr => {
    const open = () => { const r = flat[Number(tr.dataset.idx)]; if (r) openAuthModal(r); };
    tr.addEventListener('click', open);
    tr.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
  });
  document.querySelectorAll('#day-badges .badge').forEach(el => {
    const toggle = () => {
      if (el.dataset.disc !== undefined) {
        state.filterDisc = !state.filterDisc;
      } else if (el.dataset.flag !== undefined && el.dataset.flag !== '') {
        const k = el.dataset.flag; state.filterFlag = state.filterFlag === k ? null : k;
      } else {
        const s = el.dataset.dec; state.filterDecision = state.filterDecision === s ? null : s;
      }
      renderDay();
    };
    el.addEventListener('click', toggle);
    el.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); } });
  });
  const clearEl = document.getElementById('clear-filter');
  if (clearEl) clearEl.addEventListener('click', () => { state.filterDecision = null; state.filterFlag = null; state.filterDisc = false; renderDay(); });

  setStat(all.length, 'auths on this day');
}

function openAuthModal(r){
  document.getElementById('modal-eyebrow').textContent = 'Authorization Detail';
  document.getElementById('modal-title').textContent = r.member || '(no name)';
  const subParts = [];
  if (r.auth) subParts.push(`Auth #${r.auth}`);
  subParts.push(cmLabel(cmKeyOf(r)));
  document.getElementById('modal-subtitle').textContent = subParts.join(' · ');

  const fields = [
    ['Authorization #', r.auth],
    ['Change Type', r.change],
    ['Case Manager', r.cm],
    ['Lookup Status', r.status],
    ['Client ID', r.client],
    ['Member DOB', r.dob],
    ['Health Plan ID', r.hpid],
    ['Medicaid ID', r.medicaid],
    ['Auth Period', (r.start || r.end) ? `${r.start || '?'} – ${r.end || '?'}` : ''],
    ['Address', r.address, true],
    ['Service Suspended', r.suspended ? (r.suspended + (r.suspended_start ? ` (since ${r.suspended_start})` : '')) : '', true],
    ['Services', r.services, true],
    // Only rendered when there IS a finding — detail is '' when the auth agrees
    // with the plan, and the empty-value filter below drops the row entirely.
    ['Plan Discrepancy',
     r.plan_disc_detail ? ((r.plan_disc ? r.plan_disc + '\n' : '') + r.plan_disc_detail) : '',
     true],
    ['Notes', r.notes, true],
    ['Journal Note', r.journal, true],
  ];
  const grid = fields.filter(f => f[1] && String(f[1]).trim())
    .map(f => `<div class="modal-field ${f[2] ? 'full' : ''}"><span class="label">${escapeHtml(f[0])}</span><span class="value">${escapeHtml(f[1])}</span></div>`)
    .join('');
  const b = decBucket(r.decision);
  const fk = flagKeyOf(r);
  const flagPill = fk
    ? `<span class="status-pill"><span class="status-dot" style="background:${FLAG_DEFS[fk].color}"></span>${escapeHtml(fk)}</span>`
    : '';
  const clientMeta = (r.client && r.client.trim())
    ? `<span class="modal-event-meta">Client ${escapeHtml(r.client)}</span>`
    : `<span class="modal-event-meta" style="color:var(--dec-denied)">No Client ID</span>`;
  let html = `
    <div class="modal-event-head">
      ${flagPill}
      <span class="status-pill"><span class="status-dot" style="background:${DEC_COLOR[b]}"></span>${escapeHtml(r.decision || '—')}</span>
      ${clientMeta}
    </div>
    <div class="modal-grid">${grid}</div>`;
  if (r.pdf) html += `<a class="pdf-link" href="${escapeHtml(r.pdf)}" target="_blank" rel="noopener">Open in SharePoint ↗</a>`;
  document.getElementById('modal-body').innerHTML = html;
  document.getElementById('modal-backdrop').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeModal(){
  document.getElementById('modal-backdrop').classList.remove('open');
  document.body.style.overflow = '';
}
document.getElementById('modal-close').addEventListener('click', closeModal);
document.getElementById('modal-backdrop').addEventListener('click', e => { if (e.target.id === 'modal-backdrop') closeModal(); });

function drillToDay(dateStr){
  state.selectedDate = dateStr;
  state.filterDecision = null;
  state.filterCM = '';
  const [y, m, d] = dateStr.split('-').map(Number);
  state.cursor = { year: y, month: m - 1, day: d };
  state.view = 'day';
  render();
}

function exitSearch(){
  if (!state.query) return;
  state.query = '';
  const si = document.getElementById('global-search');
  if (si) si.value = '';
  const sx = document.getElementById('search-clear-x');
  if (sx) sx.style.display = 'none';
}

document.querySelectorAll('.view-toggle button').forEach(btn => {
  btn.addEventListener('click', () => {
    exitSearch();
    state.view = btn.dataset.view;
    if (state.view === 'day' && !state.selectedDate) {
      state.selectedDate = isoDate(state.cursor.year, state.cursor.month, state.cursor.day);
    }
    render();
  });
});

document.getElementById('prev-btn').addEventListener('click', () => {
  exitSearch();
  if (state.view === 'month') {
    state.cursor.month--;
    if (state.cursor.month < 0) { state.cursor.month = 11; state.cursor.year--; }
  } else if (state.view === 'week') {
    state.cursor = addDaysCursor(state.cursor, -7);
  } else {
    state.cursor = addDaysCursor(state.cursor, -1);
    state.selectedDate = isoDate(state.cursor.year, state.cursor.month, state.cursor.day);
    state.filterDecision = null; state.filterCM = '';
  }
  render();
});

document.getElementById('next-btn').addEventListener('click', () => {
  exitSearch();
  if (state.view === 'month') {
    state.cursor.month++;
    if (state.cursor.month > 11) { state.cursor.month = 0; state.cursor.year++; }
  } else if (state.view === 'week') {
    state.cursor = addDaysCursor(state.cursor, 7);
  } else {
    state.cursor = addDaysCursor(state.cursor, 1);
    state.selectedDate = isoDate(state.cursor.year, state.cursor.month, state.cursor.day);
    state.filterDecision = null; state.filterCM = '';
  }
  render();
});

document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  if (document.getElementById('modal-backdrop').classList.contains('open')) { closeModal(); return; }
  if (state.view !== 'month') { state.view = 'month'; state.selectedDate = null; render(); }
});

document.getElementById('filter-practice').addEventListener('change', e => {
  state.filterCM = e.target.value;
  renderDay();
});

const searchInput = document.getElementById('global-search');
const searchClearX = document.getElementById('search-clear-x');
function applySearch(){
  state.query = searchInput.value.trim();
  searchClearX.style.display = state.query ? 'block' : 'none';
  render();
}
searchInput.addEventListener('input', applySearch);
searchClearX.addEventListener('click', () => {
  searchInput.value = '';
  state.query = '';
  searchClearX.style.display = 'none';
  searchInput.focus();
  render();
});
searchInput.addEventListener('keydown', e => {
  if (e.key === 'Escape' && state.query) { e.stopPropagation(); searchInput.value = ''; state.query = ''; searchClearX.style.display = 'none'; render(); }
});

// ===== Init =====
document.getElementById('generated-at').textContent = `Generated · ${GENERATED}`;
const allDates = Object.keys(byDate).sort();
if (allDates.length) {
  const last = allDates[allDates.length - 1];
  const [y, m, d] = last.split('-').map(Number);
  state.cursor = { year: y, month: m - 1, day: d };
}
render();
</script>
</body>
</html>"""


def _created_local_date(it: dict) -> str:
    """The list item's createdDateTime as a local (America/New_York) YYYY-MM-DD,
    so 'date received' lands on the right calendar day. Falls back to the UTC
    date if tz data is unavailable."""
    raw = (it.get("createdDateTime") or "").strip()
    if not raw:
        return ""
    try:
        dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError:
        return raw[:10]
    try:
        from zoneinfo import ZoneInfo
        dt = dt.astimezone(ZoneInfo("America/New_York"))
    except Exception:
        pass
    return dt.strftime("%Y-%m-%d")


def _calendar_date(it: dict) -> str:
    """Day this auth sits on (YYYY-MM-DD): its payload Review Date when present,
    else the row's creation date — so back-filled / late-ingested rows land on
    their real day, not the SharePoint ingest day."""
    review = _norm_dob((it.get("fields") or {}).get("ReviewDate"))
    return review or _created_local_date(it)


def _cal_record(it: dict) -> dict:
    f = it.get("fields", {}) or {}
    return {
        "date": _calendar_date(it),
        "auth": (f.get("Title") or "").strip(),
        "member": (f.get("MemberName") or "").strip(),
        "dob": _xls_fmt_date(f.get("MemberDOB")),
        "decision": (f.get("OverallDecision") or "").strip(),
        "change": (f.get("ChangeType") or "").strip(),
        "cm": (f.get(_XLS_CM_FIELD) or "").strip(),
        "client": (f.get("ClientID") or "").strip(),
        "status": (f.get(LOOKUP_STATUS_FIELD) or "").strip(),
        "hpid": (f.get("HealthPlanID") or "").strip(),
        "medicaid": (f.get("MedicaidID") or "").strip(),
        "address": _xls_address(f),
        "start": _xls_fmt_date(f.get("AuthPeriodStart")),
        "end": _xls_fmt_date(f.get("AuthPeriodEnd")),
        "services": (f.get("Services") or "").strip(),
        "notes": (f.get("Notes") or "").strip(),
        "journal": (f.get("JournalNote") or "").strip(),
        "plan_disc": _plain_text(f.get(PLAN_DISC_FIELD)),
        "plan_disc_detail": _plain_text(f.get(PLAN_DISC_DETAIL_FIELD)),
        "suspended": _plain_text(f.get(SUSP_SVC_FIELD)),
        "suspended_start": _xls_fmt_date(f.get(SUSP_START_FIELD)),
        "pdf": _dispform_url(it["id"]) if f.get("Attachments") else "",
    }


def build_calendar_html(items: list[dict], out_path: Path | None = None) -> Path:
    """Write the interactive calendar HTML from a fetched items list. Uses the
    same dedupe as the workbook so each authorization appears once, placed by
    _calendar_date (Review Date, else creation date)."""
    if out_path is None:
        out_path = SCRIPT_DIR / _HTML_CAL_NAME
    clean = _xls_dedupe(items)
    records = [_cal_record(it) for it in clean]
    records = [r for r in records if r["date"]]
    records.sort(key=lambda r: (r["date"], r["member"]))
    data_json = json.dumps(records, ensure_ascii=False).replace("</", "<\\/")
    html = (
        _CAL_TEMPLATE
        .replace("__DATA__", data_json)
        .replace("__GENERATED__", datetime.now().strftime("%b %d, %Y %I:%M %p"))
        .replace("__TOTAL__", str(len(records)))
    )
    out_path.write_text(html, encoding="utf-8")
    log.info("[calendar] wrote %s (%d authorizations across %d day(s))",
             out_path, len(records), len({r["date"] for r in records}))
    return out_path


# ── Azure Static Web App deploy (calendar) ─────────────────────────
# The deploy token comes from the Static Web App's "Manage deployment token".
# It lives in a .swa_deploy_token file next to this script. Site:
# https://purple-tree-023254d10.7.azurestaticapps.net
_SWA_TOKEN_FILE = SCRIPT_DIR / ".swa_deploy_token"
# The script may run from a folder other than the original one (e.g. the
# AiHub copy). Fall back to the original pipeline folder under Report
# Subscriptions for the token and the companion-page stage dir.
_ORIG_PIPELINE_DIR = (REPORT_SUBS_DIR / "Special Programs" / "Authorizations"
                      / "United Health Care")
if not _SWA_TOKEN_FILE.exists() and (_ORIG_PIPELINE_DIR / ".swa_deploy_token").exists():
    _SWA_TOKEN_FILE = _ORIG_PIPELINE_DIR / ".swa_deploy_token"

# Persistent stage dir for pages published to this site by OTHER pipelines.
# A builder writes its page here; deploy_calendar copies whatever is present
# into the upload. Mirrors Power Pages/swa/ps (index.html + daily.html).
_SWA_STAGE_DIR = SCRIPT_DIR / "swa_calendar"
if not _SWA_STAGE_DIR.exists() and (_ORIG_PIPELINE_DIR / "swa_calendar").exists():
    _SWA_STAGE_DIR = _ORIG_PIPELINE_DIR / "swa_calendar"
# Filename -> the URL it answers on. tufts-medical.html is reachable as
# /tufts-medical via the rewrite in _SWA_CONFIG_JSON below.
_SWA_COMPANION_PAGES = ("tufts-medical.html",)

# Forces Microsoft Entra (Azure AD) sign-in on every route. Requires, on the
# purple-tree Static Web App: env vars AAD_CLIENT_ID + AAD_CLIENT_SECRET (set), and
# the "UHC Authorization Dashboard" app registration (client 9dc5dab6-ee07-40d9-872b-296ce290dd29)
# with redirect URI
#   https://purple-tree-023254d10.7.azurestaticapps.net/.auth/login/aad/callback
# openIdIssuer tenant = CBES directory (tenant) ID.
_SWA_CONFIG_JSON = """{
  "$schema": "https://json.schemastore.org/staticwebapp.config.json",
  "auth": {
    "identityProviders": {
      "azureActiveDirectory": {
        "registration": {
          "openIdIssuer": "https://login.microsoftonline.com/ea4caad5-4b33-461b-8faa-9c22a94803c0/v2.0",
          "clientIdSettingName": "AAD_CLIENT_ID",
          "clientSecretSettingName": "AAD_CLIENT_SECRET"
        }
      }
    }
  },
  "routes": [
    { "route": "/login", "rewrite": "/.auth/login/aad" },
    { "route": "/logout", "redirect": "/.auth/logout" },
    { "route": "/.auth/login/github", "statusCode": 404 },
    { "route": "/.auth/login/twitter", "statusCode": 404 },
    { "route": "/tufts-medical", "rewrite": "/tufts-medical.html" },
    { "route": "/*", "allowedRoles": ["authenticated"] }
  ],
  "responseOverrides": {
    "401": { "redirect": "/.auth/login/aad", "statusCode": 302 }
  },
  "navigationFallback": {
    "rewrite": "/index.html",
    "exclude": ["/.auth/*", "*.{css,js,png,jpg,jpeg,svg,ico,webp,woff,woff2}"]
  },
  "globalHeaders": {
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "SAMEORIGIN",
    "Cache-Control": "no-store"
  }
}
"""


def _read_token_robust(token_file: Path) -> str:
    """Read a deploy token that may be an un-hydrated OneDrive placeholder.

    A rarely-touched dotfile in OneDrive can be online-only on the scheduler box,
    so a plain read fails with Errno 13/22. Retry, then force-hydrate via
    robocopy /ZB (reads locked/placeholder files) and read a local snapshot.
    Returns "" if it can't be read.
    """
    import subprocess
    import tempfile
    import time

    if not token_file.exists():
        return ""
    for delay in (0, 2, 4):
        if delay:
            time.sleep(delay)
        try:
            t = token_file.read_text(encoding="utf-8").strip()
            if t:
                return t
        except OSError:
            pass
    try:
        tmp = Path(tempfile.gettempdir()) / "uhc_swa_token"
        tmp.mkdir(parents=True, exist_ok=True)
        proc = subprocess.run(
            ["robocopy", str(token_file.parent), str(tmp), token_file.name,
             "/ZB", "/R:2", "/W:3", "/NFL", "/NDL", "/NJH", "/NJS", "/NP"],
            capture_output=True, text=True,
        )
        snap = tmp / token_file.name
        if proc.returncode < 8 and snap.exists():
            return snap.read_text(encoding="utf-8").strip()
    except Exception:
        pass
    return ""


def deploy_calendar(html_path: Path) -> None:
    """Push the calendar HTML to its Azure Static Web App. Called every run.

    Best-effort: any failure is logged and swallowed so it can never fail the
    pipeline. Reads the deploy token from .swa_deploy_token next to this script.
    """
    import shutil
    import subprocess

    if not html_path or not Path(html_path).exists():
        log.warning("[deploy] calendar HTML missing — nothing to deploy")
        return
    token = _read_token_robust(_SWA_TOKEN_FILE)
    if not token:
        log.warning("[deploy] no/empty %s — skipping Azure deploy",
                    _SWA_TOKEN_FILE)
        return

    # Stage a deploy folder with the calendar as the site root (index.html).
    #
    # Staged to LOCAL disk, not next to this script. SCRIPT_DIR lives under
    # OneDrive, and StaticSitesClient.exe intermittently dies on a OneDrive
    # path — "Deployment failed with exit code 1" with no real reason, while
    # the identical payload deploys fine from a plain local folder. Same
    # failure mode as the exec dashboard, same fix.
    try:
        swa_dir = Path(tempfile.gettempdir()) / "uhc_swa_calendar"
        swa_dir.mkdir(parents=True, exist_ok=True)
        for name in ("index.html", _HTML_CAL_NAME):
            shutil.copyfile(html_path, swa_dir / name)
        # Companion pages ride along. A `swa deploy` publishes the staged
        # folder as the WHOLE site, so anything not staged here is deleted
        # from the site -- which is why this used to be UHC-only and why a
        # second page couldn't just be pushed separately. Same shape as
        # Power Pages/swa/ps, where daily.html sits beside index.html and the
        # folder goes up as a unit.
        #
        # The page is read from the persistent stage dir, NOT from wherever it
        # was built: the Tufts pipeline drops it there when it runs, and the
        # copy survives so a UHC deploy never blanks /tufts-medical just
        # because Tufts hasn't run recently.
        for name in _SWA_COMPANION_PAGES:
            src = _SWA_STAGE_DIR / name
            if src.exists():
                shutil.copyfile(src, swa_dir / name)
                log.info("[deploy] including companion page %s", name)
            else:
                log.info("[deploy] companion page %s not staged yet — "
                         "site will not carry it", name)
        # Force Entra sign-in: ship the auth config alongside the HTML.
        (swa_dir / "staticwebapp.config.json").write_text(
            _SWA_CONFIG_JSON, encoding="utf-8")
    except Exception as e:
        log.warning("[deploy] could not stage deploy folder: %s", e)
        return

    swa_exe = os.path.expandvars(r"%APPDATA%\npm\swa.cmd")
    if not os.path.exists(swa_exe):
        log.warning("[deploy] swa CLI not found at %s — run: "
                    "npm install -g @azure/static-web-apps-cli", swa_exe)
        return
    try:
        # The swa CLI scans its working directory for a config file and the
        # deploy binary refuses to run from inside the artifact folder. A
        # scheduled task's cwd can be anything (System32 fails with EPERM),
        # so always run from a dedicated empty local folder.
        swa_cwd = Path(tempfile.gettempdir()) / "uhc_swa_cwd"
        swa_cwd.mkdir(parents=True, exist_ok=True)
        r = subprocess.run(
            [swa_exe, "deploy", str(swa_dir),
             "--deployment-token", token, "--env", "production"],
            capture_output=True, text=True, cwd=str(swa_cwd),
        )
        if r.returncode == 0:
            log.info("[deploy] OK - calendar deployed to Azure Static Web App")
        else:
            log.warning("[deploy] swa deploy failed (exit %s)\n"
                        "stdout: %s\nstderr: %s", r.returncode,
                        (r.stdout or "")[-1500:], (r.stderr or "")[-1500:])
    except Exception as e:
        log.warning("[deploy] swa deploy errored: %s", e)


def main() -> int:
    p = argparse.ArgumentParser(description="UHC Authorizations pipeline")
    p.add_argument("--limit", type=int, default=None,
                   help="Process at most N canonical items (testing).")
    p.add_argument("--since", type=str, default=None,
                   help="Only items with ReviewDate >= YYYY-MM-DD "
                        "(default: 30 days ago).")
    p.add_argument("--all", action="store_true",
                   help="Full scan: pull the WHOLE list instead of only "
                        "unprocessed items. Use for backfills/audits.")
    p.add_argument("--reprocess", action="store_true",
                   help="Re-enrich items even if Lookup Status is already set.")
    p.add_argument("--delete-duplicates", action="store_true",
                   help="DELETE duplicate items instead of flagging them.")
    p.add_argument("--no-excel", action="store_true",
                   help="Skip rebuilding the uhc_authorizations.xlsx report.")
    p.add_argument("--no-calendar", action="store_true",
                   help="Skip rebuilding the calendar HTML "
                        "(uhc_authorizations_calendar.html).")
    p.add_argument("--no-deploy", action="store_true",
                   help="Skip pushing the calendar HTML to its Azure Static "
                        "Web App.")
    p.add_argument("--no-suspensions", action="store_true",
                   help="Skip the active service-suspension sync "
                        "(Service Suspended / Suspension Start Date columns).")
    p.add_argument("--no-plan-check", action="store_true",
                   help="Skip the auth-vs-active-service-plan discrepancy check "
                        "(Plan Discrepancy / Plan Discrepancy Detail columns).")
    p.add_argument("--plan-check-dry-run", action="store_true",
                   help="Run the auth-vs-service-plan check and LOG what it "
                        "would write, without writing anything to SharePoint.")
    p.add_argument("--no-rematch", action="store_true",
                   help="Skip re-running the consumer lookup on previously "
                        "unmatched items (no Client ID / Case Manager).")
    p.add_argument("--force", action="store_true",
                   help="Ignore the list-modified gate and fetch even if the "
                        "list hasn't changed since the last run.")
    args = p.parse_args()

    # Stamp the run START time up front. On success this becomes the new
    # last_run marker, so items created WHILE this run executes aren't skipped
    # next time. UTC to match Graph's createdDateTime / lastModifiedDateTime.
    run_started = _utc_now_iso()

    log.info("=== UHC pipeline starting (mode=%s since=%s limit=%s "
             "reprocess=%s delete_dupes=%s rematch=%s) ===",
             "full" if (args.all or args.reprocess) else "incremental",
             args.since, args.limit, args.reprocess, args.delete_duplicates,
             not args.no_rematch)
    log.info("Report Subscriptions root: %s", REPORT_SUBS_DIR)
    if not CREDS_FILE.exists():
        log.error("Credentials not found at %s — check that this OneDrive "
                  "folder is fully synced, or set REPORT_SUBSCRIPTIONS_DIR.",
                  CREDS_FILE)
        return 1
    if not CONSUMERS_CSV.exists():
        log.warning("Consumers CSV not found at %s — consumer lookups will "
                    "fail until it syncs.", CONSUMERS_CSV)

    # reprocess implies a full scan — you can't re-enrich rows you didn't fetch.
    full_scan = args.all or args.reprocess
    last_run = load_last_run()
    log.info("Previous run: %s", last_run or "(none recorded)")

    try:
        creds = load_azure_creds()
        app = _msal_app(creds)
        g = graph_session(get_token(app, "https://graph.microsoft.com/.default"))
        sp = sp_session(get_token(app, f"https://{SP_HOSTNAME}/.default"))
        site_id = resolve_site_id(g)

        # Gate: a single metadata read on the list (no items, no payloads, no
        # PII). If the list hasn't changed since the last run, there are no new
        # creates to enrich — we skip the new-item fetch but still re-match old
        # misses and rebuild the reports below.
        list_changed = True
        if not full_scan and not args.force:
            list_mod = list_last_modified(g, site_id)
            log.info("List last modified: %s", list_mod or "(unknown)")
            # Compare at second precision: Graph timestamps may carry fractional
            # seconds ('...22.123Z'), which would sort before a whole-second
            # marker and cause a false "unchanged". Trim to whole seconds first.
            def _sec(ts: str) -> str:
                return ts.split(".")[0].rstrip("Z")
            if list_mod and last_run and _sec(list_mod) < _sec(last_run):
                list_changed = False
                log.info("[gate] list unchanged since last run — no new creates "
                         "to enrich (use --force to fetch anyway).")

        if full_scan:
            items = list_items(g, site_id)
            log.info("Full scan: pulled %d total list item(s)", len(items))
        elif list_changed:
            # Incremental: only items the flow created but hasn't had enriched
            # yet (LookupStatus still blank). The flow never sets LookupStatus,
            # so this set IS "new creates since the last successful run" — the
            # metadata gate above already confirmed the list changed. Cost
            # scales with NEW items, not total list size.
            #
            # NOTE: we filter on fields/LookupStatus only. Graph/SharePoint does
            # not reliably support combining a system property (createdDateTime)
            # with a fields/ property in one $filter, so the "since last run"
            # cut is enforced by the gate + last_run marker, not the OData query.
            items = list_items(g, site_id, filter_="fields/LookupStatus eq null")
            log.info("Incremental: pulled %d new unprocessed item(s) "
                     "(LookupStatus blank, since %s)",
                     len(items), last_run or "list start")
        else:
            items = []
    except Exception as e:
        log.exception("setup/fetch failed: %s", e)
        return 1

    # --since only further narrows a full scan (by ReviewDate). Incremental
    # mode already fetches exactly the unprocessed rows, so it's ignored there.
    if full_scan and args.since:
        cutoff = args.since.strip()
        kept = []
        for it in items:
            rd = (it.get("fields") or {}).get("ReviewDate") or ""
            if not rd or rd[:10] >= cutoff:
                kept.append(it)
        log.info("[filter] ReviewDate >= %s (or unset): %d of %d kept",
                 cutoff, len(kept), len(items))
        items = kept

    run(g, sp, site_id, items, delete_duplicates=args.delete_duplicates,
        reprocess=args.reprocess, limit=args.limit,
        guard_existing=not full_scan)

    # Pull the WHOLE list once. This single fetch drives BOTH the whole-map
    # re-match and the reports, so every authorization (keyed in the calendar by
    # its creation date) is present and as fully matched as the data allows.
    try:
        all_items = list_items(g, site_id)
        log.info("[report] pulled %d total item(s) for whole-map matching + reports",
                 len(all_items))
    except Exception as e:
        log.exception("[report] full pull failed: %s", e)
        all_items = None

    # Re-match EVERY item still missing a Client ID — not just the 'Not Found'
    # rows but anything never enriched — so the whole map gets populated. The
    # cascade's Medicaid + corroborated-Address signals resolve many stale
    # misses. Updates are written to SharePoint AND onto the in-memory items, so
    # the reports below reflect them without another fetch. A full scan already
    # matches everything via run(), so skip the extra pass there.
    if all_items is not None and not args.no_rematch and not full_scan:
        try:
            cols = existing_columns(g, site_id)
            missing = sum(1 for it in all_items
                          if not ((it.get("fields") or {}).get("ClientID") or "").strip()
                          and not _xls_is_duplicate(it))
            log.info("[rematch] %d item(s) missing a Client ID to retry", missing)
            rematch_unmatched(g, site_id, all_items, cols, update_in_place=True)
        except Exception as e:
            log.exception("[rematch] pass failed: %s", e)

    # Bring every row's Change Type onto the current wording. Runs every mode,
    # unconditionally — a display rename is not something to remember to go run
    # a one-off backfill for; the pipeline just repairs the list every time.
    if all_items is not None:
        try:
            normalize_change_types(g, site_id, all_items, update_in_place=True)
        except Exception as e:
            log.exception("[changetype] normalization pass failed: %s", e)

    # Refresh the active-service-suspension flags on every matched row from the
    # Misc/Service Suspensions export. Runs every mode (suspensions change daily,
    # independent of the auths) and writes only rows whose value changed.
    if all_items is not None and not args.no_suspensions:
        try:
            sync_suspensions(g, site_id, all_items, update_in_place=True)
        except Exception as e:
            log.exception("[suspend] sync pass failed: %s", e)

    # Compare each auth against the consumer's CURRENT ACTIVE service plan
    # (Misc/Service Plans export) and flag hours/week and days/week
    # discrepancies — nothing else. Runs every mode for the same reason as the
    # suspension sync — the plan changes daily, independent of the auths.
    if all_items is not None and not args.no_plan_check:
        try:
            sync_plan_discrepancies(g, site_id, all_items, update_in_place=True,
                                    dry_run=args.plan_check_dry_run)
        except Exception as e:
            log.exception("[plan] auth-vs-plan check failed: %s", e)

    # Rebuild the reports over that same whole list (now freshly matched).
    cal_path = None
    if all_items is not None and not (args.no_excel and args.no_calendar):
        try:
            if not args.no_excel:
                build_workbook(all_items)
            if not args.no_calendar:
                cal_path = build_calendar_html(all_items)
        except Exception as e:
            log.exception("[report] build failed: %s", e)

    # Push the calendar to its Azure Static Web App every run (best-effort;
    # deploy_calendar never raises, so it can't fail the pipeline).
    if not args.no_deploy and not args.no_calendar:
        deploy_calendar(cal_path or (SCRIPT_DIR / _HTML_CAL_NAME))

    # Mark this run complete so the next incremental run only looks at what
    # arrives afterward. Use the start time captured before any fetch.
    save_last_run(run_started)

    log.info("=== UHC pipeline done ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
