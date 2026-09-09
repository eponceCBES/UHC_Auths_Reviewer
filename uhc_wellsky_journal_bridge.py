"""UHC Authorizations  ->  WellSky journal-note bridge.

Reads the "United Health Care Authorizations" SharePoint list, finds rows that
have been matched to a consumer but whose journal note has NOT yet been entered
into WellSky, logs into the WellSky (SAMS) sandbox, and creates the Journal
entry on that consumer. The row's "WellSky Documentation Status" column is used
as the work queue and (optionally) stamped as each row is handled.

  Ready-to-document row  =  Lookup Status == 'Matched'
                            AND WellSky Documentation Status in (null, 'Not Documented')
                            AND Journal Note and Client ID both non-empty.

Status lifecycle (WellSky Documentation Status choice column):
    Not Documented  ->  In Progress  ->  Documented        (success)
                                     ->  Failed             (error)

This script REUSES:
  - uhc_pipeline.py                (SharePoint auth + list read/patch helpers)
  - WellSky Automation/wellsky.py  (Selenium client with .add_journal())

Guardrails
  - Subject/Comments field-mapping VERIFY before every save (catches the note
    leaking into the Subject line — fill_by_label used to grab Subject's
    textarea; if that regresses, the row fails loudly instead of saving junk).
  - Per-row retry with dialog cleanup + goto_home recovery for transient
    Selenium timeouts (SAMS cold-start, slow searches).
  - Browser AUTO-RESTART on a dead/hung session (ReadTimeout / WebDriver
    death) — the same row is retried on a fresh login, remaining rows continue.
  - Warm-up open_consumer before the loop absorbs the first-search cold start.
  - Failure screenshots saved to ./bridge_shots/.
  - Proactive browser recycle every --restart-every rows.

Nothing is committed by default. Run patterns:

    # DRY RUN — fill+verify ONE ready row, do not save, do not stamp.
    py uhc_wellsky_journal_bridge.py

    # LIVE, today's queue, all on one sandbox consumer, no stamping:
    py uhc_wellsky_journal_bridge.py --save --today --all \
        --target-consumer <sandbox-consumer-id> --no-mark

    # Re-run only specific SharePoint rows:
    py uhc_wellsky_journal_bridge.py --save --today \
        --target-consumer <sandbox-consumer-id> --no-mark --item-ids 1074,1082,1084,1086
"""

from __future__ import annotations

import argparse
import importlib.util
import re
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

# ── Locate & import the two modules we reuse ───────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
PIPELINE_PY = SCRIPT_DIR / "uhc_pipeline.py"
SHOTS_DIR = SCRIPT_DIR / "bridge_shots"

_WS_CANDIDATES = [
    Path(r"C:/Users/eponce/Desktop/AiHub/WellSky Automation"),
    Path.home() / "Desktop" / "AiHub" / "WellSky Automation",
    Path(r"C:/Users/eponce/AiHub/WellSky Automation"),
    Path.home() / "AiHub" / "WellSky Automation",
]
WS_DIR = next((p for p in _WS_CANDIDATES if (p / "wellsky.py").exists()), None)
if WS_DIR is None:
    sys.exit("[fatal] Could not locate 'WellSky Automation/wellsky.py'. "
             "Edit _WS_CANDIDATES in this script to point at it.")
sys.path.insert(0, str(WS_DIR))


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


uhc = _load_module("uhc_pipeline", PIPELINE_PY)
# NOTE: the Claude reviewer no longer runs here. review_notes.py (hourly, on the
# pipeline machine) corrects the JournalNote and stores the Care Plan summary in
# SharePoint ahead of time; this bridge just pushes what's already there.
from wellsky import WellSkyClient  # noqa: E402  (after sys.path insert)

# ── SharePoint column internal names (verified live against the list) ──
COL_LOOKUP_STATUS = "LookupStatus"
COL_DOC_STATUS = "WellSkyDocumentationStatus"
COL_JOURNAL = "JournalNote"
COL_CLIENT_ID = "ClientID"
COL_MEMBER = "MemberName"
COL_CHANGE_TYPE = "ChangeType"
COL_TITLE = "Title"  # authorization number
COL_SERVICES = "Services"          # rendered services blob (fed to the reviewer)
COL_CARE_PLAN = "CarePlanComments"  # where the reviewer's summary is stored

DOC_NOT = "Not Documented"
DOC_INPROGRESS = "In Progress"
DOC_DONE = "Documented"
DOC_FAILED = "Failed"

# ── THE CUTOFF SWITCH ──────────────────────────────────────────────────
# Only auths ADDED to the list on/after this date are pushed to WellSky.
# Everything older was entered by the team by hand and is NEVER touched.
# To hand over to the automation on a given day, set that day here.
PUSH_SINCE = "2026-09-09"   # YYYY-MM-DD
# Override for tests (e.g. a sandbox run on rows added before the cutoff):
#   set UHC_PUSH_SINCE=2026-09-01
import os as _os  # noqa: E402
PUSH_SINCE = _os.environ.get("UHC_PUSH_SINCE", PUSH_SINCE)

# ── Journal-entry Type + Subject shape (matches the WellSky screenshots) ──
# WellSky "Journal Type" dropdown value for authorization notes.
JOURNAL_TYPE = "Service Authorization"
# EXCEPTION (per Grace Guan): coverage decision letters use a different type
# and subject prefix. Coverage decision letters = Termination + Suspension
# change types (CMS-10716). e.g. subject "Coverage decision letter/HM decrease
# (SCO United)", type "Email/Fax Contact ...".
CDL_JOURNAL_TYPE = "Email/Fax Contact (text enter sender/recipient)"
CDL_SUBJECT_PREFIX = "Coverage decision letter/"
CDL_CHANGE_TYPES = {"termination", "suspension"}
# Change types that keep the journal TYPE of a coverage decision letter but
# NOT its subject prefix. The reviewer rewrote "Coverage decision letter/HDM
# termination (SCO United)" as "HDM end auth (SCO United)" on 2026-09-01;
# Grace Guan's rule about the entry TYPE is untouched.
CDL_PLAIN_SUBJECT_CHANGE_TYPES = {"termination"}
# Trailing program label on every subject, e.g. "... auth (SCO United)".
PROGRAM_LABEL = "SCO United"

# ChangeType (SharePoint) -> the verb used in the subject line.
_CHANGE_WORD = {
    "increase": "increase",
    "decrease": "decrease",
    "new": "initialize",
    "initiation": "initialize",   # list now says Initiation; payload still "New"
    "initiate": "initiation",     # the value the list ACTUALLY stores; without
                                  # it every Initiate row fell through to the
                                  # "UHC Authorization - Initiate" fallback
    "renewal": "renewal",
    "termination": "end",
    "suspension": "suspension",
    "records request": "records request",
}

# Service full-name (as it appears in the Services field) -> subject abbrev.
# Longest key wins so "adult day health" isn't shadowed by a shorter match.
_SVC_ABBR = {
    "homemaker": "HM",
    "personal care": "PC",
    "adult day health": "ADH",
    "home delivered meals": "HDM",
    "chore": "HCH 1x",
    "pers": "PERS",
    "emergency response": "PERS",   # UHC's literal wording for PERS lines
    "companion care": "Companion",
    "case management": "CDC",       # UHC's literal wording for the CDC line (team says CDC, not CDS)
    "day care services": "ADH",     # UHC's literal wording (S5100-S5105) for ADH
    "consumer directed": "CDC",
    "laundry": "Laundry",
    "transportation": "Transportation",
}


class SessionDead(Exception):
    """Raised when the WellSky browser/session is no longer usable and the
    only recovery is a full browser restart."""


class FieldMappingError(Exception):
    """Raised when a post-fill readback shows the journal fields are wrong
    (e.g. the note leaked into Subject). Never save in this state."""


# ── row helpers ────────────────────────────────────────────────────────
def is_ready(fields: dict) -> bool:
    if (fields.get(COL_LOOKUP_STATUS) or "").strip() != "Matched":
        return False
    doc = (fields.get(COL_DOC_STATUS) or "").strip()
    if doc not in ("", DOC_NOT):
        return False
    if not (fields.get(COL_JOURNAL) or "").strip():
        return False
    if not str(fields.get(COL_CLIENT_ID) or "").strip():
        return False
    return True


def clean_client_id(raw) -> str:
    cid = str(raw or "").strip()
    if cid.endswith(".0"):
        cid = cid[:-2]
    return cid


def _service_abbrevs(services_text: str, note_text: str = "") -> list[str]:
    """Distinct service abbreviations found in the Services field, in first-seen
    order. The Services column is a human-readable multiline blob like
    'Homemaker (Day): Approved — 3.75 hrs/wk'; we match known service names
    (longest first) and fall back to the leading token for anything unmapped."""
    found: list[str] = []
    seen: set[str] = set()
    # One entry per non-empty line so multi-service auths keep every service.
    for line in (services_text or "").splitlines():
        low = line.strip().lower()
        if not low:
            continue
        abbr = None
        for name in sorted(_SVC_ABBR, key=len, reverse=True):
            if name in low:
                abbr = _SVC_ABBR[name]
                break
        if abbr is None:
            # Unmapped service — use the name up to the first '(' or ':', minus
            # any literal-extract decorations (a trailing HCPCS code like S5102 /
            # T2022 and a [modifier]) so codes never leak into the subject.
            head = line.strip().split("(")[0].split(":")[0].strip()
            head = re.sub(r"\s*\[[^\]]*\]\s*$", "", head)          # "[U1]"
            head = re.sub(r"\s+[A-Z]\d{4}\s*$", "", head).strip()  # " S5102"
            abbr = head or None
        elif abbr == "PERS":
            # Reviewer wants the device spelled out: "PERS cellular w/GPS &
            # fall detection renewal auth (SCO United)", not a bare "PERS".
            # The device wording lives in the subcategory parenthetical that
            # render_services() writes, e.g. "PERS (cellular w/GPS & fall
            # detection): Approved — ...".
            m = re.search(r"\(([^)]+)\)", line)
            if m and m.group(1).strip():
                abbr = f"PERS {m.group(1).strip()}"
            else:
                # Literal extracts carry no parenthetical; the decided note
                # names the device ("PERS cellular renewal 13 units").
                d = re.search(r"\bPERS\s+(cellular|landline)\b", note_text or "", re.I)
                if d:
                    abbr = f"PERS {d.group(1).lower()}"
        if abbr and abbr not in seen:
            seen.add(abbr)
            found.append(abbr)
    # Team feedback 2026-09-09:
    #  - A CDC auth is written as CDC only. The case management (T2022),
    #    per-diem (T1020), TV and 99509 lines are program components, not
    #    separate services -> never "CDC/PC".
    #  - Transportation that rides along with ADH is part of the ADH auth.
    if "CDC" in found:
        return ["CDC"]
    if "ADH" in found:
        found = [a for a in found if a != "Transportation"]
    return found


# Phrases the extracted journal note opens with -> ChangeType, used when the
# ChangeType column is blank (Power Automate doesn't populate it).
_NOTE_CHANGE_PHRASES = [
    ("for renewal of", "Renewal"),
    ("for increase of", "Increase"),
    ("for decrease of", "Decrease"),
    ("for termination of", "Termination"),
    ("for end of", "Termination"),
    ("will be suspended", "Suspension"),
    ("for suspension of", "Suspension"),
]


def resolve_change_type(fields: dict) -> str:
    """ChangeType column if set, else inferred from the note's opening phrase."""
    ct = (fields.get(COL_CHANGE_TYPE) or "").strip()
    if ct and ct.lower() != "null":
        return ct
    note = (fields.get(COL_JOURNAL) or "").lower()
    for phrase, kind in _NOTE_CHANGE_PHRASES:
        if phrase in note:
            return kind
    return ""


def is_coverage_decision_letter(fields: dict) -> bool:
    """Coverage decision letters (CMS-10716) = Termination + Suspension change
    types. They take a different journal Type and subject prefix."""
    return resolve_change_type(fields).lower() in CDL_CHANGE_TYPES


def journal_type_for(fields: dict) -> str:
    return CDL_JOURNAL_TYPE if is_coverage_decision_letter(fields) else JOURNAL_TYPE


def build_subject(fields: dict) -> str:
    """Subject line matching the WellSky screenshots:
        normal auth : 'HM increase auth (SCO United)'
        coverage dec: 'Coverage decision letter/HM decrease (SCO United)'
                      (prefixed, and no 'auth' word)

    Falls back to the prior 'UHC Authorization - <ChangeType>' shape if the
    change type or services can't be resolved, so a save never emits a blank
    or half-formed subject. The dry run prints this for review before --save."""
    ct = resolve_change_type(fields)
    change_word = _CHANGE_WORD.get(ct.lower())
    svc = "/".join(_service_abbrevs(fields.get("Services") or "",
                                    fields.get(COL_JOURNAL) or ""))
    if svc and change_word:
        if (is_coverage_decision_letter(fields)
                and ct.lower() not in CDL_PLAIN_SUBJECT_CHANGE_TYPES):
            return f"{CDL_SUBJECT_PREFIX}{svc} {change_word} ({PROGRAM_LABEL})"
        return f"{svc} {change_word} auth ({PROGRAM_LABEL})"
    # Fallback — keep the old, safe shape rather than emit junk.
    return f"UHC Authorization - {ct}" if ct else "UHC Authorization"


def fetch_ready_rows(g, site_id) -> list[dict]:
    items = uhc.list_items(g, site_id, filter_=f"fields/{COL_LOOKUP_STATUS} eq 'Matched'")
    return [it for it in items
            if is_ready(it.get("fields", {}))
            # THE CUTOFF: skip anything added before PUSH_SINCE (done by hand).
            and (it.get("createdDateTime") or "")[:10] >= PUSH_SINCE]


def stamp(g, site_id, item_id: str, status: str, do_write: bool):
    if not do_write:
        return
    uhc.patch_fields(g, site_id, item_id, {COL_DOC_STATUS: status})


# ── WellSky JS readbacks (verify what actually landed in the form) ─────
_JS_READ_SUBJECT = r"""
const c = document.querySelector('[data-id="control-id_Subject"]');
const t = c && c.querySelector('textarea');
return t ? (t.value || '') : '';
"""

_JS_READ_COMMENTS = r"""
const named = el => el.closest('[data-id^="control-id_"]') !== null;
const ce = [...document.querySelectorAll('[contenteditable="true"]')].find(e => !named(e));
if (ce) return (ce.innerText || '');
const ta = [...document.querySelectorAll('textarea')].find(e => !named(e));
return ta ? (ta.value || '') : '';
"""


def read_subject(w) -> str:
    try:
        return (w.js(_JS_READ_SUBJECT) or "").strip()
    except Exception:
        return ""


def read_comments(w) -> str:
    try:
        return (w.js(_JS_READ_COMMENTS) or "").strip()
    except Exception:
        return ""


# Pure read-only: report the currently-selected Journal Type (never mutates the
# form, so it's safe to call for verification).
_JS_READ_JTYPE = r"""
const c = document.querySelector('[data-id="control-id_JournalTypeUuid"]')
       || document.querySelector('[data-id="JournalTypeUuid"]');
if (!c) return '';
const inp = c.querySelector('input');
if (inp && (inp.value || '').trim()) return inp.value.trim();
return (c.innerText || '').trim();
"""


def read_jtype(w) -> str:
    try:
        return (w.driver.execute_script(_JS_READ_JTYPE) or "").strip()
    except Exception:
        return ""


def fill_field(w, control_id: str, value: str):
    """Fill a textarea control the activities_generator.py way — a REAL click
    to focus (so the clear/select-all stays INSIDE the field), then clear and
    type. wellsky.fill_textarea JS-focuses instead, which often doesn't stick
    on these OpenSilver textareas, so its Ctrl+A select-all hits the whole page
    and wipes the just-committed Journal Type."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    d = w.driver
    ta = WebDriverWait(d, 20).until(EC.element_to_be_clickable(
        (By.XPATH, f"//div[@data-id='control-id_{control_id}']//textarea")))
    ta.click()
    ta.clear()
    ta.send_keys(value)


def set_journal_type(w, value: str) -> str:
    """Select the Journal Type using the picklist pattern proven in
    activities_generator.py: hover+click the anchor to open (a REAL mouse
    move+click, not a JS click — that's what actually commits), type the value,
    press ENTER globally. Under load the type-ahead occasionally commits a
    neighbour (e.g. 'On Site Assessment'), so we read the committed value back
    and RETRY until it matches. Must run on the fresh form. Returns the final
    committed text."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    d = w.driver
    got = ""
    for _ in range(5):
        anchor = WebDriverWait(d, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "//div[@data-id='control-id_JournalTypeUuid']")))
        ActionChains(d).move_to_element(anchor).click().perform()
        time.sleep(0.7)
        ActionChains(d).send_keys(value).perform()
        time.sleep(1.0)
        ActionChains(d).send_keys(Keys.ENTER).perform()
        time.sleep(1.0)
        got = read_jtype(w)
        if got.strip() == value:
            return got
    return got


def session_alive(w) -> bool:
    try:
        _ = w.driver.title
        return True
    except Exception:
        return False


def snap(w, tag: str):
    try:
        SHOTS_DIR.mkdir(exist_ok=True)
        stamp_ = datetime.now().strftime("%H%M%S")
        w.screenshot(str(SHOTS_DIR / f"{tag}_{stamp_}.png"))
    except Exception:
        pass


# ── the actual per-row work (raises on any failure) ────────────────────
def enter_journal(w, item: dict, subject: str, journal: str,
                  client_id: str, *, save: bool):
    """Open the consumer, fill the Journal form, VERIFY field mapping, then
    save. Raises FieldMappingError if the note didn't map correctly, or
    SessionDead if the browser died mid-flow."""
    jtype = journal_type_for(item["fields"])
    jtype_got = ""
    try:
        w.open_consumer(client_id)
        time.sleep(4)
        w.goto_tab("Journals", wait_after=4.0)

        # Build the entry by hand (NOT add_journal) so the Journal TYPE is set
        # FIRST, on the empty form — the only state where the OpenSilver
        # dropdown reliably takes the selection. add_journal fills Subject
        # first, which leaves the type stuck on the consumer's default
        # ('PASRR Non-Compliant'). Fill only — do NOT save yet; verify first.
        w.click_add_new()
        time.sleep(1.0)
        jtype_got = set_journal_type(w, jtype)
        # Do NOT touch EntryDate/EntryTime — WellSky pre-fills them with 'now',
        # and fill_textarea's JS-focus+Ctrl+A clear misfires on those controls,
        # firing a PAGE-WIDE select-all that reverts the Journal Type.
        if journal:
            # WellSky pre-fills a default SIGNATURE in the Comments box on a new
            # journal entry. Do NOT wipe it: read it first, then write our note
            # ON TOP with a blank line before the signature, so the signature is
            # preserved below the note. (fill_by_label clears then types, so we
            # re-supply the signature as part of the value.)
            sig = ""
            for _ in range(3):
                sig = read_comments(w)
                if sig:
                    break
                time.sleep(0.5)
            body_text = f"{journal}\n\n{sig}" if sig else journal
            w.fill_by_label("Comments", body_text)  # contenteditable: focus sticks
        # Subject via a REAL click (fill_field), so its clear stays scoped to
        # the field and never page-selects. Last, so nothing overwrites it.
        time.sleep(0.5)
        fill_field(w, "Subject", subject)
        time.sleep(0.5)
    except Exception as e:  # noqa: BLE001
        if not session_alive(w):
            raise SessionDead(str(e)) from e
        raise  # transient — caller retries within the same session

    if jtype_got and jtype not in jtype_got and jtype_got not in jtype:
        print(f"    [warn] Journal Type is {jtype_got!r}, expected {jtype!r}.")
        snap(w, f"badtype_row{item['id']}")

    # ── VERIFY field mapping before committing ─────────────────────────
    # The readback occasionally returns empty on a slow SAMS render — retry a
    # few times before trusting it, so a slow read doesn't strand a good fill.
    subj_got = body_got = ""
    for _ in range(4):
        time.sleep(1.0)
        subj_got = read_subject(w)
        body_got = read_comments(w)
        if subj_got:
            break

    # Only abort on a REAL mapping error: the Subject holds something OTHER
    # than what we set (e.g. the long note leaked in). An empty readback means
    # the read failed, not that the fill failed — trust the fill and save.
    if subj_got and subj_got != subject:
        snap(w, f"badsubject_row{item['id']}")
        raise FieldMappingError(
            f"Subject readback {subj_got[:60]!r} != expected {subject!r} "
            f"(note may have leaked into Subject) — not saving.")
    if not subj_got:
        print("    [warn] Subject readback empty (slow render) — saving on trust.")
    elif journal and not body_got:
        # Reads are working (subject came back) but the body is genuinely
        # empty — that's a real fill failure.
        snap(w, f"emptybody_row{item['id']}")
        raise FieldMappingError("Comments body empty after fill — not saving.")

    if not save:
        return "dry-run"

    try:
        w.click_save_and_close()
        time.sleep(5)
    except Exception as e:  # noqa: BLE001
        if not session_alive(w):
            raise SessionDead(str(e)) from e
        raise
    return "documented"


def _auth_date_iso(text: str) -> str:
    """First date found in `text` (the summary or note, e.g. 'effective 5/23/26')
    as 'YYYY-MM-DD', for matching against a care plan's date range. '' if none."""
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", text or "")
    if not m:
        return ""
    mo, da, yr = m.group(1), m.group(2), m.group(3)
    if len(yr) == 2:
        yr = "20" + yr
    return f"{yr}-{int(mo):02d}-{int(da):02d}"


def write_care_plan_comment(w, client_id: str, summary: str, auth_date: str, *, save: bool):
    """Write `summary` into the consumer's Care Plan > Comments box.

    `auth_date` is the auth's effective date as 'YYYY-MM-DD'. The plan chosen is
    the one whose Start-End range contains it. Returns 'no-date' if auth_date is
    missing, 'no-grid' if no care-plan grid loaded, or 'no-plan' if no plan
    covers the date — in all three the caller should skip and flag, never guess.

    Path proven live in the sandbox:
      open_consumer (puts driver inside the app iframe) -> Care Plans tab ->
      pick the plan whose date range covers auth_date -> DOUBLE-CLICK its Care
      Program cell (a single click only selects) -> the single visible
      <textarea> (enclosing data-id 'txt_textBoxControl') is the Comments field
      -> append -> Save. Everything runs in the current app-iframe context.

    Returns "documented" on save, "dry-run" when save is False (filled +
    read-back verified, not saved). Raises FieldMappingError if the read-back
    doesn't match, SessionDead if the browser died.

    APPENDS (decided): existing Comments content — prior auths or a case
    manager's own notes — is preserved and our summary is added below it.
    Idempotent: a summary already present is not re-added.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.action_chains import ActionChains

    d = w.driver
    if not (auth_date or "").strip():
        return "no-date"
    try:
        w.open_consumer(client_id)
        time.sleep(4)
        w.goto_tab("Care", "Plans", wait_after=5.0)
        time.sleep(2)

        # Pick the plan whose Start-End range CONTAINS auth_date, then open it by
        # double-clicking its Care Program cell. open_consumer already put the
        # driver INSIDE the OpenSilver app iframe, so everything here runs in the
        # CURRENT context — switching back to default content loses the app frame
        # (that was the source of earlier flakiness). OpenSilver renders each
        # cell as an absolutely-positioned div and DUPLICATES them, so JS groups
        # cells into rows by top offset and takes the row's MIN date as start and
        # MAX date as end (never "second date by position" — that's a duplicate).
        js_pick = r"""
        const auth = arguments[0];
        const toISO = s => { const m=(s||'').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
          if(!m) return null; let y=m[3]; if(y.length===2) y='20'+y;
          return y+'-'+String(m[1]).padStart(2,'0')+'-'+String(m[2]).padStart(2,'0'); };
        const cells=[];
        document.querySelectorAll('.opensilver-uielement').forEach(e=>{
          if(e.offsetParent===null) return;
          const t=(e.innerText||'').trim(); if(!t) return;
          const isDate=/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(t);
          const r=e.getBoundingClientRect();
          cells.push({e,t,top:Math.round(r.top),left:Math.round(r.left),isDate});
        });
        const rows=[];
        cells.forEach(c=>{ let row=rows.find(r=>Math.abs(r.top-c.top)<=6);
          if(!row){row={dates:[],progs:[]}; rows.push(row);}
          if(c.isDate) row.dates.push({iso:toISO(c.t),left:c.left});
          else if(c.t.length>6) row.progs.push({e:c.e,left:c.left}); });
        let hadGrid=false;
        for(const row of rows){
          const isos=row.dates.map(x=>x.iso).filter(Boolean).sort();
          if(isos.length<2 || row.progs.length===0) continue;
          hadGrid=true;
          const start=isos[0], end=isos[isos.length-1];
          const maxD=Math.max.apply(null, row.dates.map(x=>x.left));
          if(start && end && auth>=start && auth<=end){
            row.progs.sort((a,b)=>a.left-b.left);
            const prog=row.progs.find(p=>p.left>maxD) || row.progs[0];
            prog.e.setAttribute('data-pick','1'); return 'ok';
          }
        }
        return hadGrid ? 'no-plan' : 'no-grid';
        """
        res = d.execute_script(js_pick, auth_date)
        if res != "ok":
            # No plan covers the auth date (or no grid) -> don't guess.
            return res  # 'no-plan' or 'no-grid'; caller skips + flags
        cell = d.find_element(By.CSS_SELECTOR, "[data-pick='1']")
        ActionChains(d).move_to_element(cell).double_click().perform()
        time.sleep(8)

        # The plan detail renders in the SAME app iframe. The Comments box is the
        # single visible <textarea> (enclosing control 'txt_textBoxControl').
        tas = [t for t in d.find_elements(By.TAG_NAME, "textarea")
               if t.is_displayed()]
        if not (tas and d.find_elements(
                By.CSS_SELECTOR, "[data-id='txt_textBoxControl']")):
            raise FieldMappingError("Care Plan Comments textarea not found.")
        target_ta = tas[0]

        # APPEND: keep whatever is already in the Comments box (prior auths, or
        # notes a case manager typed) and add our summary below it. Idempotent —
        # if this exact summary is already present (a re-run), leave it alone.
        existing = (target_ta.get_attribute("value") or "").rstrip()
        want = summary.strip()
        if want and want in existing:
            return "dry-run" if not save else "documented"  # already there
        new_value = f"{existing}\n{want}" if existing else want
        target_ta.click()
        target_ta.clear()
        target_ta.send_keys(new_value)
        time.sleep(1)
        back = (target_ta.get_attribute("value") or "")
        # Verify our summary landed AND the prior content survived.
        if want not in back or (existing and existing not in back):
            raise FieldMappingError(
                "Care Plan Comments read-back mismatch — not saving.")
    except FieldMappingError:
        raise
    except Exception as e:  # noqa: BLE001
        if not session_alive(w):
            raise SessionDead(str(e)) from e
        raise

    if not save:
        return "dry-run"
    # Save and Close on the plan-detail toolbar.
    for xp in ("//a[normalize-space(text())='Save and Close']",
               "//span[normalize-space(text())='Save and Close']",
               "//a[normalize-space(text())='Save']",
               "//span[normalize-space(text())='Save']"):
        els = [e for e in d.find_elements(By.XPATH, xp) if e.is_displayed()]
        if els:
            d.execute_script("arguments[0].click();", els[0])
            time.sleep(5)
            break
    d.switch_to.default_content()
    return "documented"


def recover(w):
    """Best-effort: close any half-open dialog and return to the search bar so
    the next attempt starts clean. Silently ignores failures."""
    try:
        w.close_any_dialog()
    except Exception:
        pass
    try:
        w.goto_home()
    except Exception:
        pass


# ── client lifecycle ───────────────────────────────────────────────────
def new_client(args) -> WellSkyClient:
    kwargs = {"username": args.username}
    if args.password:
        kwargs["password"] = args.password
    w = WellSkyClient(**kwargs)
    print("    [session] logging in …")
    w.login()
    # Bump the HTTP read timeout to the driver so a slow-but-alive action
    # isn't killed at the 120s default (best-effort; API varies by Selenium).
    try:
        w.driver.command_executor._client_config.timeout = 300  # noqa: SLF001
    except Exception:
        pass
    return w


def warmup(w, consumer_id: str):
    """Absorb the SAMS cold-start on a throwaway open so it doesn't cost a
    real row its first attempt."""
    try:
        print(f"    [warmup] priming search on {consumer_id} …")
        w.open_consumer(consumer_id)
        w.to_default_content()
    except Exception as e:  # noqa: BLE001
        print(f"    [warmup] non-fatal: {type(e).__name__}: {e}")


def run_row(w, args, item, subject, journal, client_id, warm_id):
    """Do one row with full resilience. Returns (result, client, failure_tag).

    - Retries transient failures up to --retries within the live session
      (dialog cleanup + goto_home between attempts).
    - On a dead/hung session, restarts the browser up to --max-restarts times
      and retries the SAME row on the fresh login.
    - A FieldMappingError never saves and never retries — it's a hard abort.

    The (possibly new) client is returned so the caller keeps using it."""
    restarts = 0
    while True:
        for attempt in range(1, args.retries + 1):
            try:
                result = enter_journal(w, item, subject, journal,
                                       client_id, save=args.save)
                return result, w, None
            except FieldMappingError as e:
                print(f"    [ABORT] {e}")
                recover(w)
                return "failed", w, f"{item['id']} (mapping)"
            except SessionDead as e:
                print(f"    [session] DEAD mid-row: {str(e)[:80]}")
                break  # leave attempt loop -> restart below
            except Exception as e:  # noqa: BLE001 — transient
                print(f"    [retry {attempt}/{args.retries}] "
                      f"{type(e).__name__}: {str(e)[:80]}")
                snap(w, f"retry_row{item['id']}")
                recover(w)
                if attempt == args.retries:
                    return "failed", w, f"{item['id']} (timeout)"

        # Only reached via SessionDead break — recycle the browser.
        restarts += 1
        if restarts > args.max_restarts:
            print("    [session] restart cap reached — giving up on this row.")
            snap(w, f"deadgiveup_row{item['id']}")
            return "failed", w, f"{item['id']} (session)"
        print(f"    [session] restarting browser ({restarts}/{args.max_restarts}) …")
        try:
            w.close()
        except Exception:
            pass
        w = new_client(args)
        warmup(w, warm_id)


# ── main ───────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--save", action="store_true",
                    help="Commit the journal (and stamp status). Omit = dry run.")
    ap.add_argument("--all", action="store_true", help="Process every ready row.")
    ap.add_argument("--limit", type=int, default=1,
                    help="Max rows (default 1). Ignored with --all.")
    ap.add_argument("--client-id", default=None,
                    help="Only the ready row(s) for this Client ID.")
    ap.add_argument("--item-ids", default=None,
                    help="Comma-separated SharePoint item IDs to restrict to "
                         "(e.g. re-run specific failed rows).")
    ap.add_argument("--today", action="store_true",
                    help="Only rows whose Review Date is today.")
    ap.add_argument("--review-date", default=None,
                    help="Only rows whose Review Date == this YYYY-MM-DD.")
    ap.add_argument("--target-consumer", default=None,
                    help="Enter EVERY note on this one WellSky consumer id "
                         "instead of each row's real Client ID (sandbox test).")
    ap.add_argument("--no-mark", action="store_true",
                    help="Do NOT write WellSky Documentation Status back.")
    ap.add_argument("--retries", type=int, default=2,
                    help="Attempts per row within a live session (default 2).")
    ap.add_argument("--restart-every", type=int, default=6,
                    help="Proactively recycle the browser every N rows "
                         "(0 = never). Default 6.")
    ap.add_argument("--max-restarts", type=int, default=2,
                    help="Browser restarts allowed per row on a dead session "
                         "(default 2).")
    ap.add_argument("--username", default="CBES5")
    ap.add_argument("--password", default="Password63!",
                    help="WellSky sandbox password (baked default so the "
                         "Task Scheduler job needs no flag; override if rotated).")
    ap.add_argument("--keep-open", action="store_true",
                    help="Leave the browser open after the run (for manual "
                         "review). The process stays alive until you kill it.")
    args = ap.parse_args()

    mark = not args.no_mark

    # ── SharePoint side ────────────────────────────────────────────────
    print("[*] Authenticating to SharePoint (Graph) …")
    creds = uhc.load_azure_creds()
    app = uhc._msal_app(creds)
    token = uhc.get_token(app, "https://graph.microsoft.com/.default")
    g = uhc.graph_session(token)
    site_id = uhc.resolve_site_id(g)

    print("[*] Fetching Matched + Not-Documented rows …")
    rows = fetch_ready_rows(g, site_id)

    if args.client_id:
        want = clean_client_id(args.client_id)
        rows = [r for r in rows
                if clean_client_id(r["fields"].get(COL_CLIENT_ID)) == want]

    review_date = args.review_date or (datetime.now().strftime("%Y-%m-%d")
                                       if args.today else None)
    if review_date:
        rows = [r for r in rows
                if (r["fields"].get("ReviewDate") or "")[:10] == review_date]
        print(f"[*] Filtered to Review Date == {review_date}.")

    if args.item_ids:
        wanted = {s.strip() for s in args.item_ids.split(",") if s.strip()}
        rows = [r for r in rows if str(r["id"]) in wanted]
        print(f"[*] Restricted to item IDs {sorted(wanted)}.")

    print(f"[*] {len(rows)} ready row(s).")
    if not rows:
        print("[*] Nothing to document. Done.")
        return

    if not args.all:
        rows = rows[: max(0, args.limit)]
    print(f"[*] Processing {len(rows)} row(s)  "
          f"(save={args.save}, stamp={mark and args.save}, "
          f"retries={args.retries}, restart_every={args.restart_every}).")

    counts = {"documented": 0, "failed": 0, "dry-run": 0}
    failures: list[str] = []

    w = new_client(args)
    warm_id = clean_client_id(
        args.target_consumer or rows[0]["fields"].get(COL_CLIENT_ID))
    warmup(w, warm_id)

    try:
        done_since_restart = 0
        for n, item in enumerate(rows, 1):
            fields = item["fields"]
            item_id = item["id"]
            row_client = clean_client_id(fields.get(COL_CLIENT_ID))
            client_id = (clean_client_id(args.target_consumer)
                         if args.target_consumer else row_client)
            # Logs identify the row by auth number only — never the member
            # name — so console output / scheduler logs carry no PHI.
            auth_no = (fields.get(COL_TITLE) or "").strip() or "(no auth #)"
            # Both already corrected/stored in SharePoint by review_notes.py.
            journal = (fields.get(COL_JOURNAL) or "").strip()
            care_summary = (fields.get(COL_CARE_PLAN) or "").strip()
            subject = build_subject(fields)

            tag = f"client={client_id}" + (f" (row {row_client})"
                                           if args.target_consumer else "")
            print(f"\n[{n}/{len(rows)}] row {item_id}  auth {auth_no}  {tag}")
            print(f"    type   : {journal_type_for(fields)}")
            print(f"    subject: {subject}")
            print(f"    note   : {journal[:100]}{'…' if len(journal) > 100 else ''}")
            if care_summary:
                print(f"    summary: {care_summary.replace(chr(10), ' | ')[:120]}")

            # Proactive browser recycle to avoid the memory/DOM buildup that
            # hung the session mid-run last time.
            if args.restart_every and done_since_restart >= args.restart_every:
                print("    [session] proactive recycle …")
                try:
                    w.close()
                except Exception:
                    pass
                w = new_client(args)
                warmup(w, warm_id)
                done_since_restart = 0

            stamp(g, site_id, item_id, DOC_INPROGRESS, mark and args.save)

            result, w, fail_tag = run_row(w, args, item, subject, journal,
                                          client_id, warm_id)

            if result == "failed":
                stamp(g, site_id, item_id, DOC_FAILED, mark and args.save)
                if fail_tag:
                    failures.append(fail_tag)
                print("    [failed]")
            elif result == "dry-run":
                print("    [dry-run] filled + VERIFIED, not saved, not stamped.")
            else:  # documented
                stamp(g, site_id, item_id, DOC_DONE, mark)
                print(f"    [ok] saved; status -> {DOC_DONE}"
                      f"{' (stamp skipped)' if not mark else ''}")

            # WellSky Care Plan Comments: append the summary to the plan whose
            # date range covers the auth. Best-effort — never fails the row; if
            # no plan matches the auth date it is skipped and flagged, not guessed.
            if care_summary and result in ("documented", "dry-run"):
                ad = _auth_date_iso(care_summary) or _auth_date_iso(journal)
                try:
                    cp = write_care_plan_comment(
                        w, client_id, care_summary, ad, save=args.save)
                    if cp in ("no-date", "no-grid", "no-plan"):
                        print(f"    [care plan] SKIPPED + FLAG: {cp} "
                              f"(auth date {ad or '?'} matched no plan)")
                    else:
                        print(f"    [care plan] {cp}")
                except SessionDead as e:  # noqa: BLE001
                    print(f"    [care plan] session died: {e} — recycling")
                    try:
                        w.close()
                    except Exception:
                        pass
                    w = new_client(args)
                    warmup(w, warm_id)
                except FieldMappingError as e:  # noqa: BLE001
                    print(f"    [care plan] skipped: {e}")

            counts[result] = counts.get(result, 0) + 1
            done_since_restart += 1

        if not args.save:
            print("\n[dry-run] leaving the browser open 20s to inspect …")
            time.sleep(20)
    except Exception as e:  # noqa: BLE001 — last-ditch, keep the summary
        print(f"\n[fatal] {type(e).__name__}: {e}")
        traceback.print_exc()
    finally:
        if args.keep_open:
            print("\n[keep-open] browser left open for review — "
                  "kill this process to close it.")
            try:
                while True:
                    time.sleep(3600)
            except KeyboardInterrupt:
                pass
        try:
            w.close()
        except Exception:
            pass

    print("\n[summary] " + ", ".join(f"{k}={v}" for k, v in counts.items()))
    if failures:
        print("[failed rows] " + ", ".join(failures))
        ids = ",".join(f.split()[0] for f in failures)
        print(f"[retry cmd] add:  --item-ids {ids}")


if __name__ == "__main__":
    main()
