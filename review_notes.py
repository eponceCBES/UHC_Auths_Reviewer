"""Hourly note reviewer: corrects UHC journal notes IN SHAREPOINT, in place.

For every authorization row that has a JournalNote but no CarePlanComments yet
(= not yet reviewed), runs the Claude reviewer (note_reviewer.review_auth) and
writes back:
  * JournalNote        -> the corrected note (abbreviations, units->hours,
                          "one time", word order) -- only if it changed
  * CarePlanComments   -> the Care Plan summary ("Auth: HDM 7 meals/wk ...")

Rows already reviewed (CarePlanComments set) are skipped, so this is safe to run
hourly: the FIRST run backfills the backlog, later runs only touch new auths.
A row whose review fails (empty summary) is left untouched and retried next run.

The WellSky bridge (uhc_wellsky_journal_bridge.py) then just pushes whatever is
in SharePoint -- already correct -- with no Claude call at push time.

Run on the pipeline machine (needs Claude Code installed + logged in there):
    py review_notes.py                 # up to --limit rows (default 150/run)
    py review_notes.py --all           # full backfill in one go (~2k rows, hours)
    py review_notes.py --dry-run       # review + report, write nothing

PHI: notes are processed in memory and written back to the list; only COUNTS
are ever printed or logged. The `Notes` column (contains DOBs) is never read.
"""

import argparse
import importlib.util
import json
import re
import sys
import time
from pathlib import Path

import requests
from cryptography.fernet import Fernet

HERE = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location("note_reviewer", HERE / "note_reviewer.py")
nr = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(nr)
_cspec = importlib.util.spec_from_file_location("compose", HERE / "compose.py")
cz = importlib.util.module_from_spec(_cspec)   # for cz.lint (rule guard)
_cspec.loader.exec_module(cz)

# Same app credentials the pipeline uses (Report Subscriptions is OneDrive-synced
# to the pipeline machine under the same user path).
ENC_DIR = Path(r"C:\Users\eponce\OneDrive - Central Boston Elder Services, Inc"
               r"\Report Subscriptions\Azure\Azure Encryption")
GRAPH = "https://graph.microsoft.com/v1.0"
SITE_PATH = "centralboston.sharepoint.com:/sites/DataManagement"
LIST_ID = "9a466e43-6e94-458f-a611-ecf372914315"   # United Health Care Authorizations

COL_NOTE = "JournalNote"
COL_SERVICES = "Services"
COL_CARE_PLAN = "CarePlanComments"
# Backup of the note as it was BEFORE correction, so any correction can be
# rolled back. Written only the first time a note is changed.
COL_ORIG = "JournalNoteOriginal"


def graph_token() -> str:
    c = json.loads(Fernet((ENC_DIR / "secret.key").read_bytes())
                   .decrypt((ENC_DIR / "encrypted_creds.bin").read_bytes()).decode())
    r = requests.post(
        f"https://login.microsoftonline.com/{c['tenant_id']}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": c["client_id"],
              "client_secret": c["client_secret"],
              "scope": "https://graph.microsoft.com/.default"},
        timeout=60)
    r.raise_for_status()
    return r.json()["access_token"]


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--limit", type=int, default=150,
                    help="Max rows to review this run (keeps hourly runs short).")
    ap.add_argument("--all", action="store_true", help="Review every unreviewed row.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Review and report counts, but write nothing back.")
    ap.add_argument("--relint", action="store_true",
                    help="Also re-review rows already reviewed whose current note/comment "
                         "fails the rule guard (compose.lint). Nothing is cleared up front.")
    ap.add_argument("--since", default="",
                    help="Only rows ADDED to the list on/after this date "
                         "(YYYY-MM-DD), e.g. --since 2026-08-01 for the backfill.")
    args = ap.parse_args()

    tok = graph_token()
    H = {"Authorization": f"Bearer {tok}"}
    site = requests.get(f"{GRAPH}/sites/{SITE_PATH}", headers=H, timeout=60).json()["id"]

    # Collect unreviewed rows: has a note, no summary yet.
    todo = []
    sel = f"{COL_NOTE},{COL_SERVICES},{COL_CARE_PLAN},{COL_ORIG}"
    url = (f"{GRAPH}/sites/{site}/lists/{LIST_ID}/items"
           f"?$expand=fields($select={sel})&$top=5000")
    total = 0
    while url:
        d = requests.get(url, headers=H, timeout=120).json()
        for it in d.get("value", []):
            total += 1
            # --since: only rows ADDED to the list on/after the cutoff.
            if args.since and (it.get("createdDateTime") or "")[:10] < args.since:
                continue
            f = it.get("fields") or {}
            note = (f.get(COL_NOTE) or "").strip()
            svcs = (f.get(COL_SERVICES) or "").strip()
            plan = (f.get(COL_CARE_PLAN) or "").strip()
            if note and not plan:
                todo.append((it["id"], note, svcs, bool((f.get(COL_ORIG) or "").strip())))
            elif note and args.relint and cz.lint("", note, plan, {"services": svcs}):
                # --relint: already reviewed, but the current text fails the
                # rules -> review it again (nothing is cleared up front; a
                # rejected retry leaves the row exactly as it was).
                todo.append((it["id"], note, svcs, bool((f.get(COL_ORIG) or "").strip())))
        url = d.get("@odata.nextLink")

    unreviewed = len(todo)
    if not args.all:
        todo = todo[: max(0, args.limit)]
    print(f"[review_notes] rows={total} unreviewed={unreviewed} "
          f"this_run={len(todo)} dry_run={args.dry_run}", flush=True)

    processed = changed = failed = 0
    t0 = time.time()
    for n, (item_id, note, svcs, has_orig) in enumerate(todo, 1):
        corrected, summary = nr.review_auth(note, svcs)
        if not summary:
            failed += 1          # reviewer failed -> leave untouched, retry next run
            continue
        # Same deterministic rule guard as the Claude-decides path: one retry
        # with the violations spelled out, then leave the row for next run.
        ct = "Termination" if re.search(r"\bend(ed)? (of|effective)\b", corrected, re.I) else ""
        viol = cz.lint(ct, corrected, summary, {"services": svcs})
        if viol:
            corrected, summary = nr.review_auth(
                note, svcs, extra="YOUR PREVIOUS ANSWER BROKE THESE RULES - fix them and answer again:\n- "
                + "\n- ".join(viol))
            viol = cz.lint(ct, corrected, summary, {"services": svcs}) if summary else viol
        if viol:
            failed += 1
            print(f"[review_notes] row {item_id} rejected by rule guard: {'; '.join(viol)}", flush=True)
            continue
        body = {COL_CARE_PLAN: summary}
        if corrected.strip() != note.strip():
            body[COL_NOTE] = corrected
            if not has_orig:
                body[COL_ORIG] = note   # keep the pre-correction note for rollback
            changed += 1
        if not args.dry_run:
            r = requests.patch(
                f"{GRAPH}/sites/{site}/lists/{LIST_ID}/items/{item_id}/fields",
                headers={**H, "Content-Type": "application/json"},
                json=body, timeout=60)
            if r.status_code >= 400:
                failed += 1
                print(f"[review_notes] write failed for row {item_id}: HTTP {r.status_code}",
                      flush=True)
                continue
        processed += 1
        if n % 25 == 0:
            print(f"[review_notes] {n}/{len(todo)} ({int(time.time() - t0)}s)", flush=True)

    print(f"[review_notes] done: reviewed={processed} notes_changed={changed} "
          f"failed_or_retry={failed} remaining_unreviewed={unreviewed - processed} "
          f"in {int(time.time() - t0)}s", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
