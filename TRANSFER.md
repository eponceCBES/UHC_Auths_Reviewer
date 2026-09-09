# UHC note reviewer — set up on the hourly machine

> Also in this folder: `compose.py` + `uhc_extraction_prompt_EXTRACT_ONLY.txt` — the
> "Claude decides" upgrade (mini copies literally, Claude sets change type + writes
> the note, PHI-guarded). NOT deployed yet; see `INTEGRATION.md` before touching
> Power Apps. `py test_compose.py --offline` proves the PHI guard.

## 1. Copy this folder
Copy the whole `uhc_reviewer` folder to the same path on the new machine:
`C:\Users\eponce\AiHub\uhc_reviewer\`
(Same path matters — a few paths are hardcoded, see step 5.)

## 2. Prerequisites on that machine
- Python 3.13, then: `py -3.13 -m pip install requests cryptography pandas openpyxl selenium`
- **Claude Code installed and logged in** (the reviewer runs `claude -p`). Test: `claude -p "say OK"`
- OneDrive synced, so `Report Subscriptions\Azure\Azure Encryption` exists (the Graph credentials).
- `C:\Users\eponce\AiHub\WellSky Automation\wellsky.py` present (the bridge imports it).

## 3. uhc_pipeline.py is INCLUDED (wired, schema-adaptive)
This folder's `uhc_pipeline.py` is the production pipeline plus the "Claude decides"
wiring in `build_field_body`. It is schema-adaptive: rows from the OLD Power Apps
prompt take the original code path unchanged; rows from the EXTRACT-ONLY prompt go
through `compose.py`. So it is SAFE to run this pipeline before the prompt is swapped.

Deploy order (never the other way round):
  1. run this pipeline (hourly job) instead of the old one
  2. verify:  py -3.13 test_pipeline_wiring.py --offline   (no Claude call)
              py -3.13 test_pipeline_wiring.py             (one Claude call)
  3. THEN paste `uhc_extraction_prompt_EXTRACT_ONLY.txt` into Power Apps
Details and rollback: INTEGRATION.md.

## 4. Fix the WellSky login (stale in wellsky.py)
In `WellSky Automation\wellsky.py` set:
- username: `CBES5@agingnetwork.com`  (must be the full address)
- password: `Password64!`

## 5. Paths to verify (only if the user profile differs from `eponce`)
- `review_notes.py` → `ENC_DIR` (Azure Encryption folder)
- `note_reviewer.py` → `_find_claude()` falls back to `%APPDATA%\npm\...\bin\claude.exe`
- `uhc_wellsky_journal_bridge.py` → `_WS_CANDIDATES` (WellSky Automation folder)

## 6. Test the reviewer (writes nothing)
    cd C:\Users\eponce\AiHub\uhc_reviewer
    py -3.13 review_notes.py --dry-run --limit 2

## 6b. The pipeline needs ONE env var (it lives outside Report Subscriptions now)
It finds its credentials/consumer CSV by walking up to "Report Subscriptions"; from
AiHub it can't. Set this once on the machine (then re-open any console):
    setx REPORT_SUBSCRIPTIONS_DIR "C:\Users\eponce\OneDrive - Central Boston Elder Services, Inc\Report Subscriptions"

Run the pipeline on demand (processes pending rows now; ~2 min; Claude only for new-format rows):
    cd C:\Users\eponce\AiHub\uhc_reviewer
    py -3.13 uhc_pipeline.py --force --no-excel --no-calendar --no-deploy --no-suspensions --no-plan-check --no-rematch
(The hourly job should point at THIS folder's uhc_pipeline.py without the --no-* flags.)

NOTE: resubmitting Power Automate runs BEFORE the new prompt is published in AI Builder
just creates old-format DUPLICATE rows; the pipeline flags them as duplicates and skips
them (harmless). Publish the prompt first, then resubmit.

## 7. Schedule it hourly
    schtasks /Create /TN "CBES UHC Note Reviewer" /SC HOURLY /ST 06:00 ^
      /TR "\"C:\Users\eponce\AppData\Local\Programs\Python\Python313\python.exe\" \"C:\Users\eponce\AiHub\uhc_reviewer\review_notes.py\"" ^
      /RL LIMITED /F

It only touches new auths each run (skips anything already reviewed). Default model is Opus 5.
For a big backfill run: `set CLAUDE_REVIEW_MODEL=claude-sonnet-5` first.

## 8. When ready to push to WellSky (the bridge)
- Set the cutoff in `uhc_wellsky_journal_bridge.py`: `PUSH_SINCE = "2026-09-09"` (your Wednesday).
  Nothing added to the list before that date is ever pushed.
- Dry-run first (fills WellSky, saves nothing): `py -3.13 uhc_wellsky_journal_bridge.py --limit 1`
- Then for real: add `--save`.
- Sandbox test on rows older than the cutoff: `set UHC_PUSH_SINCE=2026-09-01` for that run only.

## 9. Deploying changes with git (replaces the zip)
This folder is a git repo. **Code only** — `.gitignore` keeps out everything generated
from the list (HTML, logs, CSV/XLSX, screenshots, run state) and the secrets (deploy
token, keys). A pre-commit hook (`.githooks/phi_check.py`) blocks a commit that contains
an auth number, client ID, Medicaid ID, DOB or SSN shape. Git history is permanent, so
the check happens before the commit exists.

One-time, dev machine (this one):
    git remote add origin <private repo url>      # GitHub private or Azure DevOps Repos
    git push -u origin main

One-time, scheduler machine:
    git clone <private repo url> C:\AiHub\uhc_reviewer
    cd C:\AiHub\uhc_reviewer
    git config core.hooksPath .githooks
    setx REPORT_SUBSCRIPTIONS_DIR "C:\Users\<user>\OneDrive - Central Boston Elder Services, Inc\Report Subscriptions"
    setx CLAUDE_REVIEW_MODEL claude-opus-5
Then point the hourly Task Scheduler action at `run_hourly.bat` (Start in: the clone folder).
It does `git pull --ff-only`, then the pipeline, then `review_notes.py`.

Every change after that: edit here → commit → `git push`. The next hourly run picks it up.
Rollback: `git revert <commit>` + push (or `git checkout <good commit>` on the scheduler box).

Secrets never enter the repo: Graph creds stay in `Report Subscriptions\Azure\Azure Encryption`,
the SWA deploy token in its original UHC folder (the pipeline falls back to it), and the
WellSky password is passed on the command line / kept in `wellsky.py` outside this repo.
