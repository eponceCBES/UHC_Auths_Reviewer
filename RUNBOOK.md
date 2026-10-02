# UHC auths → WellSky: the push flow (runbook)

This is the sequence used for every production push so far (first five laundry
auths 9/10, the 24-row laundry batch 9/11). Same steps every time.

All commands run from `C:\Users\eponce\AiHub\uhc_reviewer` with
`C:\Users\eponce\AppData\Local\Programs\Python\Python313\python.exe`.

## 1. Review the batch (read it, don't trust it)

    py tools/review_batch.py 2026-09-10

Prints every auth created on/after that date with its journal note, care-plan
comment and the rule check (`OK` / `!!` with the reason), then a total line
("N auths, M flagged"). Read the notes. A flagged row is fixed BEFORE pushing:
clear its `JournalNote` + `CarePlanComments` in SharePoint and run the pipeline
incrementally (`py uhc_pipeline.py --force --no-excel --no-calendar --no-deploy
--no-suspensions --no-plan-check --no-rematch`); if the fix is a rule, it goes
into `compose.py` (rule text + `lint()` check + a case in `test_compose.py`) first.

## 2. Decide what goes

The bridge only takes rows that are `Matched`, have a note and a Client ID, are
not `Documented`, and were added on/after the cutoff (`PUSH_SINCE` in the
bridge; `UHC_PUSH_SINCE=YYYY-MM-DD` overrides it for one window). Restrict to a
set with `--item-ids` when pushing a specific batch.

Care-plan comments are decided by the bridge itself: a laundry auth with no
detailed notes gets the journal note only; every other auth gets its summary
appended to the plan whose dates cover the auth.

## 3. Push

    $env:UHC_PUSH_SINCE = '2026-09-08'      # only if the batch is before the cutoff
    py uhc_wellsky_journal_bridge.py --save --all --item-ids <ids> --username <prod upn> --password -

`--password -` prompts; a scheduled run uses the stored password instead
(`--save-password` once). Drop `--save` for a dry run (fills, verifies, saves
nothing). Each row prints `[ok] saved; status -> Documented`. The bridge closes
the consumer window after every row, checks the right consumer is on top before
writing, recovers a dead browser on its own, and retries failed rows once at the
end of the run and again on later runs (up to 5 attempts per row).

## 4. Confirm

Re-run step 1 for the date, or check `WellSkyDocumentationStatus` in the list:
every pushed row reads `Documented`. A `Failed` row is retried automatically by
the next run; nothing needs resetting by hand.

## 5. Reports

    py tools/build_reports.py

Overwrites `..\uhc_reviewer_deliverables\reports\UHC_Auths_Sep5-9_2026.xlsx`
(all auths since 9/5 with worker + WellSky status + link) and
`UHC_Laundry_Auths_by_GSSC.xlsx` (laundry, one tab per worker).

## Hourly, unattended

`run_hourly.bat` = `git pull` → pipeline → reviewer; add the bridge line from
TRANSFER.md §8 after them once the team is ready for automatic pushes.

## Rules added 2026-10-02 (renewal-notes review, 13 comments)

Prompt + `lint()` + `test_compose.py` cases, same as every rule:

* **Schedule in the note.** Hour services carry the printed schedule in the
  note, not only the summary: `HM renewal 3 hrs/wk (weekday)`. "(Day)" hours
  and "Weekday Hours" both mean weekday.
* **HDM cultural + dietary type.** Both go in the note and summary whenever
  printed, even "Regular": `..., Chinese cultural meal, regular diet`.
* **ADH level + center.** Named when printed; when the auth prints neither,
  the note ends `ADH level and center not specified on the authorization,
  GSSC to confirm with SCO United.` Never invented.
* **PERS device type.** Only landline/cellular, only when printed. Features
  ("Fall Detection") follow the units. No type printed → the note ends
  `PERS device type (landline/cellular) not specified on the authorization,
  GSSC to follow up with SCO United.`
* **Member Letter** (new change type). The member's copy of an approval letter
  ("Member Service number on the back of your member ID card", "cc: Central
  Boston Elder Services") is not an authorization to CBES. The note says so,
  and the bridge never pushes a Member Letter row to WellSky.
* **Laundry follow-up party.** The laundry no-notes template says `Program
  Manager was notified for follow up with SCO United` and the summary
  `(no detailed notes, Program Manager notified)`. "GSSC" there is a violation
  (team 2026-10-02, row 92 of the Sept renewal report).
* **PERS features in the summary**, **"Companion" never "COMP"**, and the
  **ADH center from the WellSky service plan** (`service_plan_adh_center`,
  looked up after the consumer match) — same review.
* Old-format rows (AI Builder payloads from before the EXTRACT-ONLY prompt)
  are rewritten with `py tools/rewrite_legacy_notes.py --item-ids ...`.
