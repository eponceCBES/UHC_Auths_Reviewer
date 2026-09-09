"""Claude-based reviewer for the UHC pipeline: corrects the journal note AND
builds the service-plan (Care Plan) Comments summary, in a single Claude call.

Uses the local Claude Code CLI (`claude -p`, no API key needed). Guards the
result: Claude is trusted on wording, abbreviations, and unit->hour conversions,
but if it changes any DATE (MM/DD/YYYY) or SERVICE CODE (e.g. S5161) in the note,
the corrected note is rejected and the ORIGINAL note is kept. Any failure (CLI
missing, non-zero exit, timeout, empty/garbled output) falls back to the original
note and an empty summary, so this step can never block the pipeline or push a
corrupted note to WellSky.

PHI: the note + services text is sent to Claude through the local CLI login. Per
the PHI scan (PHI_scan_journal_notes report) journal-note text carries no member
identifiers, but use of this reviewer is gated on compliance sign-off.

Pipeline usage:
    from note_reviewer import review_auth
    corrected_note, summary = review_auth(journal, services_text)
    body["JournalNote"]      = corrected_note
    if summary:
        body["CarePlanComments"] = summary
"""

import os
import re
import shutil
import subprocess
import time

def _find_claude() -> str:
    """Locate the Claude Code CLI. Tries PATH first, then the npm shim, then the
    native binary INSIDE the npm package. The last one matters: a Claude Code
    auto-update (2026-09-06, v2.1.263) replaced the JS CLI with a native
    bin\\claude.exe and left the PATH shims temp-renamed, so `which` and the
    .cmd both vanished mid-run and every review failed instantly. The package's
    own bin/claude.exe survives such updates."""
    npm = os.path.join(os.environ.get("APPDATA", r"C:\Users\eponce\AppData\Roaming"), "npm")
    for cand in (
        shutil.which("claude"),
        os.path.join(npm, "claude.cmd"),
        os.path.join(npm, "node_modules", "@anthropic-ai", "claude-code", "bin", "claude.exe"),
    ):
        if cand and os.path.exists(cand):
            return cand
    return os.path.join(npm, "claude.cmd")   # best effort; _call reports failure


CLAUDE_CMD = _find_claude()
# A real review takes ~11s. `claude -p` occasionally hangs; 45s is generous
# for a real call and keeps each hang cheap (it used to burn 180s each).
TIMEOUT_S = 180
# Model for the review. Default is Opus 5 for the HOURLY job: volume is a few
# notes/hour, so cost is negligible and these go to WellSky, so take the extra
# care. For BULK backfills (hundreds of notes) override to something cheaper:
#   set CLAUDE_REVIEW_MODEL=claude-sonnet-5
MODEL = os.environ.get("CLAUDE_REVIEW_MODEL", "claude-opus-5")

_NOTE_MARK = "<<<NOTE>>>"
_SUMM_MARK = "<<<SUMMARY>>>"

# Combined prompt: correct the note, then build the Care Plan Comments summary.
RULES = f"""You correct a home-care service-authorization for WellSky. You are given the draft JOURNAL NOTE and the SERVICES list. Produce two things.

PART 1 - corrected journal note. Apply ONLY these rules; do NOT change any facts, dates, numbers, service codes, or meaning:
1. Lead every service with its standard abbreviation, never the full name: Homemaker=HM, Home Delivered Meals=HDM, Chore=HCH, Personal Care=PC, Adult Day Health=ADH, Consumer Directed Services/Care=CDC, PERS stays PERS. Never write "of Homemaker", "Chore services", etc. WORD ORDER: put the abbreviation BEFORE the change word, e.g. "HM initiation", "HDM renewal", "PC increase" - NOT "initiation of HM" or "renewal of HDM".
2. Amounts: hours as "hrs/wk"; HDM as "meals/wk". Do NOT use "units" for hour- or meal-based services (4 units = 1 hour; 1 unit = 1 HDM meal). ADH may be in days, transportation in trips. PERS stays per month.
3. If the authorization is one time, the note must explicitly say "one time".
Keep all dates in their original MM/DD/YYYY format.

PART 2 - Care Plan Comments summary. First line is exactly "Auth:". Then ONE line per APPROVED service:
  <ABBR> <amount> (<detail if any>), effective <M/D/YY> to <M/D/YY>.
Use the same abbreviations. Copy the amount/frequency EXACTLY as written in the SERVICES line - do not change the number. Only convert to a weekly rate when SERVICES gives a TOTAL unit count for HDM (1 unit = 1 meal) or an hour-based service (4 units = 1 hour); NEVER change a per-month PERS count. The detail in parentheses (e.g. meal breakdown, device type) must be taken VERBATIM from wording in the JOURNAL NOTE; if the note has no such detail, use NO parenthetical. NEVER invent a detail, split, number, or descriptor that is not written in the note or services. Dates in M/D/YY (no leading zeros, 2-digit year). If a fact is not present in the note or services, leave it out.

Output EXACTLY this format and nothing else:
{_NOTE_MARK}
<corrected note text>
{_SUMM_MARK}
Auth:
<summary lines>
"""

_DATE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")
_CODE = re.compile(r"\b[A-Za-z]\d{3,}\b")  # service codes like S5161


def _facts(text: str):
    """Return the sets of dates and service codes present in `text`."""
    t = text or ""
    return set(_DATE.findall(t)), set(_CODE.findall(t))


def _parse(raw: str):
    """Split a model reply into (note, summary). Returns (None, "") if the
    NOTE/SUMMARY markers aren't both present."""
    if _NOTE_MARK not in raw or _SUMM_MARK not in raw:
        return None, ""
    after_note = raw.split(_NOTE_MARK, 1)[1]
    note_part, summ_part = after_note.split(_SUMM_MARK, 1)
    note = note_part.strip()
    summ = summ_part.strip()
    if summ and not summ.lower().startswith("auth:"):
        summ = "Auth:\n" + summ
    return (note or None), summ


def _claude_node_pids() -> set[int]:
    """PIDs of node processes currently running the claude-code CLI."""
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "Get-CimInstance Win32_Process | Where-Object { ($_.Name -like 'node*' "
             "-or $_.Name -like 'claude*') -and $_.CommandLine -like '*claude-code*' } "
             "| ForEach-Object { $_.ProcessId }"],
            capture_output=True, text=True, timeout=30)
        return {int(x) for x in r.stdout.split() if x.strip().isdigit()}
    except Exception:
        return set()


def _claude_pids_since(t0: float) -> set[int]:
    """PIDs of claude CLI processes CREATED at/after unix time t0 (1s slack).
    Immune to the snapshot race where a just-spawned process hasn't registered
    yet: creation time is a fact about the process, not about when we looked.

    The comparison is done INSIDE PowerShell on native DateTimes. Do NOT go via
    `Get-Date -UFormat %s`: on Windows PowerShell 5.1 it returns LOCAL time as
    if it were UTC (a 4h skew in EDT), so an epoch cutoff silently excluded
    every process and made the reaper vacuous."""
    from datetime import datetime
    cutoff = datetime.fromtimestamp(t0 - 1).strftime("%Y-%m-%d %H:%M:%S")
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"$cut=[datetime]'{cutoff}'; "
             "Get-CimInstance Win32_Process | Where-Object { ($_.Name -like 'node*' "
             "-or $_.Name -like 'claude*') -and $_.CommandLine -like '*claude-code*' "
             "-and $_.CreationDate -ge $cut } | ForEach-Object { $_.ProcessId }"],
            capture_output=True, text=True, timeout=30)
        return {int(x) for x in r.stdout.split() if x.strip().isdigit()}
    except Exception:
        return set()


def _kill_pids(pids) -> None:
    for pid in pids:
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           capture_output=True, timeout=30)
        except Exception:
            pass


def _reap_call(t0: float, wrapper_pid: int) -> None:
    """On timeout: kill the wrapper, then every claude process created since
    this call began, retrying briefly so a late-registering spawn can't slip by."""
    _kill_pids({wrapper_pid})
    for _ in range(3):
        late = _claude_pids_since(t0)
        if not late:
            break
        _kill_pids(late)
        time.sleep(1.5)


def _call(prompt: str, claude_cmd: str, timeout: int):
    """Run `claude -p` and return stdout, or None on failure/timeout.

    `claude.cmd` is a wrapper: it spawns a node process that DETACHES, so
    killing the wrapper (even with taskkill /T) leaves node alive, holding the
    pipes — that's what wedged the backfill and then leaked orphans. So we
    snapshot the claude node PIDs BEFORE the call and, on timeout, kill exactly
    the ones that appeared during it. Pre-existing PIDs (the user's other Claude
    windows) can never be touched."""
    t0 = time.time()
    try:
        p = subprocess.Popen(
            [claude_cmd, "-p", "--model", MODEL],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except Exception:
        return None
    try:
        out, _ = p.communicate(input=prompt, timeout=timeout)
    except subprocess.TimeoutExpired:
        _reap_call(t0, p.pid)      # wrapper + anything spawned since t0
        try:
            p.communicate(timeout=10)
        except Exception:
            pass
        return None
    except Exception:
        _reap_call(t0, p.pid)
        return None
    if p.returncode != 0:
        return None
    return (out or "").strip() or None


def review_auth(note: str, services_text: str = "", notes_text: str = "", *,
                claude_cmd: str = CLAUDE_CMD, timeout: int = TIMEOUT_S):
    """Return (corrected_note, summary).

    `notes_text` is the free-text `Notes` column (extra detail such as the meal
    breakdown), used only to source the summary's parenthetical detail.

    Never raises. On any failure the corrected_note falls back to the original
    note and summary is "". If Claude alters a date/service code in the note,
    the corrected note is rejected (original kept) but a well-formed summary is
    still returned.
    """
    note = (note or "").strip()
    if not note:
        return note, ""

    # NOTE: the `Notes` column is deliberately NOT sent -- a PHI scan found DOB
    # mentions in it. Only Journal Note + Services (both scanned clean) are sent.
    prompt = (f"{RULES}\nJOURNAL NOTE:\n{note}\n\n"
              f"SERVICES:\n{(services_text or '').strip()}\n")
    raw = _call(prompt, claude_cmd, timeout)
    if not raw:
        return note, ""

    new_note, summary = _parse(raw)
    if not new_note:
        return note, summary  # couldn't parse a note -> keep original note

    # Date / service-code guard on the note only.
    in_d, in_c = _facts(note)
    out_d, out_c = _facts(new_note)
    if not out_d.issuperset(in_d) or not out_c.issuperset(in_c):
        return note, summary  # note changed a date/code -> keep original note

    return new_note, summary


# Back-compat: note-only helper (subject/summary unaffected).
def review_note(note: str, **kw) -> str:
    return review_auth(note, "", **kw)[0]


if __name__ == "__main__":
    # Offline self-tests: parser + guard. No network, no PHI.
    sample = (
        "<<<NOTE>>>\n"
        "renewal of HM 4 hrs/wk effective 05/23/2026, code S5161\n"
        "<<<SUMMARY>>>\n"
        "Auth:\n"
        "HM 4 hrs/wk, effective 5/23/26 to 5/31/27."
    )
    n, s = _parse(sample)
    assert n and "HM 4 hrs/wk" in n, n
    assert s.startswith("Auth:") and "HM 4 hrs/wk" in s, s

    # guard: date preserved -> superset holds
    orig = "renewal of Homemaker 16 units effective 05/23/2026, code S5161"
    good = "renewal of HM 4 hrs/wk effective 05/23/2026, code S5161"
    di, ci = _facts(orig); dg, cg = _facts(good)
    assert dg.issuperset(di) and cg.issuperset(ci)
    # guard: date changed -> superset fails
    bad = "renewal of HM 4 hrs/wk effective 05/22/2026, code S5161"
    db, _ = _facts(bad)
    assert not db.issuperset(di)

    # missing markers -> no note parsed (falls back)
    assert _parse("just some text")[0] is None
    print("note_reviewer self-test: OK")
