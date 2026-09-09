"""compose.py -- "Claude decides": from mini's LITERAL extract, decide the change
type and write the journal note + Care Plan summary.

PHI GUARD (the whole point of this module's design):
  * ALLOWLIST. Only the keys in ALLOWED_KEYS are ever placed in the prompt.
    Identifier keys (IDENTIFIER_KEYS) are removed and their absence is ASSERTED
    before the prompt is built -- build_payload() raises if any slips through.
  * REDACTION of the one free-text field (notification_notes_verbatim):
    the member's name, known from the extract, is replaced deterministically
    (full name and each name token), then DOB / long-ID / SSN / phone / street
    address shapes are stripped by regex.
  * The returned decision carries `sent_fields` and `redactions` (counts only)
    so every call can be audited without logging any text.

Pairs with uhc_extraction_prompt_EXTRACT_ONLY.txt (mini copies; this decides).
Reuses note_reviewer's Claude call, model choice, reaper and date/code guard.

    from compose import compose
    decision = compose(extract_json)      # {'change_type','journal_note','summary',
                                          #  'sent_fields','redactions'}
"""

from __future__ import annotations

import importlib.util
import re
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location("note_reviewer", _HERE / "note_reviewer.py")
nr = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(nr)

# ── PHI allowlist ─────────────────────────────────────────────────────────────
IDENTIFIER_KEYS = frozenset({
    "member_name", "member_dob", "medicaid_id", "health_plan_id", "address",
    "fax_header", "authorization_number",
})
ALLOWED_KEYS = frozenset({
    "provider_name", "review_date", "auth_period_start", "auth_period_end",
    "document_header_labels", "overall_decision_verbatim", "services",
    "notification_notes_verbatim",
})
SERVICE_KEYS = ("description_verbatim", "service_code", "modifier", "amount_verbatim",
                "from_date", "to_date", "status_verbatim", "denial_reason_verbatim")

# Order matters for the audit counts: specific shapes first, then the DOB label
# with a short span, so a phone/ID that follows "DOB ..." is counted as itself.
_RX = [
    ("ssn",     re.compile(r"\b\d{3}-\d{2}-\d{4}\b")),
    ("phone",   re.compile(r"\b\d{3}[-.\s]\d{3}[-.\s]\d{4}\b")),
    ("long_id", re.compile(r"\b\d{8,}\b")),
    ("email",   re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")),
    ("address", re.compile(r"\b\d{1,6}\s+(?:[A-Z][a-z]+\.?\s+){1,3}(?:Street|St|Avenue|Ave|Road|Rd|"
                           r"Drive|Dr|Lane|Ln|Boulevard|Blvd|Court|Ct|Place|Pl|Way|Terrace|Ter)\b")),
    ("dob",     re.compile(r"\b(?:dob|date of birth|d\.o\.b\.?)\b[\s:]*\d{1,4}[/-]\d{1,2}[/-]\d{1,4}", re.I)),
]


def redact(text: str, member_name: str | None) -> tuple[str, dict]:
    """De-identify free text. Returns (clean_text, {rule: count})."""
    counts: dict[str, int] = {}
    t = text or ""
    if member_name:
        full = member_name.strip()
        # "Last, First" -> also try "First Last"
        variants = {full}
        if "," in full:
            last, first = [p.strip() for p in full.split(",", 1)]
            variants.add(f"{first} {last}")
        for v in variants:
            if len(v) >= 3:
                t, n = re.subn(re.escape(v), "the member", t, flags=re.I)
                counts["name"] = counts.get("name", 0) + n
        for tok in re.split(r"[\s,]+", full):
            if len(tok) >= 3:
                t, n = re.subn(rf"\b{re.escape(tok)}\b", "the member", t, flags=re.I)
                counts["name"] = counts.get("name", 0) + n
    for rule, rx in _RX:
        t, n = rx.subn(f"[{rule} removed]", t)
        if n:
            counts[rule] = counts.get(rule, 0) + n
    return t, counts


def build_payload(extract: dict) -> tuple[dict, dict]:
    """The de-identified dict that becomes the prompt. RAISES if an identifier
    key would be included -- this is the guard the unit test exercises."""
    member_name = (extract or {}).get("member_name")
    payload = {}
    for k in ALLOWED_KEYS:
        if k not in (extract or {}):
            continue
        v = extract[k]
        if k == "services":
            v = [{sk: (s or {}).get(sk) for sk in SERVICE_KEYS} for s in (v or [])]
        payload[k] = v
    notes, red = redact(payload.get("notification_notes_verbatim") or "", member_name)
    payload["notification_notes_verbatim"] = notes
    leaked = IDENTIFIER_KEYS & set(payload)
    if leaked:
        raise ValueError(f"PHI guard: identifier keys in prompt payload: {sorted(leaked)}")
    return payload, red


# ── The decision prompt ───────────────────────────────────────────────────────
DECIDE = """You are the reviewer for UnitedHealthcare (SCO United) home-care authorizations at Central Boston Elder Services. You receive a LITERAL extract of one authorization document (service lines exactly as printed, the notification notes verbatim, header labels, dates). Decide what it means and write the journal note. Facts come ONLY from the extract; never invent a service, date, amount, or reason.

STEP 1 - CHANGE TYPE. Output exactly one of: Initiate, Renewal, Increase, Decrease, Termination, Suspension, Records Request.
- "Transition of Program: ... to <other program>" together with "Current Auth End Dated: <date>" = the CURRENT authorization ENDS on that date -> Termination (effective the end date). The units table does not make it an approval/renewal.
- "End date authorization", "Loss of Medicaid Coverage", "deeming period", a service stopping/discontinued/not reauthorized, or a line carrying only an end date = Termination.
- Suspension language ("will be suspended") = Suspension.
- "Request Type: <SVC> Increase - One Time Request" or "additional ... one time increase" = Increase (one-time; see note format).
- Otherwise use the header labels and notes: new service never authorized before = Initiate; continuing an existing service = Renewal; more/less of an existing service = Increase/Decrease.

STEP 2 - AMOUNTS. Convert literal amounts to the note's units: hour-based services (Personal Care, Homemaker, Chore, CDC, Companion) 4 units = 1 hour, shown as hrs/wk; HDM 1 unit = 1 meal, shown as meals/wk; ADH in days; transportation in trips. Modifier UB = weekday/day hours, U2 = weekend hours. When a service line gives a TOTAL unit count, the weekly rate is units ÷ weeks in THAT LINE'S PRINTED from/to period (weeks = days ÷ 7, days from the line's own dates) — NEVER assume 52 weeks or a calendar year, even for a termination. Round to the nearest 0.25 hr. Keep the weekday/weekend split when separate lines (UB/U2) give it, and state the total as their sum.
- PERS: the amount is the printed TOTAL units (e.g. "13 units"), never "1 per month", and the device type (landline / cellular, plus GPS/fall detection if printed) is always named: "PERS cellular renewal 13 units".
- CDC (consumer directed) auths: ONLY the consumer-directed hours line (T1019 U1 / "consumer directed") is the service. Case management T2022, per-diem T1020, T1019 TV and 99509 lines are program components of CDC — never list them, never call them PC, never mention their codes. Write "CDC renewal 8.75 hrs/wk".
- ADH: name the level and the center as printed ("ADH initiate basic level, 5 days/wk with <center name>") and fold any transportation line into the same sentence ("with nonemergency transportation 10 trips/wk"); transportation is part of the ADH auth, not a separate service.
- Drop zero or empty qualifiers: never write "weekend 0", "night 0", "No Known Food Allergies". "(weekday)" alone is fine when only weekday hours are authorized.

STEP 3 - JOURNAL NOTE. One paragraph. Abbreviations: Homemaker=HM, Home Delivered Meals=HDM, Chore=HCH, Personal Care=PC, Adult Day Health=ADH, Consumer Directed=CDC, PERS stays PERS. Abbreviation BEFORE the action word ("HM renewal", "PC increase" - never "renewal of HM"). Never the word "units" for hour/meal services. One-time increases must say "one time". Dates as MM/DD/YYYY. Templates:
- Normal: "Authorization received via UHC e-fax 617-275-4711 for <SVC> <action> <amount>, effective <start> to <end> with Central Boston Elder Services."
- Termination: "Authorization received via UHC e-fax 617-275-4711. <SVC> ended effective <end date> due to <reason from the notes, e.g. member disenrollment / transition to the PCA program / loss of Medicaid coverage>." NO amount, NO hrs/wk or meals/wk, NO start date - a termination only says what ended, when, and why.
- One-time increase: "Auth received via UHC efax 617-275-4711 for an additional <SVC> one time increase of <n> hrs for <date>." (or "... of <n> hrs a week for <start> to <end> (<split>)" when the notes give a range.) If the notes list MORE THAN ONE one-time date, name every date in the one note ("... of 3 hrs for 09/03/2026 and 09/08/2026") — never drop a date.
Special instructions: when the notification notes carry an instruction that is not a service line (e.g. "MassHealth reinstated as of 9/1/2026", "Redistribution of PERS unit type from Landline to Cellular"), append it to the note as a final sentence: "Special instructions: <the instruction as printed>." Always, for every change type — the team relies on it.
Never include a member name, ID, DOB, or address in the note.

STEP 4 - CARE PLAN SUMMARY. First line exactly "Auth:". Then one line per service: "<ABBR> <amount> (<detail only if literally in the extract>), effective <M/D/YY> to <M/D/YY>." For a Termination write "<ABBR> ends <M/D/YY> (<reason>)."
- Never print HCPCS codes or modifiers (T1019, T2022, U1, UB, U2, TV...) in the summary; say "weekday"/"weekend" instead.
- CDC auth: exactly one line, the CDC hours ("CDC 8.75 hrs/wk, effective ..."). No case management / per diem / TV / 99509 lines.
- One-time increase: the summary is ONLY the one-time line ("PC one time increase of 5 hrs, effective 9/4/26."). Do not restate the existing weekly authorization lines.
- PERS: "PERS <landline|cellular> <n> units, effective ...". Put special instructions in the note, not in the summary.
- ADH: one line: "ADH <level> <n> days/wk with round trip transportation, effective ... with <center name>."
- Keep details short: meal type / diet only if printed; never allergies, never zero quantities.

Output EXACTLY this and nothing else:
<<<CHANGE_TYPE>>>
<one word from Step 1>
<<<NOTE>>>
<the journal note>
<<<SUMMARY>>>
Auth:
<summary lines>
"""

_CT = "<<<CHANGE_TYPE>>>"
_NT = "<<<NOTE>>>"
_SM = "<<<SUMMARY>>>"
VALID_CT = ("Initiate", "Renewal", "Increase", "Decrease", "Termination", "Suspension", "Records Request")


def _parse(raw: str) -> tuple[str, str, str]:
    if not (raw and _CT in raw and _NT in raw and _SM in raw):
        return "", "", ""
    ct = raw.split(_CT, 1)[1].split(_NT, 1)[0].strip()
    rest = raw.split(_NT, 1)[1]
    note, summ = rest.split(_SM, 1)
    ct = next((v for v in VALID_CT if v.lower() == ct.lower()), "")
    summ = summ.strip()
    if summ and not summ.lower().startswith("auth:"):
        summ = "Auth:\n" + summ
    return ct, note.strip(), summ


def lint(ct: str, note: str, summ: str, payload: dict) -> list[str]:
    """Deterministic check of the team's rules (2026-09-09 feedback). Returns
    the list of violations; empty = clean. This is what turns the prompt's
    rules into a guarantee: a violating answer is retried, then rejected."""
    v: list[str] = []
    src = " ".join([note or "", summ or ""])
    low = src.lower()
    ex = (_json_dumps(payload)).lower()
    lines = [l.strip() for l in (summ or "").splitlines() if l.strip() and l.strip().lower() != "auth:"]

    # Termination: what ended, when, why - never an amount or a period
    if (ct or "").lower() == "termination" and re.search(
            r"hrs/wk|meals/wk|\bunits?\b|days/wk|trips|\beffective\s+\d{1,2}/\d{1,2}/\d{2,4}\s+to\b", note or "", re.I):
        v.append("termination note carries an amount or a period (only: ended effective <date> due to <reason>)")
    # PERS: units + device, never "per month"
    if re.search(r"\bPERS\b", src, re.I):
        if re.search(r"\bper month\b", low):
            v.append('PERS written "per month" (must be total units)')
        if re.search(r"cellular|landline", ex) and not re.search(r"cellular|landline", low):
            v.append("PERS device type (cellular/landline) missing")
    # No HCPCS codes / modifiers in the summary
    if re.search(r"\b[A-Z]\d{4}\b|\b99509\b|\b(U1|UB|U2|TV)\b", summ or ""):
        v.append("HCPCS code or modifier in summary")
    # No zero qualifiers / allergies
    if re.search(r"\b(weekend|night|weekday)\s*(hours\s*)?0\b|allerg", low):
        v.append("zero qualifier or allergy text")
    # "units" for hour/meal services
    if re.search(r"\b(HM|PC|HDM|CDC|HCH|Companion)\b[^.\n]*\b\d+(\.\d+)?\s*units\b", src):
        v.append('"units" used for an hour/meal service')
    # CDC: one line, no components
    if re.search(r"\bCDC\b", src):
        if re.search(r"case management|per diem|\bT2022\b|\bT1020\b", low):
            v.append("CDC components (case management / per diem) mentioned")
        if len(lines) > 1:
            v.append("CDC summary must be a single line")
    # One-time increase: summary is only the one-time line
    if "one time" in low and len(lines) > 1:
        v.append("one-time increase summary must be a single line")
    # Special instructions carried over
    if re.search(r"reinstated|redistribution", ex) and "special instructions" not in low:
        v.append("special instructions missing from the note")
    # ADH: transportation folded in, level/center named when printed
    if re.search(r"\bADH\b", src) and "transportation" in ex and "transportation" not in low:
        v.append("ADH transportation missing")
    return v


def _json_dumps(o) -> str:
    import json as _j
    return _j.dumps(o, ensure_ascii=False)


def compose(extract: dict, *, timeout: int = nr.TIMEOUT_S) -> dict:
    """Decide + write. Never raises on Claude failure: returns empty strings
    with `error` set, so the caller can leave the row for the next run."""
    import json as _json
    payload, red = build_payload(extract)
    sent = sorted(payload)
    prompt = DECIDE + "\nEXTRACT:\n" + _json.dumps(payload, indent=1, ensure_ascii=False) + "\n"
    raw = nr._call(prompt, nr.CLAUDE_CMD, timeout)
    ct, note, summ = _parse(raw or "")
    out = {"change_type": ct, "journal_note": note, "summary": summ,
           "sent_fields": sent, "redactions": red, "error": ""}
    if not (ct and note):
        out["error"] = "claude call failed or unparseable"
        return out
    # Rule guard: one retry with the violations spelled out, then reject.
    viol = lint(ct, note, summ, payload)
    if viol:
        retry = (prompt + "\nYOUR PREVIOUS ANSWER BROKE THESE RULES - fix them and answer again:\n- "
                 + "\n- ".join(viol) + "\n")
        raw = nr._call(retry, nr.CLAUDE_CMD, timeout)
        ct2, note2, summ2 = _parse(raw or "")
        if ct2 and note2:
            ct, note, summ = ct2, note2, summ2
            out.update(change_type=ct, journal_note=note, summary=summ)
        viol = lint(ct, note, summ, payload)
        out["lint_retry"] = True
    if viol:
        out["error"] = "rule violations after retry: " + "; ".join(viol)
        out["journal_note"] = ""
        out["summary"] = ""
        return out
    # Date guard: every date in the extract's service lines / period that the
    # note cites must be a real extract date (no invented dates).
    ex_dates = set()
    for s in payload.get("services") or []:
        for k in ("from_date", "to_date"):
            if s.get(k):
                ex_dates.add(str(s[k])[:10])
    for k in ("auth_period_start", "auth_period_end", "review_date"):
        if payload.get(k):
            ex_dates.add(str(payload[k])[:10])
    # Dates printed inside the notification notes are legitimate too (one-time
    # increase dates, "MassHealth reinstated as of ...", end dates).
    for mo, da, yr in re.findall(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b",
                                 payload.get("notification_notes_verbatim") or ""):
        yr = ("20" + yr) if len(yr) == 2 else yr
        ex_dates.add(f"{yr}-{int(mo):02d}-{int(da):02d}")
    for iso in re.findall(r"\b\d{4}-\d{2}-\d{2}\b",
                          payload.get("notification_notes_verbatim") or ""):
        ex_dates.add(iso)
    for m in re.findall(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b", note):
        mo, da, yr = m
        yr = ("20" + yr) if len(yr) == 2 else yr
        iso = f"{yr}-{int(mo):02d}-{int(da):02d}"
        if iso not in ex_dates:
            out["error"] = f"note cites a date not in the extract ({iso})"
            out["journal_note"] = ""
            break
    return out


def render_services(extract: dict) -> str:
    """Human-readable Services blob for the SharePoint row, from the literal
    extract (no interpretation)."""
    lines = []
    for s in (extract or {}).get("services") or []:
        s = s or {}
        head = (s.get("description_verbatim") or "Service").split(",")[0].strip()
        mod = f" [{s['modifier']}]" if s.get("modifier") else ""
        code = f" {s['service_code']}" if s.get("service_code") else ""
        bits = [f"{head}{code}{mod}: {s.get('status_verbatim') or '—'}"]
        if s.get("amount_verbatim"):
            bits.append(str(s["amount_verbatim"]))
        if s.get("from_date") or s.get("to_date"):
            bits.append(f"{s.get('from_date') or '?'} → {s.get('to_date') or '?'}")
        line = " — ".join(bits)
        if s.get("denial_reason_verbatim"):
            line += f"\n    Reason: {s['denial_reason_verbatim']}"
        lines.append(line)
    return "\n".join(lines)
