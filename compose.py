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
    # The ADH center on the consumer's current WellSky service plan (a provider
    # name, not an identifier). The pipeline adds it after the consumer match
    # because UHC's ADH auths never print the center (team 2026-10-02).
    "service_plan_adh_center",
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

STEP 1 - CHANGE TYPE. Output exactly one of: Initiate, Renewal, Increase, Decrease, Termination, Suspension, Records Request, Member Letter.
- Member Letter = the MEMBER'S copy of an approval letter, not a service authorization to CBES. Tells (any one is enough): "call the toll-free Member Service number on the back of your member ID card", "cc: Central Boston Elder Services", "we're pleased to tell you", "your provider", "Sincerely, The UnitedHealthcare Team". A provider notice ("Service reimbursement", "non-contracted provider", "call Customer Service") is NOT a member letter. (Team 2026-10-02.)
- "Transition of Program: ... to <other program>" together with "Current Auth End Dated: <date>" = the CURRENT authorization ENDS on that date -> Termination (effective the end date). The units table does not make it an approval/renewal.
- "End date authorization", "Loss of Medicaid Coverage", "deeming period", a service stopping/discontinued/not reauthorized, or a line carrying only an end date = Termination.
- Suspension language ("will be suspended") = Suspension.
- "Request Type: <SVC> Increase - One Time Request" or "additional ... one time increase" = Increase (one-time; see note format).
- Otherwise use the header labels and notes: new service never authorized before = Initiate; continuing an existing service = Renewal; more/less of an existing service = Increase/Decrease.

STEP 2 - AMOUNTS. Convert literal amounts to the note's units: hour-based services (Personal Care, Homemaker, Chore, CDC, Companion) 4 units = 1 hour, shown as hrs/wk; HDM 1 unit = 1 meal, shown as meals/wk; ADH in days; transportation in trips. Modifier UB = weekday/day hours, U2 = weekend hours. When a service line gives a TOTAL unit count, the weekly rate is units ÷ weeks in THAT LINE'S PRINTED from/to period (weeks = days ÷ 7, days from the line's own dates) — NEVER assume 52 weeks or a calendar year, even for a termination. Round to the nearest 0.25 hr. Keep the weekday/weekend split when separate lines (UB/U2) give it, and state the total as their sum.
- SCHEDULE (team 2026-10-02): for hour services the NOTE carries the schedule exactly like the summary does: "HM renewal 3 hrs/wk (weekday)", "PC increase 9.5 hrs/wk (5 weekday, 4.5 weekend)". The schedule is printed as "HMK Weekday Hours: 3", "Homemaker (Day) Hours per week: 3", "Companion (Day) Hours", "Weekend Day Hours", "Night Hours", or as the UB/U2 modifier. "(Day)" and "Weekday" both mean weekday. Never drop it from the note when the extract prints it.
- PERS: the amount is the printed TOTAL units (e.g. "13 units"), never "1 per month". The device TYPE is landline or cellular ONLY, and only when the extract prints that word: "PERS cellular renewal 13 units". Features ("PERS Features: Fall Detection", GPS) are not a type; add them after the units: "PERS renewal 12 units with fall detection". When neither landline nor cellular is printed, the note ends with this sentence: "PERS device type (landline/cellular) not specified on the authorization, GSSC to follow up with SCO United." Never guess the type from a modifier (RR, U8).
- CDC (consumer directed) auths: ONLY the consumer-directed hours line (T1019 U1 / "consumer directed") is the service. Case management T2022, per-diem T1020, T1019 TV and 99509 lines are program components of CDC — never list them, never call them PC, never mention their codes. Write "CDC renewal 8.75 hrs/wk".
- ADH: name the level and the center as printed ("ADH initiate basic level, 5 days/wk with <center name>") and fold any transportation line into the same sentence ("with nonemergency transportation 10 trips/wk"); transportation is part of the ADH auth, not a separate service. When the extract prints NO level (basic/complex) or NO center name, the note ends with: "ADH level and center not specified on the authorization, GSSC to confirm with SCO United." (drop the word that IS printed: "ADH center not specified ..." when the level is there). Never invent a level or a center. EXCEPTION: when the extract carries "service_plan_adh_center" (the center on the consumer's current WellSky service plan), that IS the center: write "with <that center>" in the note and the summary and do not say the center is unspecified.
- Drop zero or empty qualifiers: never write "weekend 0", "night 0", "No Known Food Allergies". "(weekday)" alone is fine when only weekday hours are authorized.
- Word order and phrasing: "PERS cellular renewal 13 units" (device BEFORE the action word); "HDM initiate 7 meals/wk (5 weekday, 2 weekend)" — the total first, the split in parentheses, never "5 meals/wk weekday and 2 meals/wk weekend, total 7"; a weekday/weekend split is ONE service line in the summary ("PC 9.5 hrs/wk (5 weekday, 4.5 weekend), effective ..."), never two lines.

STEP 3 - JOURNAL NOTE. One paragraph. Abbreviations: Homemaker=HM, Home Delivered Meals=HDM, Chore=HCH, Personal Care=PC, Adult Day Health=ADH, Consumer Directed=CDC, PERS stays PERS. Companion is always written "Companion" — never "COMP", even when the fax prints "COMP Renewal" (team 2026-10-02). Abbreviation BEFORE the action word ("HM renewal", "PC increase" - never "renewal of HM"). Never the word "units" for hour/meal services. One-time increases must say "one time". Dates as MM/DD/YYYY. Templates:
- Normal: "Authorization received via UHC e-fax 617-275-4711 for <SVC> <action> <amount>, effective <start> to <end> with Central Boston Elder Services."
- Termination: "Authorization received via UHC e-fax 617-275-4711. <SVC> ended effective <end date> due to <reason from the notes, e.g. member disenrollment / transition to the PCA program / loss of Medicaid coverage>." NO amount, NO hrs/wk or meals/wk, NO start date - a termination only says what ended, when, and why.
- One-time increase: "Auth received via UHC efax 617-275-4711 for an additional <SVC> one time increase of <n> hrs for <date>." (or "... of <n> hrs a week for <start> to <end> (<split>)" when the notes give a range.) If the notes list MORE THAN ONE one-time date, name every date in the one note and state the TOTAL hours across the dates ("3 hrs on 09/03/2026 and 3 hrs on 09/08/2026" -> "one time increase of 6 hrs for 09/03/2026 and 09/08/2026"; summary "PC one time increase of 6 hrs, effective 9/3/26 and 9/8/26.") — never drop a date.
- HDM: keep the meal qualifier when printed: "7 meals/wk (5 lunch weekday, 2 weekend)" — "lunch" stays in both the note and the summary. The cultural meal type ("Type of Cultural Meal: Chinese") and the dietary meal type ("Type of Dietary Meal: Regular") ALWAYS go in the note when printed, right after the split: "HDM renewal 7 meals/wk (5 lunch weekday, 2 weekend), Chinese cultural meal, regular diet, effective ...". Both words, every time, even when both are "Regular" ("regular cultural meal, regular diet"). (Team 2026-10-02.)
- Member Letter: "Member copy of a UHC approval letter received via e-fax 617-275-4711 for <SVC> <amount>, <start> to <end>; not a service authorization to Central Boston Elder Services, GSSC to confirm with SCO United." Nothing else — it is not documented as an auth.
- ADH: the center name comes from the service line or the notes; if either names it, it goes in the note ("with Blue Hill Adult Day Health Center") and the summary ("with Blue Hill ADH").
- Suspension (coverage decision letter): "The coverage decision letter received from UHC via e-fax 617-275-4711 stated that the consumer's <service as printed, e.g. Companion care, adult (IADL/ADL)> will be suspended on <date> due to <reason from the notes>. The consumer can appeal the Plan's decision by <appeal deadline> and contact the case manager to discuss how to re-start services." If NO appeal deadline is printed, the sentence is "The consumer can appeal the Plan's decision and contact the case manager to discuss how to re-start services."
- NEVER output a placeholder anywhere: no "null", "None", "N/A", "undefined", "TBD", no empty "to" or "by". A missing fact means you DROP that clause and keep the sentence grammatical.
- Laundry with NO detailed notes (notification notes empty, or only generic eligibility / appeal boilerplate with nothing specific to this member's service): "Authorization received via UHC e-fax 617-275-4711 for laundry service, effective <start> to <end>, no detailed notes were included in the authorization, GSSC was notified for follow up with SCO United." No amount, no units. (Team rule 2026-09-09.)
Special instructions: when the notification notes carry an instruction that is not a service line (e.g. "MassHealth reinstated as of 9/1/2026", "Redistribution of PERS unit type from Landline to Cellular"), append it to the note as a final sentence: "Special instructions: <the instruction as printed>." Always, for every change type — the team relies on it.
Never include a member name, ID, DOB, or address in the note.

STEP 4 - CARE PLAN SUMMARY. First line exactly "Auth:". Then one line per service: "<ABBR> <amount> (<detail only if literally in the extract>), effective <M/D/YY> to <M/D/YY>." For a Termination write "<ABBR> ends <M/D/YY> (<reason>)."
- Never print HCPCS codes or modifiers (T1019, T2022, U1, UB, U2, TV...) in the summary; say "weekday"/"weekend" instead.
- CDC auth: exactly one line, the CDC hours ("CDC 8.75 hrs/wk, effective ..."). No case management / per diem / TV / 99509 lines.
- One-time increase: the summary is ONLY the one-time line ("PC one time increase of 5 hrs, effective 9/4/26."). Do not restate the existing weekly authorization lines. If the note gives the one-time amount as a weekly rate over a range ("17.5 hrs a week for 08/01/2026 to 08/31/2026"), the summary keeps the rate: "PC one time increase of 17.5 hrs/wk, effective 8/1/26 to 8/31/26." — never drop the /wk.
- PERS: "PERS <landline|cellular> <n> units (<features as printed: fall detection, GPS tracker>), effective ...". The features go in the summary too, every time they are printed (team 2026-10-02). Put special instructions in the note, not in the summary.
- ADH: one line: "ADH <level> <n> days/wk with round trip transportation, effective ... with <center name>."
- Keep details short: never allergies, never zero quantities. HDM carries the cultural and dietary type inside the parentheses when printed: "HDM 7 meals/wk (5 lunch weekday, 2 weekend, Chinese cultural, regular diet), effective ...".
- Hour services carry the schedule: "HM 3 hrs/wk (weekday), effective ...".
- PERS without a printed device type: "PERS 12 units (fall detection, device type not specified), effective ...". ADH without a printed level/center: "ADH 5 days/wk with round trip transportation (level and center not specified), effective ...".
- Laundry with no detailed notes: one line "Laundry service, effective <M/D/YY> to <M/D/YY> (no detailed notes, GSSC notified)."
- Member Letter: one line "Member letter only (<SVC> <amount>, <M/D/YY> to <M/D/YY>), not a CBES authorization."

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
VALID_CT = ("Initiate", "Renewal", "Increase", "Decrease", "Termination", "Suspension", "Records Request",
            "Member Letter")
# The member's copy of an approval letter (not an authorization to CBES).
# Any of these in the notification notes marks it; provider notices say
# "Service reimbursement" / "call Customer Service" instead. (Team 2026-10-02.)
MEMBER_LETTER_RX = re.compile(
    r"member service number|member id card|cc:\s*central boston|we'?re pleased to tell you|"
    r"the unitedhealthcare team", re.I)
HOUR_SVC_RX = re.compile(r"\b(HM|PC|Companion|HCH|CDC)\b[^\n]*?\bhrs/wk", re.I)
SCHEDULE_RX = re.compile(r"\b(weekday|weekend|night)s?\b", re.I)


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

    # Summary must actually say something after "Auth:"
    if summ is not None and not lines:
        v.append("summary is empty (only 'Auth:')")
    # PERS word order: "PERS <device> <action> <n> units" / "PERS <device> <n> units"
    if re.search(r"\bPERS\s+(renewal|initiat\w*|increase|decrease)\s+(landline|cellular)", src, re.I) \
            or re.search(r"\bPERS\s+\d+\s*units?\s*\((landline|cellular)\)", src, re.I):
        v.append('PERS word order (write "PERS cellular renewal 13 units" / "PERS cellular 13 units")')
    # Weekday/weekend split belongs in ONE line's parenthetical, not two lines
    _wd = {m.group(1) for m in re.finditer(r"^(\w+)\b[^\n]*\bweekdays?\b", summ or "", re.M | re.I)}
    _we = {m.group(1) for m in re.finditer(r"^(\w+)\b[^\n]*\bweekends?\b", summ or "", re.M | re.I)}
    if any(a in _we for a in _wd) and len(lines) > 1 and not re.search(
            r"^\w+\b[^\n]*\bweekdays?\b[^\n]*\bweekends?\b", summ or "", re.M | re.I):
        v.append("weekday and weekend on separate summary lines (one line: total + split in parentheses)")
    # "5 meals/wk weekday and 2 meals/wk weekend, total 7 meals/wk" -> "7 meals/wk (5 weekday, 2 weekend)"
    if re.search(r"\btotal\s+\d+(\.\d+)?\s*(meals|hrs)/wk", src, re.I):
        v.append('phrasing: write "<total> meals/wk (<a> weekday, <b> weekend)", not "... total N"')
    # One-time increase given as a weekly rate in the note -> the summary must keep /wk
    if "one time" in low and re.search(r"hrs?\s*(a|per)\s*week|hrs/wk", note or "", re.I) and not re.search(r"hrs/wk", summ or ""):
        v.append("one-time increase is a weekly rate in the note but the summary dropped /wk")
    # Placeholders: a missing fact is dropped, never printed
    if re.search(r"\b(null|none|n/a|undefined|tbd)\b|\b(to|by|on|of|for)\s*[.,]|\(\s*\)", src, re.I):
        v.append("placeholder or empty slot (null/None/N/A/TBD or a dangling 'to'/'by')")
    # Companion is the team's word; "COMP" is the fax's (team 2026-10-02)
    if re.search(r"\bCOMP\s+(renewal|initiat\w*|increase|decrease|one time|\d)", src, re.I) \
            or re.search(r"^COMP\b", summ or "", re.M):
        v.append('"COMP" used (write "Companion")')
    # Typos / doubled words that keep slipping through
    if re.search(r"instrucitions|instrutions|\b(\w+)\s+\1\b|\s{2,}\S", src, re.I):
        v.append("typo, doubled word or double space")
    # Termination: what ended, when, why - never an amount or a period
    if (ct or "").lower() == "termination" and re.search(
            r"hrs/wk|meals/wk|\bunits?\b|days/wk|trips|\beffective\s+\d{1,2}/\d{1,2}/\d{2,4}\s+to\b", note or "", re.I):
        v.append("termination note carries an amount or a period (only: ended effective <date> due to <reason>)")
    # PERS: units + device, never "per month"
    if re.search(r"\bPERS\b", src, re.I):
        if re.search(r"\bper month\b|/\s*month\b|\bmonthly\b", low):
            v.append('PERS written per month (must be total units, e.g. "13 units")')
    # HDM qualifier order: "5 lunch weekday", never "5 weekday lunch"
    if re.search(r"\b(weekday|weekend)s?\s+lunch\b", low):
        v.append('HDM qualifier order: write "5 lunch weekday, 2 weekend"')
    # Always /wk, never /week
    if re.search(r"\b(meals|hrs|hours|days|trips)\s*/\s*week\b", low):
        v.append('"/week" used (write /wk)')
    is_term = (ct or "").lower() in ("termination", "suspension", "member letter")
    # PERS device type: printed -> named; not printed -> never invented, flagged for GSSC.
    if re.search(r"\bPERS\b", src, re.I) and not is_term:
        ex_dev = re.search(r"cellular|landline", ex)
        note_dev = re.search(r"cellular|landline", re.sub(r"\(landline/cellular\)", "", low))
        if ex_dev and not note_dev:
            v.append("PERS device type (cellular/landline) missing")
        if note_dev and not ex_dev:
            v.append("PERS device type is not printed on the auth (never guess it from a modifier)")
        if not ex_dev and not (re.search(r"not specified", low) and "gssc" in low):
            v.append('PERS device type not printed: end the note with "PERS device type (landline/cellular) '
                     'not specified on the authorization, GSSC to follow up with SCO United."')
        for feat in ("fall detection", "gps"):
            if feat in ex and feat not in (note or "").lower():
                v.append(f"PERS feature '{feat}' printed on the auth but missing from the note")
            elif feat in ex and feat not in (summ or "").lower():
                v.append(f"PERS feature '{feat}' missing from the summary (write \"PERS cellular 13 units ({feat})\")")
        if "change pers unit" in ex and "change pers unit" not in (note or "").lower():
            v.append('special instructions missing: "Special instructions: Change PERS Unit."')
    # Schedule (team 2026-10-02): when the auth prints weekday/weekend/night/(Day)
    # hours, the NOTE and the summary both carry the qualifier.
    sched_vals = re.findall(r"(?:weekday|weekend(?: day)?|night|\(day\))\s*hours(?: per week)?:\s*(\d+(?:\.\d+)?)", ex)
    # UB/U2 counts only on an hour-service line, not on CDC program components
    # (99509 U2, T2022 U1, T1020) — a CDC note is one line with no split.
    mod_sched = any((s or {}).get("modifier", "") and str(s["modifier"]).upper() in ("UB", "U2")
                    and str((s or {}).get("service_code") or "").upper() not in ("99509", "T2022", "T1020")
                    for s in payload.get("services") or [])
    sched_printed = any(float(x) > 0 for x in sched_vals) or mod_sched
    if sched_printed and not is_term and not re.search(r"\bCDC\b", src):
        if HOUR_SVC_RX.search(note or "") and not SCHEDULE_RX.search(note or ""):
            v.append('schedule missing from the note (write "HM renewal 3 hrs/wk (weekday)")')
        if HOUR_SVC_RX.search(summ or "") and not SCHEDULE_RX.search(summ or ""):
            v.append('schedule missing from the summary (write "HM 3 hrs/wk (weekday)")')
    # HDM cultural / dietary meal type (team 2026-10-02): printed -> in note and summary.
    if re.search(r"\bHDM\b", src) and not is_term:
        for kind, m in re.findall(r"type of (cultural|dietary) meal:\s*([a-z][a-z-]*)", ex):
            if m in ("n", "na", "none", "null", "nka", "no"):
                continue
            label = "cultural" if kind == "cultural" else "diet"
            if m not in low or label not in low:
                v.append(f'HDM {kind} meal type "{m}" printed on the auth but missing '
                         f'(write "..., {m} cultural meal, regular diet")')
            elif m not in (summ or "").lower():
                v.append(f'HDM {kind} meal type "{m}" missing from the summary')
    # ADH level + center (team 2026-10-02): printed -> named; not printed -> flagged for GSSC.
    if re.search(r"\bADH\b", src) and not is_term:
        ex_level = re.search(r"\b(basic|complex)\b", ex)
        plan_center = str(payload.get("service_plan_adh_center") or "").strip()
        # A center is printed as "... Adult Day Health Center" or named by
        # program: "ADH Basic at FUENTE DE VIDA ADH".
        ex_center = (re.search(r"\bcent(er|re)\b", ex) or bool(plan_center)
                     or re.search(r"\b(at|with)\s+[\w'&.-]+(\s+[\w'&.-]+){0,5}\s+(adh|adult day health|adult day)\b", ex))
        if plan_center:
            first = re.split(r"\s+", plan_center.lower())[0]
            if first not in low or first not in (summ or "").lower():
                v.append(f"ADH center from the WellSky service plan ('{plan_center}') missing from the note or summary")
            if "not specified" in low:
                v.append("ADH center is known from the service plan: do not write 'not specified'")
        if ex_level and ex_level.group(1) not in low:
            v.append(f"ADH level '{ex_level.group(1)}' printed on the auth but missing from the note")
        if not ex_level and not ("level" in low and "not specified" in low):
            v.append('ADH level not printed: end the note with "ADH level and center not specified on the '
                     'authorization, GSSC to confirm with SCO United."')
        if not ex_center and not ("center" in low and "not specified" in low):
            v.append('ADH center not printed: say "center not specified on the authorization, GSSC to confirm"')
        if (not ex_level or not ex_center) and "gssc" not in low:
            v.append("ADH level/center not printed: the note must name GSSC for follow-up")
    # Member letter (team 2026-10-02): the member's copy of an approval is not an auth.
    ml = bool(MEMBER_LETTER_RX.search(payload.get("notification_notes_verbatim") or ""))
    if ml and (ct or "").lower() != "member letter":
        v.append("this is the member's copy of an approval letter: change type must be Member Letter")
    if (ct or "").lower() == "member letter":
        if not ml:
            v.append("Member Letter chosen but the notes carry no member-letter wording")
        if "not a service authorization" not in low:
            v.append('Member Letter note must say "not a service authorization to Central Boston Elder Services"')
    # No HCPCS codes / modifiers in the summary
    if re.search(r"\b[A-Z]\d{4}\b|\b99509\b|\b(U1|UB|U2|TV)\b", summ or ""):
        v.append("HCPCS code or modifier in summary")
    # No zero qualifiers / allergies
    if re.search(r"\b(weekend|night|weekday)\s*(hours\s*)?0\b|allerg", low):
        v.append("zero qualifier or allergy text")
    # "units" for hour/meal services
    if re.search(r"\b(HM|PC|HDM|CDC|HCH|Companion)\b[^.\n]*\b\d+(\.\d+)?\s*units\b", src):
        v.append('"units" used for an hour/meal service')
    # Laundry: never an amount/units; with no detailed notes -> the fixed template
    if re.search(r"\blaundry\b", low) and re.search(r"\d+(\.\d+)?\s*units?\b|/wk", low):
        v.append("laundry note carries units/amount (use the laundry template)")
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
    ex_dates = extract_dates(payload)
    for m in re.findall(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b", note):
        mo, da, yr = m
        yr = ("20" + yr) if len(yr) == 2 else yr
        iso = f"{yr}-{int(mo):02d}-{int(da):02d}"
        if iso not in ex_dates:
            out["error"] = f"note cites a date not in the extract ({iso})"
            out["journal_note"] = ""
            break
    return out


def extract_dates(payload: dict) -> set[str]:
    """Every date the note may legitimately cite, as YYYY-MM-DD: service line
    dates, the auth period, the review date, and dates printed in the notes.
    A notes date with no year ("3 Hours on 10/06") is taken in every year the
    extract spans (period start/end, review date) — a one-time PC increase
    was rejected on 2026-10-02 because its "10/06" carried no year."""
    ex_dates: set[str] = set()
    for s in payload.get("services") or []:
        for k in ("from_date", "to_date"):
            if (s or {}).get(k):
                ex_dates.add(str(s[k])[:10])
    for k in ("auth_period_start", "auth_period_end", "review_date"):
        if payload.get(k):
            ex_dates.add(str(payload[k])[:10])
    notes = payload.get("notification_notes_verbatim") or ""
    for mo, da, yr in re.findall(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b", notes):
        yr = ("20" + yr) if len(yr) == 2 else yr
        ex_dates.add(f"{yr}-{int(mo):02d}-{int(da):02d}")
    for iso in re.findall(r"\b\d{4}-\d{2}-\d{2}\b", notes):
        ex_dates.add(iso)
    years = {d[:4] for d in ex_dates}
    for mo, da in re.findall(r"\b(\d{1,2})/(\d{1,2})\b(?!/)", notes):
        if 1 <= int(mo) <= 12 and 1 <= int(da) <= 31:
            for yr in years:
                ex_dates.add(f"{yr}-{int(mo):02d}-{int(da):02d}")
    return ex_dates


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
