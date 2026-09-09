# Wiring "Claude decides" into the pipeline

Do these together, in this order. Doing only one breaks the list.

## 1. Pipeline: `uhc_pipeline.py` → `build_field_body(parsed)`
`parsed` is now mini's LITERAL extract (schema in `uhc_extraction_prompt_EXTRACT_ONLY.txt`).

```python
from compose import compose, render_services   # compose.py sits next to the pipeline

def build_field_body(parsed: dict) -> dict:
    body = {}
    # identifiers -> SharePoint as today (never to Claude)
    auth_no = (str(parsed.get("authorization_number") or "")).strip()
    if auth_no: body["Title"] = auth_no
    ... keep the existing TEXT_FIELD_MAP / DATE_FIELD_MAP / ADDRESS_FIELD_MAP code ...
    # renamed keys in the new schema:
    #   overall_decision  -> parsed["overall_decision_verbatim"]
    #   provider          -> parsed["provider_name"]
    body["Services"] = render_services(parsed)           # literal, no interpretation
    body["Notes"]    = (parsed.get("notification_notes_verbatim") or "").strip()

    d = compose(parsed)                                   # Claude decides (PHI-guarded)
    if not d["error"]:
        body[CHANGE_TYPE_FIELD] = d["change_type"]        # Initiate/Renewal/Increase/...
        body["JournalNote"]     = d["journal_note"]
        body["CarePlanComments"] = d["summary"]
    # on error: leave ChangeType/JournalNote/CarePlanComments EMPTY. The hourly
    # review_notes.py already treats "no CarePlanComments" as unreviewed and will
    # retry the row next run -- nothing is lost, nothing half-written.
    return body
```
Log `d["sent_fields"]` and `d["redactions"]` (counts only) with the row id for the audit trail.

## 2. Reviewer job: `review_notes.py`
Keep it running hourly as the safety net / retry path. When the pipeline already
filled `CarePlanComments`, the row is skipped, so there's no double work.
(Optional later: have it call `compose()` for rows that have a `Notes` field but
no note, instead of "correcting" a mini note that no longer exists.)

## 3. Power Automate flow + Power Apps (AI Builder) — LAST, together
Only after 1 is deployed:
- **Parse JSON step:** replace its schema with `uhc_parse_json_schema_EXTRACT_ONLY.json`
  (every value nullable — the "generate from sample" schema fails on nulls).
- **Remove or repoint any flow action that reads dropped fields** — `journal_note`,
  `notes`, `overall_decision`, `services[].name` / `units_or_frequency` / `approved`.
  The pipeline now produces those. The flow only needs to keep writing the full model
  response into the list's `JSONPayload` column (+ the attachment) — that's what
  `parse_payload()` reads.
- **AI Builder:** replace the prompt with `uhc_extraction_prompt_EXTRACT_ONLY.txt` and
  make sure it is published/live, not just saved.
Symptom if the flow's Parse JSON is still the old one: resubmitted runs fail at Parse
JSON, or produce old-format duplicate rows the pipeline flags and skips.
Verify one document end-to-end: the row should show literal Services, the verbatim
notes, and a Claude-written ChangeType / JournalNote / CarePlanComments.

## 4. Roll back
Put the old prompt back in Power Apps and revert `build_field_body`. Rows created
in between keep their literal Services/Notes; the reviewer job fills notes.

## Guard reminders
- `compose.IDENTIFIER_KEYS` are never sent; `build_payload` raises if one is.
- Notification notes are redacted (member name → "the member"; SSN/phone/ID/email/address/DOB stripped).
- `py test_compose.py --offline` runs the guard tests without a Claude call. Run it after ANY edit to compose.py.
