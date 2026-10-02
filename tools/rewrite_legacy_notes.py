"""Rewrite OLD-FORMAT rows (the AI Builder payload from before the EXTRACT-ONLY
prompt) under today's rules, by mapping the legacy payload into the literal
extract shape and running compose.py on it.

    py tools/rewrite_legacy_notes.py --item-ids 2116 2112 ...   [--dry-run]
    py tools/rewrite_legacy_notes.py --ids-file ids.json          [--dry-run]

Rows that are Documented in WellSky are skipped. Literal-extract rows are
skipped (the hourly pipeline owns those). Prints counts and per-row status;
notes are printed only with --dry-run (they carry no identifiers by rule).
Team 2026-10-02: the 52 red rows on the Sept renewal report.
"""
import argparse, json, sys, importlib.util, requests
from pathlib import Path

R = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(R))
import review_notes as rn
import compose as C
spec = importlib.util.spec_from_file_location("uhc_pipeline", R / "uhc_pipeline.py")
P = importlib.util.module_from_spec(spec); spec.loader.exec_module(P)


def legacy_to_literal(p: dict) -> dict:
    """Legacy payload -> the shape compose() expects. Nothing is interpreted:
    service names/subcategories become the description, units_or_frequency the
    amount, and the legacy AI summary + per-service hour lines become the notes."""
    services, lines = [], []
    for s in p.get("services") or []:
        s = s or {}
        name = (s.get("name") or "Service").strip()
        sub = (s.get("subcategory") or "").strip()
        approved = s.get("approved")
        services.append({
            "description_verbatim": f"{name}; {sub}" if sub else name,
            "service_code": None, "modifier": None,
            "amount_verbatim": s.get("units_or_frequency"),
            "from_date": s.get("start_date"), "to_date": s.get("end_date"),
            "status_verbatim": "Approved" if approved is True else ("Denied" if approved is False else None),
            "denial_reason_verbatim": s.get("denial_reason"),
        })
        if s.get("weekday_hours"):
            lines.append(f"{name} Weekday Hours: {s['weekday_hours']}")
        if s.get("weekend_hours"):
            lines.append(f"{name} Weekend Hours: {s['weekend_hours']}")
    ct = (p.get("change_type") or "").strip()
    head = []
    if ct and services:
        head.append(f"Request Type: {services[0]['description_verbatim'].split(';')[0]} {ct}")
    notes = "\n".join(head + lines + ([p["notes"].strip()] if p.get("notes") else []))
    return {
        "member_name": p.get("member_name"),
        "provider_name": "Central Boston Elder Services",
        "review_date": p.get("review_date"),
        "auth_period_start": p.get("auth_period_start"),
        "auth_period_end": p.get("auth_period_end"),
        "document_header_labels": [x for x in [p.get("overall_decision")] if x],
        "overall_decision_verbatim": p.get("overall_decision"),
        "services": services,
        "notification_notes_verbatim": notes or None,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--item-ids", nargs="*", default=[])
    ap.add_argument("--ids-file")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    ids = list(a.item_ids) + (json.load(open(a.ids_file)) if a.ids_file else [])
    if not ids:
        ap.error("no item ids")
    tok = rn.graph_token(); H = {"Authorization": f"Bearer {tok}", "Content-Type": "application/json"}
    site = requests.get(f"{rn.GRAPH}/sites/{rn.SITE_PATH}", headers=H, timeout=60).json()["id"]
    n = {"written": 0, "skipped": 0, "failed": 0}
    for iid in ids:
        f = requests.get(f"{rn.GRAPH}/sites/{site}/lists/{rn.LIST_ID}/items/{iid}"
                         f"?$expand=fields($select=Title,JSONPayload,ClientID,WellSkyDocumentationStatus)",
                         headers=H, timeout=60).json().get("fields") or {}
        p = P.parse_payload(f.get("JSONPayload"))
        if not p or P._is_literal_extract(p):
            print(f"[skip] {iid}: not a legacy payload"); n["skipped"] += 1; continue
        if (f.get("WellSkyDocumentationStatus") or "") == "Documented":
            print(f"[skip] {iid}: already Documented"); n["skipped"] += 1; continue
        ex = legacy_to_literal(p)
        if any("adult day" in (s["description_verbatim"] or "").lower() for s in ex["services"]):
            center = P.plan_adh_center(f.get("ClientID"))
            if center:
                ex["service_plan_adh_center"] = center
        d = C.compose(ex)
        if d.get("error"):
            print(f"[fail] {iid}: {d['error']}"); n["failed"] += 1; continue
        ct = d["change_type"]
        body = {"JournalNote": d["journal_note"], "CarePlanComments": d["summary"],
                P.CHANGE_TYPE_FIELD: P.CHANGE_TYPE_DISPLAY.get(ct.lower(), ct)}
        if a.dry_run:
            print(f"[dry] {iid} {ct}\n   {d['journal_note']}\n   {d['summary'].replace(chr(10), ' | ')}")
            n["written"] += 1; continue
        r = requests.patch(f"{rn.GRAPH}/sites/{site}/lists/{rn.LIST_ID}/items/{iid}/fields",
                           headers=H, json=body, timeout=60)
        if r.ok:
            print(f"[ok] {iid} {ct}"); n["written"] += 1
        else:
            print(f"[fail] {iid}: PATCH {r.status_code}"); n["failed"] += 1
    print(n)
    return 0 if not n["failed"] else 1


if __name__ == "__main__":
    sys.exit(main())
