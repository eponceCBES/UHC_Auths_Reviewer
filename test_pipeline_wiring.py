"""test_pipeline_wiring.py -- proves build_field_body is schema-adaptive.

    py -3.13 test_pipeline_wiring.py            # legacy path (offline) + literal path (1 Claude call)
    py -3.13 test_pipeline_wiring.py --offline  # legacy path only

Uses the FAKE (made-up) A304117409 extract from test_compose.py.
"""
import importlib.util
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
spec = importlib.util.spec_from_file_location("uhc_pipeline", HERE / "uhc_pipeline.py")
up = importlib.util.module_from_spec(spec)
spec.loader.exec_module(up)
from test_compose import FAKE, IDENT_VALUES  # noqa: E402

LEGACY = {  # old AI Builder schema, fake data
    "member_name": "Testperson, Zelda", "health_plan_id": "1300000099", "medicaid_id": None,
    "member_dob": "1941-03-09", "address": None, "authorization_number": "A300000001",
    "review_date": "2026-09-01", "auth_period_start": "2026-09-01", "auth_period_end": "2027-08-31",
    "change_type": "Renewal",
    "services": [{"name": "Homemaker", "subcategory": "Day", "approved": True, "start_date": "2026-09-01",
                  "end_date": "2027-08-31", "units_or_frequency": "3.25 hrs/wk", "weekday_hours": None,
                  "weekend_hours": None, "denial_reason": None}],
    "overall_decision": "approved", "notes": "Renewal.",
    "journal_note": "Authorization received via UHC e-fax 617-275-4711 for HM renewal 3.25 hrs/wk (day), effective 09/01/2026 to 08/31/2027 with Central Boston Elder Services.",
}


def test_legacy_path():
    assert not up._is_literal_extract(LEGACY)
    b = up.build_field_body(LEGACY)
    assert b["Title"] == "A300000001" and b["MemberName"] == "Testperson, Zelda"
    assert b["ChangeType"] == "Renewal" and b["JournalNote"].startswith("Authorization received")
    assert b["Services"].startswith("Homemaker (Day): Approved") and up.NOTE_QA_FIELD in b
    assert "CarePlanComments" not in b            # legacy path never calls Claude
    print("legacy path: OK ->", sorted(b))


def test_literal_path():
    assert up._is_literal_extract(FAKE)
    b = up.build_field_body(FAKE)
    # identifiers -> SharePoint (never to Claude)
    assert b["Title"] == "A304117409" and b["MemberName"] == "Testperson, Zelda"
    # dates are noon-UTC instants by design (SP date-only columns render same day in Eastern)
    assert b["MemberDOB"].startswith("1941-03-09") and b["HealthPlanID"] == "1300000099"
    assert b["StreetAddress"] == "12 Example Street" and b["ZipCode"] == "02118"
    assert b["AuthPeriodStart"].startswith("2026-01-01") and b["AuthPeriodEnd"].startswith("2026-10-03")
    # literal, uninterpreted
    assert "T1019 [UB]" in b["Services"] and "2080 Units" in b["Services"]
    assert b["Notes"].startswith("Transition of Program")
    # decided by Claude
    assert b.get("ChangeType") == "Termination", b.get("ChangeType")
    assert "10/03/2026" in b.get("JournalNote", "") and b.get("CarePlanComments", "").startswith("Auth:")
    low = (b["JournalNote"] + b["CarePlanComments"]).lower()
    for v in IDENT_VALUES:
        assert v.lower() not in low, f"identifier in Claude output: {v}"
    assert up.NOTE_QA_FIELD not in b                 # flag column is legacy-only
    print("literal path: OK -> ChangeType =", b["ChangeType"])
    print("  note:", b["JournalNote"][:110] + "…")


if __name__ == "__main__":
    test_legacy_path()
    if "--offline" not in sys.argv:
        test_literal_path()
    print("ALL OK")
