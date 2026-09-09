"""test_compose.py -- proves the PHI guard and checks the A304117409 decision.

    py -3.13 test_compose.py          # guard tests (offline) + one live Claude call

The FAKE identifiers below are made up (not a real member). They exist only so
the test can assert they never reach the prompt.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import compose as C

FAKE = {  # A304117409 as mini's literal extract would produce it, with FAKE identifiers
    "member_name": "Testperson, Zelda",
    "health_plan_id": "1300000099",
    "medicaid_id": "100099999999",
    "member_dob": "1941-03-09",
    "address": {"street": "12 Example Street", "city": "Boston", "state": "MA", "zip": "02118"},
    "authorization_number": "A304117409",
    "provider_name": "Central Boston Elder Services",
    "review_date": "2026-09-05",
    "auth_period_start": "2026-01-01",
    "auth_period_end": "2026-10-03",
    "document_header_labels": ["LTSS HCBS Letter"],
    "notification_notes_verbatim": ("Transition of Program: HMK/COMP/PC to PCA\n"
                                    "Current Auth End Dated: 10/3/2026\n"
                                    "Member Zelda Testperson DOB 03/09/1941, call 617-555-0100, "
                                    "ID 100099999999, 12 Example Street"),
    "services": [
        {"description_verbatim": "Personal care services, per 15 minutes, not for an inpatient or resident of a hospital, nursing facility, ICF/MR or IMD, part of the individualized plan of treatment",
         "service_code": "T1019", "modifier": "UB", "amount_verbatim": "2080 Units",
         "from_date": "2026-01-01", "to_date": "2026-10-03", "status_verbatim": None, "denial_reason_verbatim": None},
        {"description_verbatim": "Personal care services, per 15 minutes, not for an inpatient or resident of a hospital, nursing facility, ICF/MR or IMD, part of the individualized plan of treatment",
         "service_code": "T1019", "modifier": "U2", "amount_verbatim": "832 Units",
         "from_date": "2026-01-01", "to_date": "2026-10-03", "status_verbatim": None, "denial_reason_verbatim": None},
    ],
    "overall_decision_verbatim": None,
    "fax_header": {"received_datetime": "09-05-2026 8:47 AM", "pages": "pg 3 of 3"},
    "change_type": None, "journal_note": None,
}
IDENT_VALUES = ["Testperson", "Zelda", "1300000099", "100099999999", "1941-03-09", "03/09/1941",
                "Example Street", "617-555-0100", "A304117409", "02118"]


def test_guard():
    payload, red = C.build_payload(FAKE)
    text = json.dumps(payload)
    for k in C.IDENTIFIER_KEYS:
        assert k not in payload, f"identifier KEY leaked: {k}"
    for v in IDENT_VALUES:
        assert v.lower() not in text.lower(), f"identifier VALUE leaked: {v}"
    assert "the member" in payload["notification_notes_verbatim"]
    # Outcome-based: every identifier shape present in the notes must have fired
    # its own rule (audit counts), and nothing may remain (asserted above).
    assert {"name", "dob", "phone", "long_id", "address"} <= set(red), red
    assert sum(red.values()) >= 5, red
    # the transition facts must SURVIVE redaction (they are what Claude needs)
    assert "Transition of Program" in payload["notification_notes_verbatim"]
    assert "10/3/2026" in payload["notification_notes_verbatim"]
    print("guard: OK  redactions =", red)


def test_guard_is_active():
    saved = C.ALLOWED_KEYS
    try:
        C.ALLOWED_KEYS = frozenset(saved | {"member_name"})   # simulate a bad edit
        try:
            C.build_payload(FAKE)
        except ValueError as e:
            print("guard trips on a bad allowlist: OK ->", str(e)[:70])
            return
        raise AssertionError("guard did NOT trip when an identifier key was allowlisted")
    finally:
        C.ALLOWED_KEYS = saved


def test_render():
    s = C.render_services(FAKE)
    assert "T1019 [UB]" in s and "2080 Units" in s and "2026-10-03" in s, s
    print("render_services: OK")


def test_lint():
    """The team's 2026-09-09 corrections as a fixed test: every BAD output they
    flagged must be caught, every GOOD version they wrote must pass."""
    P = {"services": [{"description_verbatim": "Emergency response system; cellular"}],
         "notification_notes_verbatim": "Redistribution of PERS unit type from Landline to Cellular"}
    bad = [
        ("PERS per month", "Renewal",
         "Authorization received for PERS renewal 1 per month, effective 09/04/2026 to 08/31/2027.",
         "Auth:\nPERS 1 per month (Cellular Network), effective 9/4/26 to 8/31/27.", P),
        ("codes in summary", "Renewal",
         "Authorization received for CDC renewal 8.75 hrs/wk, effective 12/01/2026 to 08/31/2027.",
         "Auth:\nCase management 2/month (T2022 U1), effective 12/1/26 to 8/31/27.\nCDC 8.75 hrs/wk (T1019 U1), effective 12/1/26 to 8/31/27.", {}),
        ("zero qualifiers", "Initiate",
         "Authorization received for HM initiate 3 hrs/wk, effective 09/01/2026 to 09/30/2027.",
         "Auth:\nHM 3 hrs/wk (weekday hours 3, weekend 0, night 0), effective 9/1/26 to 9/30/27.", {}),
        ("one-time multi-line", "Increase",
         "Auth received for an additional PC one time increase of 5 hrs for 09/04/2026.",
         "Auth:\nPC 7 hrs/wk (weekday), effective 9/4/26 to 1/31/27.\nPC one time increase 5 hrs, effective 9/4/26.", {}),
        ("special instructions dropped", "Renewal",
         "Authorization received for HDM renewal 7 meals/wk, effective 09/01/2026 to 07/31/2027.",
         "Auth:\nHDM 7 meals/wk, effective 9/1/26 to 7/31/27.",
         {"notification_notes_verbatim": "MassHealth reinstated as of 9/1/2026"}),
        ("ADH transportation dropped", "Initiate",
         "Authorization received for ADH initiate basic level, 5 days/wk, effective 08/31/2026 to 08/31/2027.",
         "Auth:\nADH basic 5 days/wk, effective 8/31/26 to 8/31/27.",
         {"services": [{"description_verbatim": "Nonemergency transportation"}]}),
    ]
    for name, ct, note, summ, payload in bad:
        assert C.lint(ct, note, summ, payload), f"lint MISSED: {name}"
    good = [
        ("Renewal", "Authorization received via UHC e-fax 617-275-4711 for PERS cellular renewal 12 units, effective 09/04/2026 to 08/31/2027 with Central Boston Elder Services. Special instructions: Redistribution of PERS unit type from Landline to Cellular.",
         "Auth:\nPERS cellular 12 units, effective 9/4/26 to 8/31/27.", P),
        ("Renewal", "Authorization received via UHC e-fax 617-275-4711 for CDC renewal 8.75 hrs/wk, effective 12/01/2026 to 08/31/2027 with Central Boston Elder Services.",
         "Auth:\nCDC 8.75 hrs/wk, effective 12/1/26 to 8/31/27.", {}),
        ("Increase", "Auth received via UHC efax 617-275-4711 for an additional PC one time increase of 5 hrs for 09/04/2026.",
         "Auth:\nPC one time increase of 5 hrs, effective 9/4/26.", {}),
        ("Initiate", "Authorization received via UHC e-fax 617-275-4711 for ADH initiate basic level, 5 days/wk with Greater Boston Golden Age Adult Day Health Center with nonemergency transportation 10 trips/wk, effective 08/31/2026 to 08/31/2027 with Central Boston Elder Services.",
         "Auth:\nADH basic 5 days/wk with round trip transportation, effective 8/31/26 to 8/31/27 with Greater Boston Golden Age Adult Day Health Center.",
         {"services": [{"description_verbatim": "Nonemergency transportation"}]}),
        ("Termination", "Authorization received via UHC e-fax 617-275-4711. HDM ended effective 08/31/2026 due to member disenrollment.",
         "Auth:\nHDM ends 8/31/26 (member disenrolled).", {}),
    ]
    # Termination with an amount/period (user 2026-09-09: "35.5 meals/wk ... we only need
    # 'HDM ended effective 08/31/2026 due to member disenrollment'") must be caught.
    assert C.lint("Termination",
                  "Authorization received via UHC e-fax 617-275-4711 for end of HDM 35.5 meals/wk, effective 08/01/2026 to 08/31/2026 with Central Boston Elder Services. HDM ended effective 08/31/2026 due to member disenrollment.",
                  "Auth:\nHDM ends 8/31/26 (member disenrolled).", {}), "lint MISSED: termination with amount"
    for ct, note, summ, payload in good:
        v = C.lint(ct, note, summ, payload)
        assert not v, f"lint FALSE POSITIVE on a team-approved note: {v}"
    print(f"lint: OK ({len(bad)} bad caught, {len(good)} good passed)")


def test_live_decision():
    d = C.compose(FAKE)
    print("sent_fields:", d["sent_fields"])
    print("change_type:", d["change_type"], "| error:", d["error"] or "none")
    print("note:", d["journal_note"])
    print("summary:", d["summary"].replace("\n", " | "))
    assert not d["error"], d["error"]
    assert d["change_type"] == "Termination", d["change_type"]
    assert "10/03/2026" in d["journal_note"] and "PC" in d["journal_note"], d["journal_note"]
    low = (d["journal_note"] + d["summary"]).lower()
    for v in IDENT_VALUES:
        assert v.lower() not in low, f"identifier in output: {v}"
    print("live decision: OK")


if __name__ == "__main__":
    test_guard(); test_guard_is_active(); test_render(); test_lint()
    if "--offline" not in sys.argv:
        test_live_decision()
    print("ALL OK")
