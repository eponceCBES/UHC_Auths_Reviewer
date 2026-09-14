r"""Add each consumer's CURRENT WellSky service plan (from HAR) to the laundry
report, as two columns after "Service Plan Comments":

    Current Service Plan (HAR)   the Active care plan: program, dates, care
                                 manager, then one line per active service
                                 allocation (service, units x frequency,
                                 provider, dates)
    Laundry on Plan              Yes / No -- is a laundry service on that plan
    Suspended Services           current HAR_SERVICE_SUSPENSIONS for the consumer
                                 (service, since/until, reason); suspended lines
                                 in the plan column are tagged ** SUSPENDED

Applied to every tab of reports\UHC_Laundry_Auths_by_GSSC.xlsx, matched on
Client ID, saved in place. Rows are PHI: this prints counts only.

"Current" = CARE_PLAN_STATUS = 'Active' (newest start date if several) and
SERVICE_ALLOCATION_STATUS = 'Active' with no end date or an end date in the
future. HAR is a nightly copy of WellSky, so "current" means as of the last
HAR refresh (printed).

    py -3.13 tools/add_service_plan.py [path-to-xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, r"C:\Users\eponce\AiHub\har_powerbi")
import har_pbi  # noqa: E402

REPORT = Path(r"C:\Users\eponce\AiHub\uhc_reviewer_deliverables\reports"
              r"\UHC_Laundry_Auths_by_GSSC.xlsx")
AFTER = "Service Plan Comments"
COL_PLAN = "Current Service Plan (HAR)"
COL_LAUNDRY = "Laundry on Plan"
COL_SUSP = "Suspended Services"
NEW_COLS = (COL_PLAN, COL_LAUNDRY, COL_SUSP)
AGENCY = "Central Boston Elder Services, Inc."

SUSP_SQL = """
SELECT c.CLIENT_ID, s.SERVICE_UUID, s.SERVICE, s.PROVIDER_UUID, s.PROVIDER,
       s.CARE_ENROLLMENT_UUID, s.START_DATE, s.END_DATE, s.SUSPENSION_REASON,
       s.LUPDATE_DATETIME
FROM HAR_SERVICE_SUSPENSIONS s
JOIN HAR_CONSUMERS c ON c.CONSUMER_UUID = s.CONSUMER_UUID
WHERE s.AGENCY = '{agency}'
  AND s.START_DATE <= GETDATE()
  AND (s.END_DATE IS NULL OR s.END_DATE >= CAST(GETDATE() AS date))
  AND c.CLIENT_ID IN ({ids})
"""

SQL = """
SELECT c.CLIENT_ID,
       p.SERVICE_UUID, p.PROVIDER_UUID,
       p.CARE_PLAN_UUID, p.CARE_PROGRAM_NAME, p.CARE_PLAN_STATUS,
       p.CARE_PLAN_START_DATE, p.CARE_PLAN_END_DATE, p.CARE_PLAN_CARE_MANAGER_NAME,
       p.CARE_PLAN_CARE_MANAGER_IS_PRIMARY, p.CARE_PLAN_LUPDATE_DATETIME,
       p.SERVICE, p.SUBSERVICE, p.SERVICE_CATEGORY, p.PROVIDER,
       p.SERVICE_ALLOCATION_STATUS, p.SERVICE_ALLOCATION_START_DATE,
       p.SERVICE_ALLOCATION_END_DATE, p.UNITS_ALLOCATED, p.FREQUENCY,
       p.ALLOCATION_TYPE, p.SERVICE_ALLOCATION_LUPDATE_DATETIME
FROM HAR_SERVICE_PLANS p
JOIN HAR_CONSUMERS c ON c.CONSUMER_UUID = p.CONSUMER_UUID
WHERE p.AGENCY = '{agency}'
  AND p.CARE_PLAN_STATUS = 'Active'
  AND c.CLIENT_ID IN ({ids})
"""


def _dt(v):
    """HAR dates arrive as text like '9/11/2026 5:07:50 PM' (or blank)."""
    return pd.to_datetime(v, errors="coerce", format="mixed")


def _d(v) -> str:
    t = _dt(v)
    return "" if pd.isna(t) else t.strftime("%m/%d/%Y")


def _s(v) -> str:
    """Text of a HAR cell; blanks and NaN become ''."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "null") else t


def _schedule(units, freq, kind) -> str:
    """'2 units biweekly' the way the service plan reads it.

    UNITS_ALLOCATED = units per period; FREQUENCY = how many periods between
    deliveries; ALLOCATION_TYPE = the period (WEEKLY / MONTHLY / ...).
    So units 2, frequency 2, WEEKLY = 2 units every 2 weeks = biweekly.
    """
    try:
        u = float(units)
        units_txt = f"{int(u)} unit{'s' if u != 1 else ''}" if u.is_integer() \
            else f"{u:g} units"
    except (TypeError, ValueError):
        units_txt = f"{_s(units)} units".strip()
    try:
        n = int(float(freq))
    except (TypeError, ValueError):
        n = 1
    kind = _s(kind).upper()
    if kind == "WEEKLY":
        every = {1: "weekly", 2: "biweekly"}.get(n, f"every {n} weeks")
    elif kind == "MONTHLY":
        every = {1: "monthly", 2: "every other month"}.get(n, f"every {n} months")
    elif kind == "DAILY":
        every = {1: "daily"}.get(n, f"every {n} days")
    elif kind == "DURATIONSPECIFIED":
        every = "for the authorization period"
    elif kind == "CAREPLAN":
        every = "per care plan"
    else:
        every = kind.lower() if kind else ""
    return f"{units_txt} {every}".strip()


def _susp_label(s: pd.Series) -> str:
    """'SUSPENDED since 09/02/2026 (Consumer Hospitalized)' or with 'until'."""
    start, end = _d(s.get("START_DATE")), _d(s.get("END_DATE"))
    reason = _s(s.get("SUSPENSION_REASON"))
    txt = f"SUSPENDED since {start}" if start else "SUSPENDED"
    if end:
        txt += f" until {end}"
    if reason:
        txt += f" ({reason})"
    return txt


def _matching_susp(a: pd.Series, susp: pd.DataFrame) -> pd.DataFrame:
    """Suspensions for this allocation: same service, and same provider when
    the suspension names one."""
    if susp.empty:
        return susp
    m = susp[susp["SERVICE_UUID"].astype(str) == str(a.get("SERVICE_UUID"))]
    prov = _s(a.get("PROVIDER_UUID"))
    if prov and not m.empty:
        p = m["PROVIDER_UUID"].map(_s)
        m = m[(p == prov) | (p == "")]
    return m


def _fmt_alloc(a: pd.Series, susp: pd.DataFrame) -> tuple[str, str | None]:
    """(line for the plan column, suspended-service text or None)."""
    svc = _s(a["SERVICE"])
    sub = _s(a.get("SUBSERVICE"))
    if sub and sub.lower() != svc.lower():
        svc = f"{svc} ({sub})"
    how = _schedule(a.get("UNITS_ALLOCATED"), a.get("FREQUENCY"),
                    a.get("ALLOCATION_TYPE"))
    prov = _s(a.get("PROVIDER"))
    start = _d(a.get("SERVICE_ALLOCATION_START_DATE"))
    end = _d(a.get("SERVICE_ALLOCATION_END_DATE"))
    dates = f"{start} to {end}" if end else (f"since {start}" if start else "")
    line = f"  - {svc}: {how}"
    if prov:
        line += f", {prov}"
    if dates:
        line += f" ({dates})"
    hit = _matching_susp(a, susp)
    if hit.empty:
        return line, None
    hit = hit.assign(_start=_dt(hit["START_DATE"])).sort_values("_start")
    label = _susp_label(hit.iloc[-1])
    return f"{line}  ** {label}", f"{svc}: {label}"


def summarize(rows: pd.DataFrame, susp: pd.DataFrame) -> tuple[str, str, str]:
    """(plan text, laundry Yes/No, suspended services) for one client."""
    if rows.empty:
        return "No ACTIVE care plan in HAR", "No", _loose_susp(susp, set())
    rows = rows.copy()
    rows["_start"] = _dt(rows["CARE_PLAN_START_DATE"])
    newest = rows.sort_values("_start", ascending=False)["CARE_PLAN_UUID"].iloc[0]
    plan = rows[rows["CARE_PLAN_UUID"] == newest]
    h = plan.iloc[0]
    cm = str(h.get("CARE_PLAN_CARE_MANAGER_NAME") or "").strip()
    head = (f"{h['CARE_PROGRAM_NAME']} | Active {_d(h['CARE_PLAN_START_DATE'])}"
            f" - {_d(h['CARE_PLAN_END_DATE']) or 'open'}"
            + (f" | CM: {cm}" if cm else ""))

    today = pd.Timestamp.today().normalize()
    end = _dt(plan["SERVICE_ALLOCATION_END_DATE"])
    live = plan[(plan["SERVICE_ALLOCATION_STATUS"].astype(str).str.strip() == "Active")
                & (end.isna() | (end >= today))]
    live = live.drop_duplicates(subset=["SERVICE", "SUBSERVICE", "PROVIDER",
                                        "UNITS_ALLOCATED", "FREQUENCY",
                                        "SERVICE_ALLOCATION_START_DATE"])
    live = live.sort_values(["SERVICE_CATEGORY", "SERVICE"], na_position="last")
    lines, suspended, seen = [], [], set()
    for _, a in live.iterrows():
        line, s_txt = _fmt_alloc(a, susp)
        lines.append(line)
        if s_txt:
            suspended.append(s_txt)
            seen.add(str(a.get("SERVICE_UUID")))
    if not lines:
        lines = ["  - (no active service allocations)"]
    n_susp = len(suspended)
    if n_susp:
        head += f" | {n_susp} SUSPENDED"
    text = "\n".join([head] + lines)
    hay = " ".join(str(v) for v in live[["SERVICE", "SUBSERVICE", "SERVICE_CATEGORY"]]
                   .fillna("").values.ravel()).lower()
    return text, ("Yes" if "laundry" in hay else "No"), \
        _loose_susp(susp, seen, suspended)


def _loose_susp(susp: pd.DataFrame, seen: set, found: list | None = None) -> str:
    """The Suspended Services cell: every suspension already tied to a plan
    line, plus any current suspension on a service that is NOT on the active
    plan (still worth knowing). 'None' when there is nothing."""
    out = list(found or [])
    for _, s in susp.iterrows():
        if str(s.get("SERVICE_UUID")) in seen:
            continue
        out.append(f"{_s(s.get('SERVICE'))}: {_susp_label(s)} [not on active plan]")
    return "\n".join(out) if out else "None"


def main(path: Path) -> int:
    ids = pd.read_excel(path, sheet_name=0)["Client ID"].dropna().astype(int).unique()
    print(f"report: {path.name}  ({len(ids)} distinct Client IDs)")

    sql = SQL.format(agency=AGENCY, ids=",".join(str(i) for i in ids))
    har = har_pbi.query(sql)
    har["CLIENT_ID"] = pd.to_numeric(har["CLIENT_ID"], errors="coerce").astype("Int64")
    fresh = _dt(pd.concat([har["CARE_PLAN_LUPDATE_DATETIME"],
                           har["SERVICE_ALLOCATION_LUPDATE_DATETIME"]])).max()
    print(f"HAR: {len(har):,} plan/allocation rows for {har['CLIENT_ID'].nunique()} "
          f"clients; newest update in HAR {_d(fresh)}")

    susp = har_pbi.query(SUSP_SQL.format(agency=AGENCY,
                                         ids=",".join(str(i) for i in ids)))
    susp["CLIENT_ID"] = pd.to_numeric(susp["CLIENT_ID"], errors="coerce").astype("Int64")
    print(f"HAR: {len(susp):,} current suspensions for "
          f"{susp['CLIENT_ID'].nunique()} clients")

    plans: dict[int, tuple[str, str, str]] = {}
    for cid in ids:
        plans[int(cid)] = summarize(har[har["CLIENT_ID"] == int(cid)],
                                    susp[susp["CLIENT_ID"] == int(cid)])
    n_plan = sum(1 for t, _, _ in plans.values() if not t.startswith("No ACTIVE"))
    n_laun = sum(1 for _, y, _ in plans.values() if y == "Yes")
    n_susp = sum(1 for _, _, s in plans.values() if s != "None")
    n_lsusp = sum(1 for _, _, s in plans.values()
                  if s != "None" and "laundry" in s.lower())
    print(f"clients with an Active plan: {n_plan}/{len(ids)}; laundry on plan: {n_laun}")
    print(f"clients with a current suspension: {n_susp}; laundry suspended: {n_lsusp}")

    wb = load_workbook(path)
    for ws in wb.worksheets:
        hdr = [c.value for c in ws[1]]
        if COL_PLAN in hdr:                       # re-run: replace, don't duplicate
            for name in reversed(NEW_COLS):
                if name in hdr:
                    ws.delete_cols(hdr.index(name) + 1)
                    hdr = [c.value for c in ws[1]]
        if AFTER not in hdr or "Client ID" not in hdr:
            print(f"  skip tab {ws.title!r}: expected columns missing")
            continue
        at = hdr.index(AFTER) + 2
        cid_col = hdr.index("Client ID") + 1
        # openpyxl's insert_cols moves cells but NOT hyperlinks or column
        # widths -- capture both by header and put them back afterwards.
        links = {}
        for c0 in range(at, len(hdr) + 1):
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(r, c0)
                if cell.hyperlink is not None:
                    links[(r, c0)] = cell.hyperlink.target
        widths = {h: ws.column_dimensions[get_column_letter(i + 1)].width
                  for i, h in enumerate(hdr)}
        k = len(NEW_COLS)
        ws.insert_cols(at, k)
        for (r, c0), target in links.items():
            ws.cell(r, c0 + k).hyperlink = target
        for i, h in enumerate(hdr):
            if i + 1 >= at and widths.get(h):
                ws.column_dimensions[get_column_letter(i + 1 + k)].width = widths[h]
        for j, name in enumerate(NEW_COLS):
            h = ws.cell(1, at + j, name)
            h.font = Font(bold=True, color="FFFFFF")
            h.fill = PatternFill("solid", fgColor="1F4E79")
        for r in range(2, ws.max_row + 1):
            cid = ws.cell(r, cid_col).value
            try:
                text, yes, sus = plans.get(int(cid), ("", "", ""))
            except (TypeError, ValueError):
                text, yes, sus = "", "", ""
            ws.cell(r, at, text).alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(r, at + 1, yes).alignment = Alignment(vertical="top",
                                                          horizontal="center")
            c_s = ws.cell(r, at + 2, sus)
            c_s.alignment = Alignment(wrap_text=True, vertical="top")
            if sus and sus != "None":
                c_s.font = Font(bold=True, color="C00000")
        ws.column_dimensions[get_column_letter(at)].width = 70
        ws.column_dimensions[get_column_letter(at + 1)].width = 12
        ws.column_dimensions[get_column_letter(at + 2)].width = 40
        # widths to the right shifted by two: re-apply from the old letters
        ws.auto_filter.ref = ws.dimensions
    wb.save(path)
    print(f"saved -> {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main(Path(sys.argv[1]) if len(sys.argv) > 1 else REPORT))
