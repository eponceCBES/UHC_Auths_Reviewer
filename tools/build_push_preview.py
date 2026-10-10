"""Excel preview of exactly what the push will write to WellSky for the rows
added to the UHC list in a date window (created date, inclusive).

    py -3.13 tools/build_push_preview.py 2026-10-01 2026-10-08

One row per SharePoint item the bridge would take (Matched, note, Client ID,
not Documented, not a Member Letter). Columns = what gets written: Journal
Type, Subject, Journal Note, Care Plan Comment (or why it is skipped),
Attachment Description, plus the auth facts and a link to the item.
Output: ../uhc_reviewer_deliverables/reports/UHC_Push_Preview_<from>_<to>.xlsx
"""
import importlib.util, sys
from pathlib import Path
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill

R = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(R))
spec = importlib.util.spec_from_file_location("bridge", R / "uhc_wellsky_journal_bridge.py")
BR = importlib.util.module_from_spec(spec); spec.loader.exec_module(BR)
U = BR.uhc

d_from, d_to = sys.argv[1], sys.argv[2]
creds = U.load_azure_creds(); app = U._msal_app(creds)
g = U.graph_session(U.get_token(app, "https://graph.microsoft.com/.default"))
site = U.resolve_site_id(g)
items = U.list_items(g, site, filter_=f"fields/{BR.COL_LOOKUP_STATUS} eq 'Matched'")
web = "https://centralboston.sharepoint.com/sites/DataManagement/Lists/United%20Health%20Care%20Authorizations"

rows = []
for it in items:
    created = (it.get("createdDateTime") or "")[:10]
    if not (d_from <= created <= d_to):
        continue
    f = it["fields"]
    ready = BR.is_ready(f, it["id"])
    summ = (f.get(BR.COL_CARE_PLAN) or "").strip()
    if not ready:
        cp = ""
    elif summ and BR.needs_care_plan(f):
        cp = summ
    elif summ:
        cp = "(skipped: laundry auth with no detailed notes, journal note only)"
    else:
        cp = "(none)"
    rows.append({
        "Received": created, "SP Item": it["id"], "Auth #": f.get("Title"),
        "Client ID": f.get("ClientID"), "Member": f.get("MemberName"),
        "GSSC": f.get("PrimaryCareManager") or "(unassigned)",
        "Change Type": f.get("ChangeType"),
        "Auth Start": (f.get("AuthPeriodStart") or "")[:10], "Auth End": (f.get("AuthPeriodEnd") or "")[:10],
        "Will Push": "Yes" if ready else ("Documented" if (f.get(BR.COL_DOC_STATUS) or "") == BR.DOC_DONE
                                           else "No: " + ("no journal note yet" if not (f.get(BR.COL_JOURNAL) or "").strip()
                                                          else "Member Letter" if (f.get(BR.COL_CHANGE_TYPE) or "").lower() == "member letter"
                                                          else "no Client ID" if not str(f.get(BR.COL_CLIENT_ID) or "").strip() else "not ready")),
        "Journal Type": BR.journal_type_for(f),
        "Subject": BR.build_subject(f),
        "Journal Note": f.get(BR.COL_JOURNAL),
        "Care Plan Comment": cp,
        "Attachment Description": BR.build_attachment_desc(f),
        "Attachment Desc Source": "list" if (f.get(BR.COL_ATTACH_DESC) or "").strip() else "built",
        "Services": f.get("Services"),
        "SharePoint Item": f"{web}/DispForm.aspx?ID={it['id']}",
    })

df = pd.DataFrame(rows).sort_values(["Received", "SP Item"])
out = R.parent / "uhc_reviewer_deliverables" / "reports" / f"UHC_Push_Preview_{d_from}_{d_to}.xlsx"
widths = {"Journal Note": 70, "Care Plan Comment": 45, "Services": 40, "Subject": 40, "Attachment Description": 44,
          "Member": 22, "GSSC": 20, "Auth #": 13, "SharePoint Item": 12, "Journal Type": 24, "Will Push": 22}
with pd.ExcelWriter(out, engine="openpyxl") as xw:
    df.to_excel(xw, index=False, sheet_name="Push preview")
    ws = xw.sheets["Push preview"]
    hdr = [c.value for c in ws[1]]; link = hdr.index("SharePoint Item") + 1
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E78")
    for r in range(2, ws.max_row + 1):
        c_ = ws.cell(row=r, column=link); c_.hyperlink = c_.value; c_.value = "Open"
        c_.font = Font(color="0563C1", underline="single")
    for i, h in enumerate(hdr, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = widths.get(h, 13)
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=i).alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
print(f"{len(df)} rows ({(df['Will Push']=='Yes').sum()} will push) -> {out}")
