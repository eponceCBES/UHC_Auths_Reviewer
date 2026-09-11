"""Builds the two deliverable workbooks (overwrites in place):
  reports/UHC_Auths_Sep5-9_2026.xlsx        all auths since 9/5 with GSSC + WellSky status
  reports/UHC_Laundry_Auths_by_GSSC.xlsx     laundry auths, one tab per worker

    py tools/build_reports.py
"""
import re, sys, requests, pandas as pd, importlib.util
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
R=Path(r"C:\Users\eponce\AiHub\uhc_reviewer"); sys.path.insert(0,str(R)); sys.path.insert(0, r"C:\Users\eponce\AiHub\WellSky Automation")
spec=importlib.util.spec_from_file_location("bridge", R/"uhc_wellsky_journal_bridge.py"); br=importlib.util.module_from_spec(spec); spec.loader.exec_module(br)
import review_notes as rn
tok=rn.graph_token(); H={"Authorization":f"Bearer {tok}"}
site=requests.get(f"{rn.GRAPH}/sites/{rn.SITE_PATH}",headers=H,timeout=60).json()["id"]
web="https://centralboston.sharepoint.com/sites/DataManagement/Lists/United%20Health%20Care%20Authorizations"
url=f"{rn.GRAPH}/sites/{site}/lists/{rn.LIST_ID}/items?$expand=fields($select=Title,ClientID,MemberName,PrimaryCareManager,ChangeType,JournalNote,CarePlanComments,Services,AuthPeriodStart,AuthPeriodEnd,WellSkyDocumentationStatus)&$top=5000"
best={}
while url:
    d=requests.get(url,headers=H,timeout=120).json()
    for it in d.get("value",[]):
        f=it.get("fields") or {}; t=(f.get("Title") or "").strip(); created=(it.get("createdDateTime") or "")[:10]
        if not t or created < "2026-09-05" or not (f.get("JournalNote") or "").strip(): continue
        if t not in best or int(it["id"]) < int(best[t][0]["id"]): best[t]=(it,f,created)
    url=d.get("@odata.nextLink")
def row(it,f,created):
    return {"Received":created,"Auth #":f.get("Title"),"GSSC":f.get("PrimaryCareManager") or "(unassigned)","Client ID":f.get("ClientID"),"Member":f.get("MemberName"),
        "Change Type":f.get("ChangeType"),"Auth Start":(f.get("AuthPeriodStart") or "")[:10],"Auth End":(f.get("AuthPeriodEnd") or "")[:10],
        "Services":f.get("Services"),"Subject":br.build_subject(f),"Journal Note":f.get("JournalNote"),"Service Plan Comments":f.get("CarePlanComments"),
        "WellSky Status":f.get("WellSkyDocumentationStatus"),"SharePoint Item":f"{web}/DispForm.aspx?ID={it['id']}"}
df=pd.DataFrame([row(*v) for v in best.values()]).sort_values(["Received","Auth #"])
widths={"Journal Note":75,"Service Plan Comments":45,"Services":45,"Subject":38,"Member":22,"GSSC":22,"Auth #":13,"SharePoint Item":15}
def style(ws):
    hdr=[c.value for c in ws[1]]; link=hdr.index("SharePoint Item")+1
    for cell in ws[1]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="1F4E79")
    for r in range(2,ws.max_row+1):
        c_=ws.cell(row=r,column=link); c_.hyperlink=c_.value; c_.value="Open"; c_.font=Font(color="0563C1",underline="single")
    for i,h in enumerate(hdr,1):
        ws.column_dimensions[ws.cell(row=1,column=i).column_letter].width=widths.get(h,13)
        for r in range(2,ws.max_row+1): ws.cell(row=r,column=i).alignment=Alignment(wrap_text=True,vertical="top")
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
rep=Path(r"C:\Users\eponce\AiHub\uhc_reviewer_deliverables\reports")
# 1) full corrected report with GSSC
out1=rep/"UHC_Auths_Sep5-9_2026.xlsx"; df.to_excel(out1,index=False); wb=load_workbook(out1); style(wb.active); wb.save(out1)
# 2) laundry report, one tab per GSSC
ldf=df[df["Journal Note"].str.contains("laundry service",case=False,na=False)].sort_values(["GSSC","Received","Auth #"])
out2=rep/"UHC_Laundry_Auths_by_GSSC.xlsx"
with pd.ExcelWriter(out2,engine="openpyxl") as xw:
    ldf.to_excel(xw,sheet_name="All laundry",index=False)
    for g,sub in ldf.groupby("GSSC",sort=True):
        name=re.sub(r"[\[\]\*\?/\\:]","",str(g))[:31] or "unassigned"
        sub.to_excel(xw,sheet_name=name,index=False)
wb=load_workbook(out2)
for ws in wb.worksheets: style(ws)
wb.save(out2)
print(f"{len(df)} rows -> {out1}"); print(f"{len(ldf)} laundry rows -> {out2}"); print("laundry per GSSC:"); print(ldf.groupby("GSSC").size().to_string())

