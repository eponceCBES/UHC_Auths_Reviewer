import sys, requests
sys.path.insert(0, r"C:\Users\eponce\AiHub\uhc_reviewer")
import review_notes as rn, compose as C
tok = rn.graph_token(); H={"Authorization":"Bearer "+tok}
site = requests.get(f"{rn.GRAPH}/sites/{rn.SITE_PATH}", headers=H, timeout=60).json()["id"]
url = f"{rn.GRAPH}/sites/{site}/lists/{rn.LIST_ID}/items?$expand=fields($select=Title,ChangeType,JournalNote,CarePlanComments,Services)&$top=5000"
best = {}
while url:
    d = requests.get(url, headers=H, timeout=120).json()
    for it in d.get("value", []):
        f = it["fields"]; c = (it.get("createdDateTime") or "")[:10]; t = (f.get("Title") or "").strip()
        if not t or not (f.get("JournalNote") or "").strip(): continue
        if c < "2026-09-05": continue
        if t not in best or int(it["id"]) < int(best[t][0]): best[t] = (int(it["id"]), c, f)
    url = d.get("@odata.nextLink")
bad = 0
for t, (iid, c, f) in sorted(best.items(), key=lambda kv: (kv[1][1], kv[0])):
    v = C.lint(f.get("ChangeType") or "", f.get("JournalNote") or "", f.get("CarePlanComments") or "", {"services": f.get("Services") or ""})
    flag = "OK " if not v and (f.get("CarePlanComments") or "").strip() else "!! "
    if flag == "!! ": bad += 1
    print(f"{flag}{c} {t} {f.get('ChangeType')}" + (f"  <- {v or 'NO SUMMARY'}" if flag == "!! " else ""))
    print("   N:", f.get("JournalNote")); print("   P:", (f.get("CarePlanComments") or "").replace("\n", " | "))
print(f"\n{len(best)} auths, {bad} flagged")

