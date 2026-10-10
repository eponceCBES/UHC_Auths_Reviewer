"""Journal note into production WellSky, then the auth fax PDF as a File
Attachment, for one or more rows in ONE browser session.

    py tools/push_one_with_attachment.py --item-ids 2169,2170 --username eponce@agingnetwork.com

Password: WELLSKY_PASSWORD env, else the file saved by
`uhc_wellsky_journal_bridge.py --save-password`. Headed browser. Prints
statuses only: no names, no note text, no file names.
Steps: (1) download the SharePoint attachment with the app certificate into a
temp folder; (2) log in once; (3) journal note, row stamped Documented; (3b) care-plan comment appended to the
plan covering the auth date (bridge rule); (4) File Attachments > Add New, folder 'Service Planning', description =
the list's Attachment Description (fallback 'UHC <SVC> Authorization'), save; (5) temp folder deleted, always.
"""
import argparse, importlib.util, os, re, sys, tempfile, time

# Logs must appear as they happen, not when the script ends.
sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)
from pathlib import Path

R = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(R))
import sp_attachments as SA
_spec = importlib.util.spec_from_file_location("bridge", R / "uhc_wellsky_journal_bridge.py")
BR = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(BR)

ap = argparse.ArgumentParser()
ap.add_argument("--item-ids", required=True, help="comma-separated SharePoint item ids, done in order in ONE login")
ap.add_argument("--username", required=True)
ap.add_argument("--folder", default="Service Planning")
ap.add_argument("--attachment-only", action="store_true", help="rows already documented: add the fax attachment only; no journal, no stamp change")
ap.add_argument("--care-plan-only", action="store_true", help="rows already documented: append the care-plan comment only; no journal, no attachment")
ap.add_argument("--no-care-plan", action="store_true", help="skip the Care Plan Comments append for this run")
a = ap.parse_args()
os.environ.setdefault("UHC_PUSH_SINCE", "2026-09-01")

class _A: pass
pa = _A(); pa.username = a.username; pa.password = None; pa.restart_every = 999; pa.max_restarts = 1; pa.retries = 1
pa.password = BR.resolve_password(pa)
if not pa.password:
    print("[stopped] no WellSky password (set WELLSKY_PASSWORD or run the push script once with --save-password)"); sys.exit(2)

creds = BR.uhc.load_azure_creds(); app = BR.uhc._msal_app(creds)
g = BR.uhc.graph_session(BR.uhc.get_token(app, "https://graph.microsoft.com/.default"))
site_id = BR.uhc.resolve_site_id(g)
ids = [x.strip() for x in a.item_ids.split(",") if x.strip()]

def load_row(item_id):
    item = g.get(f"{BR.uhc.GRAPH_BASE}/sites/{site_id}/lists/{BR.uhc.SP_LIST_ID}/items/{item_id}?$expand=fields").json()
    f = item.get("fields") or {}
    client_id = BR.clean_client_id(f.get(BR.COL_CLIENT_ID))
    if not client_id or (f.get(BR.COL_LOOKUP_STATUS) or "") != "Matched":
        return None, "not Matched with a Client ID"
    if (a.attachment_only or a.care_plan_only) and (f.get(BR.COL_DOC_STATUS) or "") != BR.DOC_DONE:
        return None, "not Documented yet (attachment-only / care-plan-only mode)"
    if not (a.attachment_only or a.care_plan_only) and (f.get(BR.COL_DOC_STATUS) or "") == BR.DOC_DONE:
        return None, "already Documented"
    journal = (f.get(BR.COL_JOURNAL) or "").strip()
    if not journal:
        return None, "no journal note"
    svc = "/".join(BR._service_abbrevs(f.get("Services") or "", journal)) or "Service"
    # The list's own Attachment Description (filled by the pipeline since 10/2,
    # e.g. "UHC HM Renewal Authorization 10/1/26 to 12/31/26") wins; a blank
    # column gets the same wording built from the row (bridge.build_attachment_desc).
    desc = BR.build_attachment_desc(f)
    return dict(item=item, f=f, client_id=client_id, journal=journal, subject=BR.build_subject(f),
                desc=desc, summary=(f.get(BR.COL_CARE_PLAN) or "").strip()), None

rows = []
for item_id in ids:
    r, why = load_row(item_id)
    if why:
        print(f"[skip] item {item_id}: {why}"); continue
    rows.append((item_id, r))
if not rows:
    print("[stopped] nothing to do"); sys.exit(2)
print(f"[plan] {len(rows)} row(s) in one login")

sp = SA.sp_cert_session()
w = BR.acquire_session(pa)
print("[login] ok")
done = failed = 0
try:
    BR.warmup(w, rows[0][1]["client_id"])
    for n, (item_id, r) in enumerate(rows, 1):
        print(f"\n[{n}/{len(rows)}] item {item_id} | subject '{r['subject']}' | '{r['desc']}'")
        with tempfile.TemporaryDirectory(prefix="uhc_att_") as tmp:
            try:
                names = [] if a.care_plan_only else SA.list_attachments(sp, item_id)
                name = re.sub(r"(\.pdf)+$", ".pdf", names[0]["name"], flags=re.I) if names else None   # original fax name, one .pdf ending
                pdf = SA.download_attachment(sp, item_id, names[0]["name"], tmp, as_name=name) if names else None
                print(f"    attachment downloaded: {os.path.getsize(pdf):,} bytes" if pdf else "    no attachment on the SharePoint item")
                if a.attachment_only or a.care_plan_only:
                    result = "documented"
                else:
                    BR.stamp(g, site_id, item_id, BR.DOC_INPROGRESS, True)
                    result = BR.enter_journal(w, r["item"], r["subject"], r["journal"], r["client_id"], save=True)
                if result != "documented":
                    BR.stamp(g, site_id, item_id, BR.DOC_FAILED, True)
                    print(f"    journal not confirmed ({result}); attachment skipped"); failed += 1
                    BR.close_all_windows(w); continue
                if not (a.attachment_only or a.care_plan_only):
                    BR.stamp(g, site_id, item_id, BR.DOC_DONE, True)
                    print("    journal saved; row stamped Documented")
                BR.close_all_windows(w)
                # Care Plan Comments: same rule as the bridge. Appended to the plan
                # whose dates cover the auth; skipped + flagged if none matches;
                # laundry auths with no detailed notes get the journal note only.
                if not a.no_care_plan and not a.attachment_only and r["summary"] and BR.needs_care_plan(r["f"]):
                    ad = BR._auth_date_iso(r["summary"]) or BR._auth_date_iso(r["journal"])
                    cp = BR.write_care_plan_comment(w, r["client_id"], r["summary"], ad, save=True)
                    print(f"    care plan: {'SKIPPED + FLAG ' + cp + ' (auth date ' + (ad or '?') + ' matched no plan)' if cp in ('no-date','no-grid','no-plan') else cp}")
                    BR.close_all_windows(w)
                elif not a.attachment_only and r["summary"]:
                    print("    care plan: skipped (laundry auth with no detailed notes; journal note only)")
                if pdf:
                    w.open_consumer(r["client_id"]); time.sleep(4)
                    BR.ensure_front(w, r["client_id"])
                    w.goto_tab("File", "Attachments", wait_after=5.0)
                    st = w.add_file_attachment(a.folder, r["desc"], pdf, save=True)
                    print(f"    attachment saved: {st.get('file_type')} {st.get('file_size')} bytes")
                    BR.close_all_windows(w)
                done += 1
            except Exception as e:  # noqa: BLE001
                failed += 1
                print(f"    FAILED: {type(e).__name__}: {str(e)[:140]}")
                if not (a.attachment_only or a.care_plan_only):   # row already Documented in those modes: leave it
                    try:
                        BR.stamp(g, site_id, item_id, BR.DOC_FAILED, True)
                    except Exception:  # noqa: BLE001
                        pass
                try:
                    BR.recover(w)
                except Exception:  # noqa: BLE001
                    pass
finally:
    time.sleep(3)
    w.close()
print(f"\n[summary] done={done} failed={failed}")
