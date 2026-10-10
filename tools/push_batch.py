"""One consumer window per row: journal note -> care-plan comment -> fax PDF
attachment, then the window is closed. Same rules as push_one_with_attachment
(which opened the consumer three times per row; 2026-10-09 timing ~4 min/row).

    py tools/push_batch.py --item-ids 2801,2802 --username <prod upn>

Options: --attachment-only / --care-plan-only (rows already Documented),
--no-care-plan, --no-attachment, --folder. Sandbox testing: --target-consumer
<id> writes every row onto that consumer, --no-mark leaves the SharePoint
status alone, --test-pdf <file> attaches that file instead of the fax.
Prints statuses and per-row seconds only: no names, no note text.
"""
import argparse, importlib.util, os, re, sys, tempfile, time
sys.stdout.reconfigure(line_buffering=True); sys.stderr.reconfigure(line_buffering=True)
from pathlib import Path
R = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(R))
import sp_attachments as SA
_spec = importlib.util.spec_from_file_location("bridge", R / "uhc_wellsky_journal_bridge.py")
BR = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(BR)

ap = argparse.ArgumentParser()
ap.add_argument("--item-ids", required=True)
ap.add_argument("--username", required=True)
ap.add_argument("--password", default=None, help="'-' prompts; default = stored password")
ap.add_argument("--folder", default="Service Planning")
ap.add_argument("--attachment-only", action="store_true")
ap.add_argument("--care-plan-only", action="store_true")
ap.add_argument("--no-care-plan", action="store_true")
ap.add_argument("--no-attachment", action="store_true")
ap.add_argument("--target-consumer", default=None, help="sandbox: write onto this consumer instead of the row's Client ID")
ap.add_argument("--no-mark", action="store_true", help="never change WellSky Documentation Status")
ap.add_argument("--test-pdf", default=None, help="sandbox: attach this file instead of the SharePoint fax")
ap.add_argument("--restart-every", type=int, default=8,
                help="log out and start a fresh browser every N rows (WellSky 'Memory Notice' after ~18 consumers in one session, 2026-10-09). 0 = never.")
a = ap.parse_args()
os.environ.setdefault("UHC_PUSH_SINCE", "2026-09-01")
mark = not a.no_mark and not (a.attachment_only or a.care_plan_only)


class _A:
    pass


pa = _A(); pa.username = a.username; pa.password = None; pa.restart_every = 999; pa.max_restarts = 1; pa.retries = 1
if a.password == "-":
    import getpass
    pa.password = getpass.getpass("WellSky password: ")
elif a.password:
    pa.password = a.password
else:
    pa.password = BR.resolve_password(pa)
if not pa.password:
    print("[stopped] no WellSky password"); sys.exit(2)

creds = BR.uhc.load_azure_creds(); app = BR.uhc._msal_app(creds)


def fresh_graph():
    """Graph session with a current token. Tokens last ~1 h and a batch runs
    for hours (2026-10-09: every row after the 19th failed with 401); MSAL
    caches the token and re-acquires it when it is about to expire, so this is
    cheap to call once per row."""
    return BR.uhc.graph_session(BR.uhc.get_token(app, "https://graph.microsoft.com/.default"))


g = fresh_graph()
site_id = BR.uhc.resolve_site_id(g)
ids = [x.strip() for x in a.item_ids.split(",") if x.strip()]


def load_row(item_id):
    item = g.get(f"{BR.uhc.GRAPH_BASE}/sites/{site_id}/lists/{BR.uhc.SP_LIST_ID}/items/{item_id}?$expand=fields").json()
    f = item.get("fields") or {}
    client_id = BR.clean_client_id(f.get(BR.COL_CLIENT_ID))
    if not client_id or (f.get(BR.COL_LOOKUP_STATUS) or "") != "Matched":
        return None, "not Matched with a Client ID"
    doc = (f.get(BR.COL_DOC_STATUS) or "")
    if (a.attachment_only or a.care_plan_only) and doc != BR.DOC_DONE:
        return None, "not Documented yet (attachment-only / care-plan-only mode)"
    if not (a.attachment_only or a.care_plan_only) and doc == BR.DOC_DONE and not a.target_consumer:
        return None, "already Documented"
    journal = (f.get(BR.COL_JOURNAL) or "").strip()
    if not journal:
        return None, "no journal note"
    return dict(item=item, f=f,
                client_id=BR.clean_client_id(a.target_consumer) if a.target_consumer else client_id,
                journal=journal, subject=BR.build_subject(f), desc=BR.build_attachment_desc(f),
                summary=(f.get(BR.COL_CARE_PLAN) or "").strip()), None


rows = []
for item_id in ids:
    r, why = load_row(item_id)
    if why:
        print(f"[skip] item {item_id}: {why}"); continue
    rows.append((item_id, r))
if not rows:
    print("[stopped] nothing to do"); sys.exit(2)
do_cp = not a.no_care_plan and not a.attachment_only
do_att = not a.no_attachment and not a.care_plan_only
print(f"[plan] {len(rows)} row(s), one window each, one login  (stamp={mark}, care_plan={do_cp}, attachment={do_att})")

sp = None if (a.test_pdf or not do_att) else SA.sp_cert_session()
w = BR.acquire_session(pa); print("[login] ok")
done = failed = 0; t_all = time.time()
skipped_plan = 0
row_secs = []


def progress(n, total):
    """One line per row: counts, average pace, and when the batch should finish."""
    avg = sum(row_secs) / len(row_secs) if row_secs else 0
    left = total - n
    eta = time.strftime("%H:%M", time.localtime(time.time() + avg * left)) if avg else "?"
    print(f"    [progress] {n}/{total} | done={done} failed={failed} plan-skipped={skipped_plan} "
          f"| avg {avg:.0f}s/row | {left} left, ETA ~{eta}")
try:
    BR.warmup(w, rows[0][1]["client_id"])
    since_restart = 0
    for n, (item_id, r) in enumerate(rows, 1):
        t0 = time.time()
        if a.restart_every and since_restart >= a.restart_every:
            print(f"    [session] fresh browser after {since_restart} rows (WellSky memory)")
            try:
                w.close()
            except Exception:  # noqa: BLE001
                pass
            w = BR.acquire_session(pa); BR.warmup(w, r["client_id"]); since_restart = 0
        since_restart += 1
        g = fresh_graph()                                   # Graph token: per row
        if sp is not None:
            sp = SA.sp_cert_session()                       # SharePoint REST token: per row
        print(f"\n[{n}/{len(rows)}] item {item_id} | subject '{r['subject']}' | '{r['desc']}'")
        with tempfile.TemporaryDirectory(prefix="uhc_att_") as tmp:
            try:
                pdf = None
                if a.test_pdf:
                    pdf = a.test_pdf
                elif sp is not None:
                    names = SA.list_attachments(sp, item_id)
                    name = re.sub(r"(\.pdf)+$", ".pdf", names[0]["name"], flags=re.I) if names else None
                    pdf = SA.download_attachment(sp, item_id, names[0]["name"], tmp, as_name=name) if names else None
                    print(f"    attachment downloaded: {os.path.getsize(pdf):,} bytes" if pdf else "    no attachment on the SharePoint item")
                # ONE window for the whole row
                BR.close_all_windows(w)
                w.open_consumer(r["client_id"]); time.sleep(4); BR.ensure_front(w, r["client_id"])
                # 1) journal
                if not (a.attachment_only or a.care_plan_only):
                    BR.stamp(g, site_id, item_id, BR.DOC_INPROGRESS, mark)
                    result = BR.enter_journal(w, r["item"], r["subject"], r["journal"], r["client_id"], save=True, reopen=False)
                    if result != "documented":
                        BR.stamp(g, site_id, item_id, BR.DOC_FAILED, mark)
                        print(f"    journal not confirmed ({result}); row skipped"); failed += 1
                        BR.close_all_windows(w); continue
                    BR.stamp(g, site_id, item_id, BR.DOC_DONE, mark)
                    stamped = "; row stamped Documented" if mark else ""
                    print(f"    journal saved{stamped}  ({time.time()-t0:.0f}s)")
                # 2) care-plan comment (bridge rule), same window
                if do_cp and r["summary"] and BR.needs_care_plan(r["f"]):
                    ad = BR._auth_date_iso(r["summary"]) or BR._auth_date_iso(r["journal"])
                    cp = BR.write_care_plan_comment(w, r["client_id"], r["summary"], ad, save=True, reopen=False)
                    if cp in ("no-date", "no-grid", "no-plan"):
                        skipped_plan += 1
                        print(f"    care plan: SKIPPED + FLAG {cp} (auth date {ad or '?'} matched no plan)  ({time.time()-t0:.0f}s)")
                    else:
                        print(f"    care plan: {cp}  ({time.time()-t0:.0f}s)")
                elif do_cp and r["summary"]:
                    print("    care plan: skipped (laundry auth with no detailed notes; journal note only)")
                # 3) attachment, same window
                if pdf and do_att:
                    BR.ensure_front(w, r["client_id"])
                    w.goto_tab("File", "Attachments", wait_after=5.0)
                    st = w.add_file_attachment(a.folder, r["desc"], pdf, save=True)
                    print(f"    attachment saved: {st.get('file_type')} {st.get('file_size')} bytes  ({time.time()-t0:.0f}s)")
                BR.close_all_windows(w)
                done += 1
                row_secs.append(time.time() - t0)
                print(f"    row done in {time.time()-t0:.0f}s")
                progress(n, len(rows))
            except Exception as e:  # noqa: BLE001
                failed += 1
                print(f"    FAILED: {type(e).__name__}: {str(e)[:140]}")
                if not (a.attachment_only or a.care_plan_only):
                    try:
                        BR.stamp(g, site_id, item_id, BR.DOC_FAILED, mark)
                    except Exception:  # noqa: BLE001
                        pass
                try:
                    BR.recover(w)
                except Exception:  # noqa: BLE001
                    pass
                progress(n, len(rows))
finally:
    time.sleep(2)
    try:
        w.close()
    except Exception:  # noqa: BLE001
        pass
print(f"\n[summary] done={done} failed={failed} plan-skipped={skipped_plan} in {(time.time()-t_all)/60:.1f} min")
