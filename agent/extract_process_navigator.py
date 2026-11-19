"""Process Navigator extract - standalone. Run by the app's "Refresh" button, or directly:

  python agent/extract_process_navigator.py --ids-file ids.json --countries DE,US [--l4 "L4_Plan.xlsx"] [--force]

1. Opens a visible Edge window; the user signs in to SAP themselves (credentials are never seen/stored).
2. Fetches, for each scope item: description, process flows/activities, document links
   (OData API of SAP Process Navigator, latest release).
3. Downloads the SAP Cloud ALM test script (xlsx) for each requested country; falls back to the global (XX),
   then US script when a country has none.
4. Writes agent/cache/scope_items/*.json, agent/cache/test_scripts/*.xlsx, agent/cache/extract_meta.json,
   and live progress to agent/cache/extract_status.json.
The generator only reads this cache; it never talks to SAP.
"""
import argparse
import datetime as dt
import html
import json
import re
import sys
import time
import traceback
from pathlib import Path

CACHE = Path(__file__).resolve().parent / "cache"
SCOPE_DIR = CACHE / "scope_items"
TEST_DIR = CACHE / "test_scripts"
STATUS = CACHE / "extract_status.json"
META = CACHE / "extract_meta.json"
URL = "https://me.sap.com/processnavigator"
SCEN = "EARL_SolS-013"  # SAP Best Practices for SAP S/4HANA Cloud Public Edition


def status(state, message, done=0, total=0, **extra):
    CACHE.mkdir(exist_ok=True)
    d = dict(state=state, message=message, done=done, total=total, updated=dt.datetime.now().isoformat(timespec="seconds"), **extra)
    tmp = STATUS.with_suffix(".tmp")
    tmp.write_text(json.dumps(d), encoding="utf-8")
    tmp.replace(STATUS)


def txt(h):
    h = re.sub(r"(?is)<(script|style).*?</\1>", "", h or "")
    h = re.sub(r"(?i)<br\s*/?>|</(p|div|li|h\d|tr)>", "\n", h)
    h = re.sub(r"<[^>]+>", "", h)
    return re.sub(r"\n\s*\n+", "\n", html.unescape(h)).strip()


JS = r"""
async ([ids, conc]) => {
  const E="https://pr.alm.me.sap.com/ui/earl-pn-ui/v1/odata/v4/EAXService/", C="https://pr.alm.me.sap.com/ui/earl-pn-ui/v1/odata/v4/BCMService/";
  const SCEN="%SCEN%";
  const g=async(u)=>{for(let i=0;i<3;i++){try{const r=await fetch(u,{headers:{Accept:"application/json"},credentials:"include"});if(r.ok)return (await r.json()).value||[];}catch(e){} await new Promise(r=>setTimeout(r,1500));}return null;};
  const q=encodeURIComponent;
  async function one(sid){
    const rec={scopeItemId:sid};
    const s=await g(C+"SolutionProcessForGlobalSearchWithUBCM?$filter="+q(`contains(fuzzy,'${sid}') and lanCode eq 'en-US'`)+"&$top=20");
    if(s===null){rec.error="search failed";return rec;}
    const hits=s.filter(v=>v.externalId===sid&&v.solutionScenarioStableId===SCEN).concat(s.filter(v=>v.externalId===sid));
    if(!hits.length){rec.found=false;return rec;}
    const h=hits[0],sp=h.solutionProcessId;
    Object.assign(rec,{found:true,name:h.name,release:h.targetRelease,scenario:h.solutionScenarioName,solutionProcessId:sp,countries:h.countryCodeText,integration:h.integration,setUpInstructions:h.setUpInstructions});
    const [d,fds]=await Promise.all([
      g(E+"SolutionProcessTranslation?$filter="+q(`externalId eq '${sid}' and solutionScenarioStableId eq '${SCEN}' and lanCode eq 'en-US'`)+"&$top=1"),
      g(E+`SolutionProcess2SolutionProcessFlowDiagram?$filter=solutionProcessID eq ${sp}&$top=50`)]);
    if(d&&d[0]){rec.description_html=d[0].description;rec.license=d[0].license;rec.changeCategory=d[0].changeCategory;}
    const dids=(fds||[]).map(x=>x.solutionProcessFlowDiagramID);
    const pf=[sp,...dids].map(p=>"parentEntityId eq "+p).join(" or ");
    const df=dids.map(p=>"spfde_solutionProcessDiagram_ID eq "+p).join(" or ");
    const nf=dids.map(p=>"ID eq "+p).join(" or ");
    const [docs,acts,names]=await Promise.all([
      g(E+"BomItemURLsWithMultiLanguageTranslation?$filter="+q(`(${pf}) and lanCode eq 'en-US'`)+"&$top=1000"),
      dids.length?g(E+"SolutionActivityTranslation?$filter="+q(`(${df}) and lanCode eq 'en-US'`)+"&$top=3000"):[],
      dids.length?g(E+"SolutionProcessFlowDiagramTranslation?$filter="+q(`(${nf}) and lanCode eq 'en-US'`)+"&$select=ID,name"):[]]);
    rec.documents=(docs||[]).map(x=>({name:x.name,bomType:x.bomType,url:x.url,countryName:x.countryName,country_ID:x.country_ID,accessLevel:x.accessLevel,isValid:x.isValid,parent:x.parentEntityId===sp?"scopeItem":"flow"}));
    const nm={};(names||[]).forEach(n=>nm[n.ID]=n.name);
    rec.flows=dids.map(id=>({diagramId:id,name:nm[id]||null,activities:(acts||[]).filter(a=>a.spfde_solutionProcessDiagram_ID===id).map(a=>({name:a.solutionActivityName,application:a.applicationName,appRole:a.applicationRoleName,businessRole:a.businessRoleName,lob:a.lineOfBusiness,businessArea:a.businessArea}))}));
    if(docs===null||acts===null)rec.partial=true;
    return rec;
  }
  const out=[];let i=0;
  await Promise.all(Array.from({length:conc},async()=>{while(i<ids.length){const s=ids[i++];try{out.push(await one(s));}catch(e){out.push({scopeItemId:s,error:String(e)});}}}));
  return out;
}
""".replace("%SCEN%", SCEN)


def navigator_loaded(page):
    for f in page.frames:
        if "alm.me.sap.com" in f.url and "accounts.sap.com" not in f.url:
            try:
                if f.locator("body *").count() > 20:
                    return True
            except Exception:
                pass
    return False


def l4_ids(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ids = []
    for r in wb.active.iter_rows(values_only=True):
        v = r[4] if len(r) > 4 else None
        if isinstance(v, str) and re.fullmatch(r"[0-9A-Z]{3}", v.strip()) and v.strip() not in ids:
            ids.append(v.strip())
    wb.close()
    return ids


def pick_script(doc_list, country):
    """Return (doc, country_used) for the Cloud ALM xlsx: requested country, else XX, else US."""
    cands = [x for x in doc_list if x.get("name") == "Test script (SAP Cloud ALM)" and x.get("url")]
    for c in (country, "XX", "US"):
        d = next((x for x in cands if x.get("country_ID") == c), None)
        if d:
            return d, c
    return None, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ids-file", help="JSON list of scope item IDs")
    ap.add_argument("--l4", help="Take scope IDs from this L4 plan instead")
    ap.add_argument("--countries", default="", help="Comma-separated country codes, e.g. DE,US")
    ap.add_argument("--force", action="store_true", help="Re-fetch and overwrite cached data")
    a = ap.parse_args()
    ids = json.loads(Path(a.ids_file).read_text()) if a.ids_file else l4_ids(a.l4)
    ids = list(dict.fromkeys(i.strip().upper() for i in ids if i.strip()))
    countries = [c.strip().upper() for c in a.countries.split(",") if c.strip()] or ["US"]
    SCOPE_DIR.mkdir(parents=True, exist_ok=True); TEST_DIR.mkdir(parents=True, exist_ok=True)
    notes = []
    try:
        from playwright.sync_api import sync_playwright
        status("waiting_login", "Sign in to SAP in the Edge window that opened (up to 10 minutes).", 0, len(ids))
        with sync_playwright() as p:
            browser = p.chromium.launch(channel="msedge", headless=False)
            ctx = browser.new_context()
            page = ctx.new_page()
            page.goto(URL)
            for _ in range(300):
                page.wait_for_timeout(2000)
                if "me.sap.com/processnavigator" in page.url and navigator_loaded(page):
                    break
            else:
                raise RuntimeError("Timed out waiting for SAP sign-in.")
            fr = next(f for f in page.frames if "pr.alm.me.sap.com/launchpad" in f.url)
            todo = [i for i in ids if a.force or not (SCOPE_DIR / f"{i}.json").exists()]
            status("fetching", f"Fetching {len(todo)} scope items from Process Navigator", 0, len(ids))
            for k in range(0, len(todo), 12):
                for rec in fr.evaluate(JS, [todo[k:k + 12], 6]):
                    sid = rec["scopeItemId"]
                    if rec.get("error") or rec.get("partial"):
                        notes.append(f"{sid}: incomplete ({rec.get('error', 'partial')}) - will retry on next refresh")
                        continue
                    if rec.get("description_html"):
                        rec["description_text"] = txt(rec.pop("description_html"))
                    (SCOPE_DIR / f"{sid}.json").write_text(json.dumps(rec, indent=1, ensure_ascii=False), encoding="utf-8")
                    if not rec.get("found"):
                        notes.append(f"{sid}: not found in Process Navigator")
                status("fetching", f"Fetched {min(k + 12, len(todo))}/{len(todo)} scope items", len(ids) - len(todo) + min(k + 12, len(todo)), len(ids))
            # test scripts
            jobs = []
            for sid in ids:
                f = SCOPE_DIR / f"{sid}.json"
                if not f.exists():
                    continue
                d = json.loads(f.read_text(encoding="utf-8"))
                if not d.get("found"):
                    continue
                seen = set()
                for c in countries:
                    doc, used = pick_script(d.get("documents", []), c)
                    if not doc:
                        notes.append(f"{sid}: no SAP Cloud ALM test script for {c}, global or US")
                    elif used not in seen:
                        seen.add(used); jobs.append((sid, used, doc["url"]))
            for n, (sid, used, url) in enumerate(jobs, 1):
                dest = TEST_DIR / f"{sid}_{used}.xlsx"
                if dest.exists() and dest.stat().st_size > 1000 and not a.force:
                    continue
                status("downloading", f"Downloading test scripts {n}/{len(jobs)}", n, len(jobs))
                try:
                    r = ctx.request.get(url, timeout=60000)
                    body = r.body()
                    if r.status == 200 and body[:2] == b"PK":
                        dest.write_bytes(body)
                    else:
                        notes.append(f"{sid}_{used}: download failed (HTTP {r.status}, not an Excel file)")
                except Exception as e:
                    notes.append(f"{sid}_{used}: download error {str(e)[:80]}")
            browser.close()
        found = [json.loads(f.read_text(encoding="utf-8")) for f in SCOPE_DIR.glob("[0-9A-Z]*.json")]
        rel = sorted({d.get("release") for d in found if d.get("found") and d.get("release")})
        META.write_text(json.dumps(dict(
            last_refreshed=dt.datetime.now().isoformat(timespec="seconds"), releases=rel,
            scope_items_cached=sum(1 for d in found if d.get("found")), scope_items_not_found=sum(1 for d in found if not d.get("found")),
            test_scripts_cached=len(list(TEST_DIR.glob("*.xlsx"))), countries=countries, requested_ids=len(ids), notes=notes[:200]), indent=1), encoding="utf-8")
        status("done", f"Done. {len(jobs)} test scripts checked, {len(notes)} notes.", len(ids), len(ids), notes=notes[:50])
    except Exception as e:
        status("error", f"{type(e).__name__}: {e}", notes=notes[:50], trace=traceback.format_exc()[-1500:])
        sys.exit(1)


if __name__ == "__main__":
    main()
