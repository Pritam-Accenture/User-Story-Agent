"""SAP S/4HANA Public Cloud user story agent - generator.

Usage:
  python agent/build_plan.py --input Agent_Input.xlsx --l4 "L4_Plan_Rev1_Full (49).xlsx" \
         --sscui SSCUI_List_EN_XX.xlsm --out Output/plan.xlsx [--no-llm] [--model <model>]

Reads  : input workbook (Scope_Items, Countries, KDD, Product_Backlog), L4 plan, SSCUI list,
         cached SAP Process Navigator data (agent/cache/scope_items, agent/cache/test_scripts).
Writes : a NEW Excel file in the sample layout (L4 Activity > User Story > Task > Sub-task),
         with a Country column. Owner is left blank. Existing files are never modified.
Claude : uses the signed-in Claude Code login (see llm.py); if not signed in, template wording is used.
"""
import argparse, collections, datetime as dt, html, json, math, os, re, sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))   # so `import llm` works under python -I
CACHE = Path(__file__).resolve().parent / "cache"

# ----------------------------------------------------------------- helpers
def clean_html(h):
    h = re.sub(r"(?is)<(script|style).*?</\1>", "", h or "")
    h = re.sub(r"(?i)<br\s*/?>|</(p|div|li|h\d|tr)>", "\n", h)
    h = re.sub(r"<[^>]+>", "", h)
    return re.sub(r"\n\s*\n+", "\n", html.unescape(h)).strip()

def short(s, n=250):
    s = re.sub(r"\s+", " ", s or "").strip()
    return s if len(s) <= n else s[: n - 1] + "…"

def to_date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    try: return dt.datetime.fromisoformat(str(v)).date()
    except Exception: return None

def sprint_no(v):
    m = re.search(r"\d+", str(v or ""))
    return int(m.group()) if m else None

# ----------------------------------------------------------------- inputs
def load_input(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    scope, override = [], {}
    for r in wb["Scope_Items"].iter_rows(min_row=2, values_only=True):
        sid = str(r[0] or "").strip().upper()
        if not sid or "example" in str(r[2] or "").lower(): continue
        scope.append(sid)
        if r[1]: override[sid] = [c.strip().upper() for c in str(r[1]).split(",") if c.strip()]
    countries = [str(r[0]).strip().upper() for r in wb["Countries"].iter_rows(min_row=2, values_only=True) if r[0]]
    kdd = collections.defaultdict(list)
    if "KDD" in wb.sheetnames:
        for r in wb["KDD"].iter_rows(min_row=2, values_only=True):
            if r[0] and "example" not in str(r[1] or "").lower():
                kdd[str(r[2] or "").strip().upper()].append(dict(id=r[0], title=r[1], decision=r[3], rationale=r[4]))
    backlog = collections.defaultdict(list)
    if "Product_Backlog" in wb.sheetnames:
        for r in wb["Product_Backlog"].iter_rows(min_row=2, values_only=True):
            if r[0]: backlog[str(r[3] or "").strip().upper()].append(dict(id=r[0], title=r[1], desc=r[2], prio=r[4]))
    names = {}
    if "Lookup_Countries" in wb.sheetnames:
        names = {str(r[0]): r[1] for r in wb["Lookup_Countries"].iter_rows(min_row=2, values_only=True) if r[0]}
    return dict(scope=list(dict.fromkeys(scope)), override=override, countries=countries, kdd=kdd,
                backlog=backlog, country_names=names)

def load_l4(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = []
    for r in wb.active.iter_rows(values_only=True):
        if r[0] and re.fullmatch(r"\d+(\.\d+)*", str(r[0]).strip()):
            rows.append(dict(wbs=str(r[0]).strip(), phase=r[1], l2=r[2], l3=r[3], sid=(r[4] or "").strip() or None,
                             task=r[5], type=r[6], start=to_date(r[9]), end=to_date(r[10]), s_sprint=r[11], e_sprint=r[12],
                             purpose=r[27] if len(r) > 27 else None))
    return rows

def load_sscui(path):
    """scope item id -> list of dict(area, sub, name, id, share_count)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_vba=False)
    ws = next(wb[n] for n in wb.sheetnames if re.fullmatch(r"\d{4}", n))
    release = ws.title
    by = collections.defaultdict(list)
    for r in ws.iter_rows(min_row=5, values_only=True):
        if not r[4]: continue
        ids = [x.strip().upper() for x in re.split(r"[,;]", str(r[9] or "")) if x.strip()]
        for i in ids:
            by[i].append(dict(area=r[0], sub=r[1], name=(r[2] or r[4]), item_id=r[3], act_id=r[8],
                              approach=r[6], shared=len(ids)))
    return release, by

def load_scope_cache(sid):
    f = CACHE / "scope_items" / f"{sid}.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else None

def load_script(sid, countries):
    """Return (steps_by_activity, [(label_countries, country_used, fallback_flag)]). One script per resolved country."""
    groups = collections.OrderedDict()
    for c in countries:
        for cand in (c, "XX", "US"):
            f = CACHE / "test_scripts" / f"{sid}_{cand}.xlsx"
            if f.exists():
                groups.setdefault(cand, []).append(c)
                break
        else:
            groups.setdefault(None, []).append(c)
    out = []
    for used, cs in groups.items():
        if used is None:
            out.append(dict(countries=cs, used=None, fallback=True, activities=[]))
            continue
        wb = openpyxl.load_workbook(CACHE / "test_scripts" / f"{sid}_{used}.xlsx", read_only=True, data_only=True)
        rows = list(wb.active.iter_rows(min_row=6, values_only=True))
        acts = collections.OrderedDict()
        for r in rows:
            if len(r) > 19 and r[14]:
                a = acts.setdefault(str(r[14]).strip(), [])
                a.append(dict(action=str(r[18] or ""), text=clean_html(r[19] if len(r) > 19 else ""),
                              expected=clean_html(r[20] if len(r) > 20 else "")))
        out.append(dict(countries=cs, used=used, fallback=any(c != used for c in cs), activities=acts))
    return out

# ----------------------------------------------------------------- story content
FIB = [1, 2, 3, 5, 8, 13]
def points(n_cfg, n_steps):
    x = 1 + n_cfg / 6 + n_steps / 25
    return next((f for f in FIB if f >= x), 13)

def pick_sscui(items, k=10, max_shared=3):
    """Only SSCUIs that name at most `max_shared` scope items are treated as specific to this scope item."""
    items = sorted((x for x in items if x["shared"] <= max_shared), key=lambda x: (x["shared"], str(x["name"])))
    seen, out = set(), []
    for it in items:
        n = str(it["name"]).strip()
        if n.lower() in seen: continue
        seen.add(n.lower()); out.append(it)
        if len(out) >= k: break
    return out, len(items) - len(out)

def llm_story(model, ctx):
    prompt = f"""You are an SAP S/4HANA Public Cloud implementation consultant using SAP Activate.
Write backlog content for ONE scope item. Use only the facts given; do not invent configuration values.

Scope item: {ctx['id']} {ctx['name']} (release {ctx['release']})
Countries in scope: {', '.join(ctx['countries'])}
SAP description: {ctx['desc']}
Process flows: {'; '.join(ctx['flows'])}
Key activities: {'; '.join(ctx['activities'])}
Configuration items (SSCUI): {'; '.join(ctx['sscui'])}
Key design decisions already recorded: {json.dumps(ctx['kdd'])}
Additional project context supplied by the user (workshop notes, BPD, free text): {ctx.get('extra') or 'none'}

Return ONLY JSON with keys:
 user_story  - one sentence "As a <role>, I want to <capability> so that <benefit>"
 acceptance_criteria - 3 to 5 testable "Given/When/Then" strings
 design_questions - 2 to 4 fit-to-standard questions to confirm with the business
 build_summary - one sentence on what is configured"""
    import llm
    return llm.ask_json(prompt, model)

def template_story(ctx):
    return dict(
        user_story=f"As a business user, I want to run {ctx['name']} in standard SAP S/4HANA Public Cloud so that the process is supported end to end.",
        acceptance_criteria=[f"Given the scope item is activated, when the '{f}' flow is executed, then it completes as described in the SAP test script"
                             for f in ctx["flows"][:4]] or ["Given the scope item is activated, when the SAP test script is executed, then all steps pass"],
        design_questions=[f"Does the standard process '{f}' fit your requirement? Capture gaps as KDD or backlog items." for f in ctx["flows"][:3]],
        build_summary=f"Activate and configure {ctx['name']} per SAP Best Practices.")

def llm_reqs(model, ctx, items, extra=""):
    import llm
    lst = json.dumps([dict(id=i["id"], title=i["title"], description=i["description"], category=i["category"],
                           priority=i["priority"], decision=i["decision"]) for i in items], ensure_ascii=False)
    prompt = f"""You are an SAP S/4HANA Public Cloud functional consultant using SAP Activate.
Scope item: {ctx['id']} {ctx['name']}. Standard process flows: {'; '.join(ctx['flows'][:8])}
Project context: {extra[:1500] or 'none'}

Turn each requirement below into a sprint-ready user story. Use only the facts given; do not invent details.
Requirements: {lst}

Return ONLY JSON: {{"stories": [{{"id": <same id>, "user_story": "As a <role>, I want ... so that ...",
 "acceptance_criteria": ["Given ... When ... Then ...", ... 2 to 4 items]}}]}}"""
    d = llm.ask_json(prompt, model)
    return {x["id"]: x for x in d.get("stories", []) if x.get("id")}

def template_reqs(ctx, items):
    out = {}
    for i in items:
        d = short(i["description"] or i["title"], 200)
        out[i["id"]] = dict(id=i["id"], user_story=f"As a business user, I want {i['title'].rstrip('.')} so that the requirement agreed with the business is met.",
                            acceptance_criteria=[f"Given the solution is configured, when the process is executed, then: {d}"])
    return out

WRICEF = {"workflow", "report", "interface", "conversion", "enhancement", "form"}

def add_req_story(rows, sid, sname, it, st, base, parent, story_id, inp):
    def add(**k): rows.append(k)
    cat = (it.get("category") or "").strip()
    is_w = cat.lower() in WRICEF
    st = st or dict(user_story=it["title"], acceptance_criteria=[short(it["description"], 200)] if it["description"] else [])
    origin = {"workshop": "Workshop", "gap": "Gap / WRICEF", "backlog": "Product Backlog"}.get(it.get("kind"), "User input")
    kdd_refs = [k["id"] for k in inp["kdd"].get(sid, [])] if sid else []
    ac = "\n".join(f"{i}. {a}" for i, a in enumerate(st.get("acceptance_criteria", []), 1))
    desc = short(f"{it['description']} " + (f"Priority: {it['priority']}. " if it["priority"] else "") + (f"Decision: {it['decision']}." if it["decision"] else ""), 600)
    trace = f"{it['id']} ({it['source']})" + (f"; KDD: {', '.join(map(str, kdd_refs))}" if kdd_refs else "")
    b = dict(base)
    add(type="User Story", id=story_id, title=st.get("user_story") or it["title"], origin=origin, build=cat or "Configuration", parent=parent,
        country="All", pts=5 if is_w else 3, ac=ac, desc=desc, trace=trace, src=it["source"], lvl=1, req_id=it["id"], req_kind=it.get("kind"), **b)
    steps = ([("Specify and agree the design", ["Write the functional specification for the " + cat.lower(), "Agree the design and acceptance criteria with the process owner", "Record the decision in the KDD register"]),
              ("Develop and unit test", [f"Build the {cat.lower()} in the development tenant", "Peer review and unit test with agreed sample data", "Document the object in the WRICEF list"]),
              ("Test and hand over", ["Write the test case and link it to this story", "Transport to the test tenant", "Demo at sprint review and get acceptance"])]
             if is_w else
             [("Confirm design inputs", ["Confirm the requirement and values with the process owner", "Check the SAP test script and configuration activities", "Record the decision against the requirement in Cloud ALM"]),
              ("Configure in the development tenant", ["Make the agreed configuration changes", "Check dependencies with other scope items"]),
              ("Test and hand over", ["Unit test with agreed sample data and save evidence", "Write the test case and link it to this story", "Transport to the test tenant", "Demo at sprint review and get acceptance"])])
    for ti, (tt, subs) in enumerate(steps, 1):
        tid = f"{story_id}.T{ti}"
        add(type="Task", id=tid, title=tt, origin=origin, build=cat or "Configuration", parent=story_id, country="All", pts=None, desc="", trace=trace, src=it["source"], lvl=2, **b)
        for si, sub in enumerate(subs, 1):
            add(type="Sub-task", id=f"{tid}.S{si}", title=sub, origin=origin, build=cat or "Configuration", parent=tid, country="All", pts=None, desc="", trace=trace, src=it["source"], lvl=3, **b)

# ----------------------------------------------------------------- build rows
COLS = ["Row Type", "ID", "Phase (L1)", "Workstream / Module (L2)", "SAP Area", "L3 Deliverable",
        "L4 Activity / Story / Task / Sub-task", "Story Origin", "Scope Items", "Build Type", "Owner", "Parent ID",
        "Country", "Sprint", "Story Points", "Start Date", "End Date", "Status",
        "Acceptance Criteria", "Description", "Traceability", "Source Reference"]

def build(inp, l4, sscui, use_llm, model, log, progress=None):
    selected = set(inp["scope"])
    # contexts
    meta, ctxs = {}, {}
    for sid in inp["scope"]:
        c = load_scope_cache(sid)
        countries = inp["override"].get(sid, inp["countries"])
        if not c or not c.get("found"):
            log.append(f"{sid}: not found in cached Process Navigator data - story skipped.")
            continue
        avail = [x.strip() for x in (c.get("countries") or "").split(",") if x.strip()]
        if avail and any(x not in avail for x in countries):
            miss = [x for x in countries if x not in avail]
            log.append(f"{sid}: not available for country {', '.join(miss)} per SAP (available: {', '.join(avail) or 'global only'}).")
        flows = [f["name"] for f in c["flows"] if f.get("name")]
        acts = [a["name"] for f in c["flows"] for a in f["activities"] if a.get("name")]
        sc, more = pick_sscui(sscui.get(sid, []))
        kd = inp["kdd"].get(sid, [])
        ctx = dict(id=sid, name=re.sub(r"\s*\(%s\)\s*$" % sid, "", c["name"]), release=c["release"], countries=countries,
                   desc=short(c.get("description_text"), 1800), flows=flows, activities=list(dict.fromkeys(acts))[:25],
                   sscui=[str(x["name"]) for x in sc], kdd=[dict(id=k["id"], title=k["title"], decision=k["decision"]) for k in kd],
                   extra=short(inp.get("extra", ""), 3000))
        meta[sid] = dict(sscui=sc, sscui_more=more, kdd=kd, scripts=load_script(sid, countries), c=c)
        ctxs[sid] = ctx

    done = [0]
    def gen(sid):
        try:
            if use_llm: return llm_story(model, ctxs[sid])
        except Exception as e:
            log.append(f"{sid}: Claude call failed ({str(e)[:80]}); template wording used.")
        finally:
            done[0] += 1
            if progress: progress(f"Writing stories {done[0]}/{len(ctxs)}")
        return template_story(ctxs[sid])
    with ThreadPoolExecutor(4) as ex:
        stories = dict(zip(ctxs, ex.map(gen, ctxs)))

    def gen_reqs(sid):
        its = inp.get("reqs", {}).get(sid, [])
        if not its: return sid, {}
        try:
            if use_llm: return sid, llm_reqs(model, ctxs[sid], its, inp.get("extra", ""))
        except Exception as e:
            log.append(f"{sid}: Claude could not word the requirement stories ({str(e)[:80]}); template wording used.")
        return sid, template_reqs(ctxs[sid], its)
    if progress: progress("Writing requirement stories")
    with ThreadPoolExecutor(4) as ex:
        inp["_req_texts"] = dict(ex.map(gen_reqs, list(ctxs)))
    for sid in list(inp.get("reqs", {})):
        if sid not in ctxs:
            log.append(f"{sid}: {len(inp['reqs'][sid])} requirement(s) belong to a scope item with no cached data - added as unmapped.")
            inp["unmapped"] = inp.get("unmapped", []) + inp["reqs"].pop(sid)
    rows = []
    def add(**k): rows.append(k)
    used_story_ids = set()
    attach = {}                      # L4 row wbs -> sid to hang the story under
    for r in l4:
        if r["sid"] in ctxs and r["type"] == "Configuration" and r["sid"] not in attach.values():
            attach[r["wbs"]] = r["sid"]
    for sid in ctxs:                 # scope item without a config row: attach to its first row
        if sid not in attach.values():
            first = next((r for r in l4 if r["sid"] == sid), None)
            if first: attach[first["wbs"]] = sid
            else: log.append(f"{sid}: no rows in the L4 plan - story added at the end.")
    for r in l4:
        if r["sid"] and r["sid"] not in selected: continue
        add(type="L4 Activity", id=f"L4-{r['wbs']}", phase=r["phase"], l2=r["l2"], area=r["l3"] if r["sid"] else "",
            l3=r["l3"], title=r["task"], origin="", scope=r["sid"] or "", build=r["type"], parent="", country="All",
            sprint=sprint_no(r["s_sprint"]), pts=None, start=r["start"], end=r["end"], status="Not started", desc=short(r["purpose"], 400), src="L4 plan", lvl=0)
        sid = attach.get(r["wbs"])
        if sid:
            add_story(rows, sid, ctxs[sid], meta[sid], stories[sid], r, inp, used_story_ids)
    unm = inp.get("unmapped", [])
    if unm:
        add(type="L4 Activity", id="L4-UNMAPPED", phase="Explore", l2="", area="", l3="Fit-to-Standard", title="Requirements not yet mapped to a scope item - assign a scope item",
            origin="", scope="", build="Task", parent="", country="All", sprint=None, pts=None, start=None, end=None, status="Not started", desc="", src="User attachments", lvl=0)
        gctx = dict(id="-", name="Unmapped requirements", release="", countries=inp["countries"], desc="", flows=[], activities=[], sscui=[], kdd=[], extra="")
        try:
            texts = llm_reqs(model, gctx, unm, inp.get("extra", "")) if use_llm else template_reqs(gctx, unm)
        except Exception:
            texts = template_reqs(gctx, unm)
        gbase = dict(phase="Realize", l2="", area="", l3="Sprint build", scope="", sprint=None, start=None, end=None, status="To do")
        for n, it in enumerate(unm, 1):
            add_req_story(rows, "", "", it, texts.get(it["id"]), gbase, "L4-UNMAPPED", f"US-UNM-R{n}", inp)
    return rows, meta, ctxs

def add_story(rows, sid, ctx, m, st, l4row, inp, used):
    def add(**k): rows.append(k)
    base = dict(phase="Realize", l2=l4row["l2"], area=l4row["l3"], l3="Sprint build", scope=f"{sid} {ctx['name']}",
                sprint=sprint_no(l4row["s_sprint"]), start=l4row["start"], end=l4row["end"], status="To do")
    sidn = f"US-{sid}"
    parent = f"L4-{l4row['wbs']}"
    n_steps = sum(len(v) for s in m["scripts"] for v in (s["activities"] or {}).values())
    n_cfg = len(m["sscui"]) + m["sscui_more"]
    ac = "\n".join(f"{i}. {a}" for i, a in enumerate(st.get("acceptance_criteria", []), 1))
    add(type="User Story", id=sidn, title=st.get("user_story") or ctx["name"], origin="Core", build="Configuration", parent=parent,
        country="All", pts=points(n_cfg, n_steps), ac=ac, desc="", trace=f"Scope item {sid} (SAP Process Navigator {ctx['release']})", src=f"SAP Process Navigator {sid} release {ctx['release']}", lvl=1, **base)
    # Task 1 confirm design
    t1 = f"{sidn}.T1"
    add(type="Task", id=t1, title="Confirm design inputs", origin="Core", build="Configuration", parent=sidn, country="All",
        pts=None, desc="Fit-to-Standard questions:\n" + "\n".join(f"- {q}" for q in st.get("design_questions", [])), src="Fit-to-Standard", lvl=2, **base)
    for i, q in enumerate(st.get("design_questions", [])[:4], 1):
        add(type="Sub-task", id=f"{t1}.S{i}", title=short(q, 220), origin="Core", build="Configuration", parent=t1, country="All", pts=None, desc="", src="Fit-to-Standard", lvl=3, **base)
    k0 = len(st.get("design_questions", [])[:4])
    for j, k in enumerate(m["kdd"], 1):
        add(type="Sub-task", id=f"{t1}.S{k0+j}", title=short(f"Apply {k['id']}: {k['title']}", 220), origin="KDD", build="Configuration", parent=t1, country="All",
            pts=None, desc=short(f"{k['decision'] or ''} {k['rationale'] or ''}", 400), src=str(k["id"]), trace=str(k["id"]), lvl=3, **base)
    if not m["kdd"]:
        add(type="Sub-task", id=f"{t1}.S{k0+1}", title="Record decisions in the KDD register and link them to this story", origin="Core", build="Configuration", parent=t1, country="All", pts=None, desc="", src="KDD register", lvl=3, **base)
    # Task 2 build
    t2 = f"{sidn}.T2"
    add(type="Task", id=t2, title="Build in the development tenant", origin="Core", build="Configuration", parent=sidn, country="All", pts=None,
        desc=st.get("build_summary", ""), src="SSCUI list / SAP test script", lvl=2, **base)
    n = 0
    prelim = []
    for s in m["scripts"]:
        for name in (s["activities"] or {}):
            if name.lower().startswith("additional information"):
                lbl = name.split(":", 1)[-1].strip()
                if lbl and lbl.lower() != "additional information": prelim.append((lbl, s))
    shown = set()
    for lbl, s in prelim:
        if lbl in shown or len(shown) >= 6: continue
        shown.add(lbl); n += 1
        add(type="Sub-task", id=f"{t2}.S{n}", title=short(lbl, 220), origin="Core", build="Configuration", parent=t2, country=", ".join(s["countries"]),
            pts=None, desc="SAP preliminary step", src=f"Test script {sid}_{s['used']}", lvl=3, **base)
    for it in m["sscui"]:
        n += 1
        add(type="Sub-task", id=f"{t2}.S{n}", title=short(f"Configure SSCUI: {it['name']}", 220), origin="Core", build="Configuration", parent=t2, country="All",
            pts=None, desc=short(f"{it['area']} / {it['sub']}", 120), src=f"SSCUI {it['act_id']}", lvl=3, **base)
    if m["sscui_more"]:
        n += 1
        add(type="Sub-task", id=f"{t2}.S{n}", title=f"Review {m['sscui_more']} further scope-specific configuration activities in the SSCUI list", origin="Core", build="Configuration", parent=t2,
            country="All", pts=None, desc="", src="SSCUI list", lvl=3, **base)
    if n == 0:
        add(type="Sub-task", id=f"{t2}.S1", title="Activate the scope item in Manage Your Solution and complete Configure Your Solution", origin="Core", build="Configuration", parent=t2, country="All", pts=None, desc="", src="SAP Best Practices", lvl=3, **base)
    # Task 3 test
    t3 = f"{sidn}.T3"
    add(type="Task", id=t3, title="Test and hand over", origin="Core", build="Configuration", parent=sidn, country="All", pts=None, desc="", src="SAP test script", lvl=2, **base)
    n = 0
    for s in m["scripts"]:
        if s["used"] is None:
            n += 1
            add(type="Sub-task", id=f"{t3}.S{n}", title=f"Obtain the SAP test script for {', '.join(s['countries'])} (not yet cached) and unit test the scope item", origin="Core", build="Configuration",
                parent=t3, country=", ".join(s["countries"]), pts=None, desc="Test script missing in local cache - download needed.", src="SAP test script", lvl=3, **base)
            continue
        tests = [(a, v) for a, v in s["activities"].items() if not a.lower().startswith("additional information")]
        for a, v in tests[:10]:
            n += 1
            steps = "\n".join(f"{i}. {x['action']}: {short(x['text'], 160)}" for i, x in enumerate(v[:8], 1))
            exp = next((short(x["expected"], 160) for x in reversed(v) if x["expected"]), "")
            tag = f" (script {s['used']}, fallback)" if s["fallback"] else ""
            add(type="Sub-task", id=f"{t3}.S{n}", title=short(f"Unit test: {a}", 220), origin="Core", build="Configuration", parent=t3, country=", ".join(s["countries"]),
                pts=None, desc=short(steps + (f"\nExpected: {exp}" if exp else ""), 900), src=f"Test script {sid}_{s['used']}{tag}", lvl=3, **base)
    for lbl, who in [("Write the test case and link it to this story", None), ("Transport configuration to the test tenant", None), ("Demo at sprint review and get acceptance", None)]:
        n += 1
        add(type="Sub-task", id=f"{t3}.S{n}", title=lbl, origin="Core", build="Configuration", parent=t3, country="All", pts=None, desc="", src="Standard", lvl=3, **base)
    # requirement stories from user attachments (workshop notes, gap/WRICEF, backlog)
    for n_r, it in enumerate(inp.get("reqs", {}).get(sid, []), 1):
        add_req_story(rows, sid, ctx["name"], it, inp.get("_req_texts", {}).get(sid, {}).get(it["id"]), base, parent, f"US-{sid}-R{n_r}", inp)

# ----------------------------------------------------------------- Excel
FILL = dict(h="1F3864", l4="DDE7F5", story="E2EFDA", task="F2F2F2")

def write(path, rows, meta, ctxs, inp, release, log, used_llm):
    wb = openpyxl.Workbook()
    rm = wb.active; rm.title = "Read Me"
    lines = ["SAP S/4HANA Public Cloud: L4 Plan and User Stories (generated)", "",
             f"Generated {dt.datetime.now():%Y-%m-%d %H:%M} by the User Story Agent. SAP release: {release}.",
             f"Countries: {', '.join(inp['countries'])}. Rows with a Country value of 'All' apply to every country.",
             "Blue rows: L4 Activity (from the L4 plan). Green: User Story. Grey: Task. White: Sub-task. Parent ID links each row upward.",
             "Owner is blank by design. Story points and sprints are agent estimates - review them.",
             "Story wording " + ("written by Claude (via the user's Claude Code login) from SAP Process Navigator data." if used_llm else "uses template wording (Claude not used for this run)."),
             "Test steps come from the SAP Cloud ALM test script for the country; where none is cached, the global (XX) or US script is used and marked 'fallback'.",
             "", "Run notes:"] + [f"- {x}" for x in (log or ["none"])]
    for l in lines: rm.append([l])
    rm["A1"].font = Font(bold=True, size=14); rm.column_dimensions["A"].width = 150

    reg = wb.create_sheet("Scope Item Register")
    reg.append(["Scope Item ID", "Name", "Release", "Countries per SAP", "Test script used (country)", "SSCUI count", "KDD count"])
    for sid, c in ctxs.items():
        m = meta[sid]
        reg.append([sid, c["name"], c["release"], m["c"].get("countries"),
                    "; ".join(f"{','.join(s['countries'])}→{s['used'] or 'MISSING'}" for s in m["scripts"]),
                    len(m["sscui"]) + m["sscui_more"], len(m["kdd"])])
    for i, w in enumerate([14, 50, 10, 60, 34, 12, 10], 1): reg.column_dimensions[get_column_letter(i)].width = w

    ws = wb.create_sheet("L4 Plan and User Stories")
    ws.append(["L4 Plan and User Stories"]); ws["A1"].font = Font(bold=True, size=14)
    ws.append(["Row types: L4 Activity > User Story > Task > Sub-task. Use outline buttons to collapse."]); ws.append([]); ws.append([])
    ws.append(COLS)
    for c in ws[5]:
        c.fill = PatternFill("solid", fgColor=FILL["h"]); c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        ws.append([r["type"], r["id"], r.get("phase"), r.get("l2"), r.get("area"), r.get("l3"), r["title"], r.get("origin"), r.get("scope"),
                   r.get("build"), None, r.get("parent"), r.get("country"), r.get("sprint"), r.get("pts"), r.get("start"), r.get("end"),
                   r.get("status"), r.get("ac"), r.get("desc"), r.get("trace"), r.get("src")])
        i = ws.max_row
        ws.row_dimensions[i].outlineLevel = r["lvl"]
        col = {"L4 Activity": "l4", "User Story": "story", "Task": "task"}.get(r["type"])
        if col:
            for c in ws[i][:len(COLS)]: c.fill = PatternFill("solid", fgColor=FILL[col])
        for k in (16, 17): ws.cell(i, k).number_format = "yyyy-mm-dd"
        for k in (19, 20): ws.cell(i, k).alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate([12, 26, 10, 22, 22, 28, 60, 14, 28, 14, 14, 24, 14, 8, 8, 12, 12, 12, 60, 50, 36, 30], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "H6"; ws.auto_filter.ref = f"A5:{get_column_letter(len(COLS))}{ws.max_row}"
    ws.sheet_properties.outlinePr.summaryBelow = False
    tr = wb.create_sheet("Traceability")
    tr.append(["Requirement / item ID", "Kind", "Source file", "Title", "Scope item", "User Story ID", "KDD references", "Tasks + sub-tasks"])
    for c in tr[1]:
        c.fill = PatternFill("solid", fgColor=FILL["h"]); c.font = Font(bold=True, color="FFFFFF")
    for r in rows:
        if r["type"] == "User Story":
            tr.append([r.get("req_id") or "Standard scope item", r.get("req_kind") or "scope item", r.get("src"), r["title"], r.get("scope") or "unmapped",
                       r["id"], (r.get("trace") or "").split("KDD: ")[-1] if "KDD:" in (r.get("trace") or "") else "", sum(1 for x in rows if x["id"].startswith(r["id"] + "."))])
    for i, w in enumerate([22, 14, 36, 70, 30, 18, 24, 12], 1): tr.column_dimensions[get_column_letter(i)].width = w
    tr.freeze_panes = "A2"
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)

def generate(cfg, progress=None):
    """cfg: scope_ids, countries, files{slot:[paths]}, free_text, l4, sscui, out, use_llm, model, output{mode:'default'|'template', template},
    optional workbook (scope/countries/KDD/backlog sheets). Returns dict(out, rows, scope_items, stories, log, used_llm)."""
    out = Path(cfg["out"])
    if out.exists(): raise FileExistsError(f"Refusing to overwrite existing file: {out}")
    log = []
    say = progress or (lambda m: None)
    say("Reading inputs")
    inp = dict(scope=[], override={}, countries=[], kdd=collections.defaultdict(list), reqs=collections.defaultdict(list), unmapped=[], extra="")
    if cfg.get("workbook"):
        w = load_input(cfg["workbook"])
        inp.update(scope=w["scope"], override=w["override"], countries=w["countries"])
        for sid, ks in w["kdd"].items():
            inp["kdd"][sid] += [dict(id=k["id"], title=k["title"], decision=k["decision"] or "", rationale=k["rationale"] or "") for k in ks]
        for sid, bs in w["backlog"].items():
            for b in bs:
                inp["reqs"][sid].append(dict(id=str(b["id"]), title=str(b["title"] or ""), description=str(b["desc"] or ""), scope_item=sid, category="",
                                             priority=str(b["prio"] or ""), decision="", rationale="", source="input workbook", kind="backlog"))
    if cfg.get("scope_ids"): inp["scope"] = list(dict.fromkeys(x.strip().upper() for x in cfg["scope_ids"] if x.strip()))
    if cfg.get("countries"): inp["countries"] = [c.strip().upper() for c in cfg["countries"] if c.strip()]
    if not inp["scope"]: raise ValueError("No scope item IDs given.")
    if not inp["countries"]: raise ValueError("No country given.")
    l4 = load_l4(cfg["l4"]); release, sscui = load_sscui(cfg["sscui"])
    use_llm = cfg.get("use_llm", True)
    if use_llm:
        import llm
        st = llm.status()
        if not st["loggedIn"]:
            log.append(st["detail"] + " Template wording used, and attachments that are not simple tables cannot be read - sign in to Claude and re-run.")
            use_llm = False
    model = cfg.get("model")
    import requirements as rq
    import readers
    names = {}
    for sid in inp["scope"]:
        c = load_scope_cache(sid)
        names[sid] = re.sub(r"\s*\(%s\)\s*$" % sid, "", c["name"]) if c and c.get("found") else sid
    ask = (lambda p: llm.ask_json(p, model)) if use_llm else None
    files = cfg.get("files", {})
    for slot in ("workshop", "kdd", "gap", "backlog"):
        if files.get(slot):
            items = rq.extract(slot, files[slot], names, use_llm, log, say, ask)
            for it in items:
                if slot == "kdd":
                    if it["scope_item"]: inp["kdd"][it["scope_item"]].append(it)
                    else: log.append(f"KDD {it['id']} could not be linked to a selected scope item.")
                elif it["scope_item"]: inp["reqs"][it["scope_item"]].append(it)
                else: inp["unmapped"].append(it)
    extra = []
    for slot in ("bpd", "trace", "other"):
        for f in files.get(slot, []):
            t, note = readers.read_text(f)
            if note: log.append(note)
            elif t: extra.append(f"[{Path(f).name}]\n{t[:4000]}")
    if cfg.get("free_text"): extra.append("[User note]\n" + cfg["free_text"][:4000])
    inp["extra"] = "\n\n".join(extra)
    if extra and not use_llm: log.append("BPD, traceability, other files and free text are only used when Claude is signed in.")
    rows, meta, ctxs = build(inp, l4, sscui, use_llm, model, log, say)
    for sid, m in meta.items():
        for s in m["scripts"]:
            if s["used"] is None: log.append(f"{sid}: test script for {', '.join(s['countries'])} not cached (no country, global or US file).")
            elif s["fallback"]: log.append(f"{sid}: no {', '.join(c for c in s['countries'] if c != s['used'])} test script cached - used {s['used']}.")
    mode = (cfg.get("output") or {}).get("mode", "default")
    say("Writing Excel")
    info = {}
    if mode == "template":
        import export_template
        info = export_template.write_template(rows, cfg["output"]["template"], out, use_llm, model, log, say)
    else:
        write(out, rows, meta, ctxs, inp, release, log, use_llm)
    return dict(out=str(out), rows=len(rows), scope_items=len(ctxs), stories=sum(r["type"] == "User Story" for r in rows),
                log=log, used_llm=use_llm, mode=mode, mapping=info.get("mapping"))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", help="workbook with Scope_Items/Countries (and optional KDD/Product_Backlog sheets)")
    ap.add_argument("--scope", help="comma-separated scope item IDs"); ap.add_argument("--countries", help="comma-separated country codes")
    ap.add_argument("--l4", required=True); ap.add_argument("--sscui", default=str(Path(__file__).parent / "reference" / "SSCUI_List.xlsm"))
    ap.add_argument("--file", action="append", default=[], help="slot=path, slot in workshop|kdd|gap|backlog|bpd|trace|other")
    ap.add_argument("--template", help="fill this output template instead of the sample layout")
    ap.add_argument("--out", required=True); ap.add_argument("--no-llm", action="store_true"); ap.add_argument("--model", default=None)
    a = ap.parse_args()
    files = collections.defaultdict(list)
    for f in a.file:
        k, v = f.split("=", 1); files[k].append(v)
    cfg = dict(workbook=a.input, scope_ids=(a.scope or "").split(",") if a.scope else None, countries=(a.countries or "").split(",") if a.countries else None,
               files=files, l4=a.l4, sscui=a.sscui, out=a.out, use_llm=not a.no_llm, model=a.model,
               output=dict(mode="template", template=a.template) if a.template else dict(mode="default"))
    r = generate(cfg, print)
    print(f"Wrote {r['out']}: {r['rows']} rows, {r['scope_items']} scope items, {r['stories']} stories (Claude: {r['used_llm']})")
    for l in r["log"]: print(" -", l)

if __name__ == "__main__":
    main()
