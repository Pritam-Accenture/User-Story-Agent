"""Fill a user-supplied output template (Cloud ALM, ADM tool, ADO/Jira Excel, anything) with the generated rows.

1. analyze()   - read the template: sheets, header rows, sample rows, dropdown lists.
2. mapping     - Claude proposes how our neutral fields map to the template columns (type names, parent link,
                 value translations). Without Claude a header-keyword heuristic is used. The mapping is cached
                 per template in agent/profiles/ and shown to the user in the app.
3. fill        - rows are written into a copy of the template, keeping its styles, dropdowns and other sheets.
"""
import copy
import datetime as dt
import hashlib
import json
import re
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

PROFILES = Path(__file__).resolve().parent / "profiles"

# neutral field -> key in our row dicts
FIELDS = {
    "type": ("type", "Row type: L4 Activity, User Story, Task or Sub-task"),
    "id": ("id", "Our unique ID"),
    "title": ("title", "Title / summary"),
    "description": ("desc", "Description"),
    "acceptance_criteria": ("ac", "Acceptance criteria (Given/When/Then lines)"),
    "description_with_ac": (None, "Description followed by the acceptance criteria in one text"),
    "parent_id": ("parent", "Parent row ID (or parent title when the template links by title)"),
    "phase": ("phase", "SAP Activate phase"),
    "workstream": ("l2", "Workstream / module"),
    "sap_area": ("area", "SAP area"),
    "deliverable": ("l3", "L3 deliverable"),
    "scope_items": ("scope", "SAP scope item(s)"),
    "country": ("country", "Country code(s) or All"),
    "build_type": ("build", "Configuration, Workflow, Report, Interface, Conversion, Enhancement, Form"),
    "story_origin": ("origin", "Core, Workshop, Gap / WRICEF, Product Backlog, KDD"),
    "sprint": ("sprint", "Sprint number"),
    "story_points": ("pts", "Story points"),
    "start_date": ("start", "Start date"),
    "end_date": ("end", "End date"),
    "status": ("status", "Status: To do / Not started"),
    "traceability": ("trace", "Requirement / KDD / scope item traceability text"),
    "source_reference": ("src", "Source reference"),
}
SYN = {
    "type": ["work item type", "type", "issue type", "row type", "item type", "entity type", "category"],
    "id": ["id", "key", "work item id", "external id", "issue key", "requirement id", "story id"],
    "title": ["title", "summary", "name", "work item title", "subject", "story", "short text", "requirement"],
    "description": ["description", "details", "long description", "user story", "text"],
    "acceptance_criteria": ["acceptance criteria", "acceptance", "ac", "definition of done"],
    "parent_id": ["parent", "parent id", "parent key", "epic link", "parent link", "parent item"],
    "phase": ["phase", "activate phase"],
    "workstream": ["workstream", "module", "area path", "work stream", "component"],
    "scope_items": ["scope item", "scope items", "scope", "scope item id"],
    "country": ["country", "countries", "region"],
    "build_type": ["build type", "object type", "wricef type"],
    "story_origin": ["story origin", "origin", "source"],
    "sprint": ["sprint", "iteration", "iteration path"],
    "story_points": ["story points", "points", "effort", "estimate", "size"],
    "start_date": ["start date", "start", "planned start"],
    "end_date": ["end date", "due date", "finish", "target date", "planned end"],
    "status": ["status", "state"],
    "traceability": ["traceability", "linked requirement", "requirement ref"],
    "source_reference": ["source reference", "reference", "source document"],
}
DEFAULT_TYPES = {"L4 Activity": "Epic", "User Story": "User Story", "Task": "Task", "Sub-task": "Task"}


def _norm(s):
    return re.sub(r"[^a-z0-9 ]", "", str(s or "").lower()).strip()


def file_hash(p):
    return hashlib.sha1(Path(p).read_bytes()).hexdigest()[:16]


def _dv_lists(ws, wb):
    """column letter -> allowed values (inline lists and simple same-workbook ranges)."""
    out = {}
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list" or not dv.formula1:
            continue
        f = dv.formula1
        vals = None
        if f.startswith('"'):
            vals = [x.strip() for x in f.strip('"').split(",")]
        else:
            m = re.fullmatch(r"=?'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", f)
            if m and m.group(1) in wb.sheetnames:
                rs = wb[m.group(1)][f"{m.group(2)}{m.group(3)}:{m.group(4)}{m.group(5)}"]
                vals = [str(c.value) for row in rs for c in row if c.value not in (None, "")]
        if not vals:
            continue
        for rng in str(dv.sqref).split():
            m = re.match(r"\$?([A-Z]+)\$?\d+", rng)
            if m:
                out[m.group(1)] = vals[:60]
    return out


def analyze(path):
    wb = openpyxl.load_workbook(path, keep_vba=str(path).lower().endswith(".xlsm"))
    sheets = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        rows = []
        for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 14)):
            cells = {c.coordinate: str(c.value)[:70] for c in r if c.value not in (None, "")}
            if cells:
                rows.append({"row": r[0].row, "cells": cells})
        sheets.append(dict(name=ws.title, max_row=ws.max_row, max_col=ws.max_column, rows=rows, dropdowns=_dv_lists(ws, wb)))
    return dict(sheets=sheets)


def heuristic_mapping(summary):
    """Pick the sheet/row that looks most like a header and match headers to fields by keyword."""
    best = None
    for sh in summary["sheets"]:
        for r in sh["rows"]:
            cmap = {}
            for coord, v in r["cells"].items():
                col = re.match(r"[A-Z]+", coord).group()
                n = _norm(v)
                n_words = re.sub(r"\d+", "", n).split()
                for field, syns in SYN.items():
                    if field in cmap.values():
                        continue
                    if n in syns or any((s_ in n_words) if " " not in s_ else (s_ in n) for s_ in syns):
                        cmap[col] = field
                        break
            score = len(cmap) + (3 if "title" in cmap.values() else 0)
            if "title" in cmap.values() and (best is None or score > best[0]):
                best = (score, sh, r["row"], cmap)
    if not best:
        raise ValueError("Could not find a header row with a Title/Summary column in the template - please check the template.")
    _, sh, hr, cmap = best
    types = {}
    tcol = next((c for c, f in cmap.items() if f == "type"), None)
    allowed = sh["dropdowns"].get(tcol, []) if tcol else []
    low = {a.lower(): a for a in allowed}
    for ours, dflt in DEFAULT_TYPES.items():
        types[ours] = low.get(dflt.lower()) or low.get(ours.lower()) or (allowed[0] if allowed and dflt not in allowed else dflt)
    value_maps = {}
    scol = next((c for c, f in cmap.items() if f == "status"), None)
    sallowed = sh["dropdowns"].get(scol, []) if scol else []
    if sallowed:
        pref = next((a for k in ("new", "to do", "todo", "open", "not started", "backlog", "proposed") for a in sallowed if a.lower() == k), sallowed[0])
        value_maps["status"] = {"To do": pref, "Not started": pref}
    return dict(sheets=[dict(sheet=sh["name"], header_row=hr, first_data_row=hr + 1, types=["L4 Activity", "User Story", "Task", "Sub-task"],
                             columns={c: {"field": f} for c, f in cmap.items()})],
                type_map=types, value_maps=value_maps, parent_mode="id", notes="Header-keyword mapping (Claude not used).")


def claude_mapping(summary, model):
    import llm
    flds = "\n".join(f"- {k}: {v[1]}" for k, v in FIELDS.items())
    prompt = f"""You are mapping generated SAP implementation backlog rows into a user-supplied Excel import template
(it may be for SAP Cloud ALM, Azure DevOps, Jira, or a custom tool).

TEMPLATE ANALYSIS (sheet names, first rows with cell coordinates, dropdown values per column letter):
{json.dumps(summary, ensure_ascii=False)[:14000]}

Our neutral fields:
{flds}
Our row types: "L4 Activity" (a plan activity, usually an Epic or Feature), "User Story", "Task", "Sub-task".
Our status values: "Not started", "To do". Our parents are linked by our own ID.

Return ONLY JSON:
{{"sheets": [{{"sheet": "<sheet name>", "header_row": <int>, "first_data_row": <int>,
   "types": [<our row types written to this sheet>],
   "columns": {{"<column letter>": {{"field": "<neutral field>"}} or {{"const": "<fixed value>"}} }}}}],
 "type_map": {{"L4 Activity": "<template work item type>", "User Story": "...", "Task": "...", "Sub-task": "..."}},
 "value_maps": {{"status": {{"To do": "<template status>", "Not started": "<template status>"}}}},
 "parent_mode": "id" | "title" | "none",
 "notes": "<one sentence on what you assumed>"}}

Rules: only use columns that exist in the template; use dropdown values exactly as listed; if the template has one sheet per level, return one entry per sheet with its `types`; leave columns you cannot fill out of `columns`; never invent column letters; owner/assignee columns stay unmapped."""
    m = llm.ask_json(prompt, model)
    for sh in m.get("sheets", []):
        sh["columns"] = {str(k).upper(): v for k, v in sh.get("columns", {}).items()}
    return m


def _check_mapping(m, summary):
    names = {s["name"]: s for s in summary["sheets"]}
    if not m.get("sheets"):
        raise ValueError("empty mapping")
    for sh in m["sheets"]:
        if sh["sheet"] not in names:
            raise ValueError(f"unknown sheet {sh['sheet']}")
        sh["header_row"] = int(sh["header_row"]); sh["first_data_row"] = int(sh.get("first_data_row") or sh["header_row"] + 1)
        sh.setdefault("types", ["L4 Activity", "User Story", "Task", "Sub-task"])
        for col, spec in list(sh["columns"].items()):
            f = spec.get("field")
            if ("const" not in spec) and (f not in FIELDS or f == "owner"):
                del sh["columns"][col]
    m.setdefault("type_map", DEFAULT_TYPES); m.setdefault("value_maps", {}); m.setdefault("parent_mode", "id")
    return m


def get_mapping(template, use_llm, model, log, say=None, force=False):
    PROFILES.mkdir(exist_ok=True)
    pf = PROFILES / f"{file_hash(template)}.json"
    summary = analyze(template)
    if pf.exists() and not force:
        d = json.loads(pf.read_text(encoding="utf-8"))
        if d.get("by_claude") or not use_llm:
            return d["mapping"], summary, True
    m, by_claude = None, False
    if use_llm:
        try:
            if say: say("Claude is reading your output template")
            m = _check_mapping(claude_mapping(summary, model), summary); by_claude = True
        except Exception as e:
            log.append(f"Claude could not map the template ({str(e)[:90]}); keyword mapping used - please check the result.")
    if m is None:
        m = _check_mapping(heuristic_mapping(summary), summary)
    pf.write_text(json.dumps(dict(mapping=m, by_claude=by_claude, template=Path(template).name), indent=1), encoding="utf-8")
    return m, summary, False


def _value(r, field, rows_by_id, mapping):
    if field == "description_with_ac":
        v = (r.get("desc") or "")
        if r.get("ac"):
            v = (v + "\n\n" if v else "") + "Acceptance criteria:\n" + r["ac"]
        return v or None
    key = FIELDS[field][0]
    v = r.get(key)
    if field == "type":
        return mapping["type_map"].get(r["type"], r["type"])
    if field == "parent_id":
        if not v or mapping["parent_mode"] == "none":
            return None
        return rows_by_id[v]["title"] if mapping["parent_mode"] == "title" and v in rows_by_id else v
    vm = mapping.get("value_maps", {}).get(field)
    if vm and v in vm:
        return vm[v]
    if isinstance(v, dt.date) or v is None:
        return v
    return v


def write_template(rows, template, out, use_llm, model, log, say=None):
    template, out = Path(template), Path(out)
    mapping, summary, cached = get_mapping(template, use_llm, model, log, say)
    if out.suffix.lower() != template.suffix.lower():
        raise ValueError(f"Output file must keep the template extension ({template.suffix}).")
    wb = openpyxl.load_workbook(template, keep_vba=template.suffix.lower() == ".xlsm")
    rows_by_id = {r["id"]: r for r in rows}
    written = {}
    problems = set()
    for sh in mapping["sheets"]:
        ws = wb[sh["sheet"]]
        dd = _dv_lists(ws, wb)
        first = sh["first_data_row"]
        cols = {column_index_from_string(c): spec for c, spec in sh["columns"].items()}
        # clear example rows (values only) in mapped columns
        for rr in range(first, ws.max_row + 1):
            for ci in cols:
                ws.cell(rr, ci).value = None
        style_src = {ci: ws.cell(first, ci) for ci in range(1, ws.max_column + 1)}
        style_copy = {ci: copy.copy(c._style) for ci, c in style_src.items()}
        rr = first
        types = set(sh["types"])
        for r in rows:
            if r["type"] not in types:
                continue
            for ci in range(1, ws.max_column + 1):
                ws.cell(rr, ci)._style = copy.copy(style_copy[ci])
            for ci, spec in cols.items():
                v = spec["const"] if "const" in spec else _value(r, spec["field"], rows_by_id, mapping)
                allowed = dd.get(get_letter(ci))
                if allowed and v not in (None, "") and str(v) not in allowed:
                    hit = next((a for a in allowed if a.lower() == str(v).lower()), None)
                    if hit: v = hit
                    else: problems.add(f"'{v}' is not in the dropdown of column {get_letter(ci)} on sheet '{sh['sheet']}'")
                ws.cell(rr, ci).value = v
                if isinstance(v, dt.date):
                    ws.cell(rr, ci).number_format = ws.cell(rr, ci).number_format if ws.cell(rr, ci).number_format != "General" else "yyyy-mm-dd"
            rr += 1
        written[sh["sheet"]] = rr - first
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    for p in sorted(problems)[:15]:
        log.append("Template check: " + p)
    log.append(f"Template: {template.name} - " + ", ".join(f"{n} rows on '{s}'" for s, n in written.items()) +
               f" ({'saved mapping reused' if cached else 'mapping created'}).")
    return dict(mapping=mapping, rows_written=written)


def get_letter(ci):
    return get_column_letter(ci)
