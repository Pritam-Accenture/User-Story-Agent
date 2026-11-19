"""Turn user attachments (workshop notes, KDD, gap/WRICEF list, backlog, ...) into structured items.

Each item: dict(id, title, description, scope_item, category, priority, decision, rationale, source)
 - Tabular files (xlsx/csv) are parsed by header keywords (no Claude needed).
 - Any other format, or a table with no recognisable header, is read by Claude (needs the Claude login).
"""
import json
import re
from pathlib import Path

import readers

KINDS = {
    "workshop": "Fit-to-Standard workshop notes / minutes. Extract every confirmed or requested requirement, gap, decision and action that needs build or configuration.",
    "kdd": "Key Design Decision (KDD) register. Extract every design decision.",
    "gap": "Gap / WRICEF list. Extract every gap or WRICEF object (Workflow, Report, Interface, Conversion, Enhancement, Form).",
    "backlog": "Product backlog. Extract every backlog item.",
}
SYN = {
    "id": ["id", "#", "no", "ref", "reference", "kdd id", "req id", "requirement id", "item id", "number", "key"],
    "title": ["title", "name", "summary", "requirement", "short description", "topic", "subject", "decision title", "gap", "object"],
    "description": ["description", "details", "detail", "requirement description", "long description", "comment", "notes", "specification"],
    "scope_item": ["scope item", "scope", "scope id", "scope item id", "si", "scope items"],
    "decision": ["decision", "outcome", "resolution", "agreed"],
    "rationale": ["rationale", "reason", "justification", "impact"],
    "priority": ["priority", "moscow", "prio"],
    "category": ["type", "category", "wricef type", "object type", "wricef", "classification", "gap type"],
}


def _norm(s):
    return re.sub(r"[^a-z0-9# ]", "", str(s or "").lower()).strip()


def _find_header(rows):
    best = (0, None, {})
    for i, r in enumerate(rows[:25]):
        cmap = {}
        for j, c in enumerate(r):
            n = _norm(c)
            if not n:
                continue
            for field, syns in SYN.items():
                if field not in cmap and n in syns:
                    cmap[field] = j
            if "id" not in cmap and n.endswith(" id") and n not in ("scope item id", "scope id"):
                cmap["id"] = j
        score = len(cmap) + (2 if "title" in cmap else 0)
        if ("title" in cmap or "description" in cmap) and score > best[0]:
            best = (score, i, cmap)
    return best[1], best[2]


def _from_table(path, source, scope_ids):
    items = []
    for _, rows in readers.table_rows(path):
        h, cmap = _find_header(rows)
        if h is None:
            continue
        for r in rows[h + 1:]:
            def g(f):
                j = cmap.get(f)
                return str(r[j]).strip() if j is not None and j < len(r) and r[j] not in (None, "") else ""
            if not g("title") and not g("description"):
                continue
            si = ""
            for cand in re.findall(r"\b[0-9A-Z]{3}\b", g("scope_item").upper()):
                if cand in scope_ids:
                    si = cand; break
            items.append(dict(id=g("id"), title=g("title") or g("description")[:120], description=g("description"), scope_item=si,
                              category=g("category"), priority=g("priority"), decision=g("decision"), rationale=g("rationale"), source=source))
    return items


def _prompt(kind, scope_names, text):
    sl = "\n".join(f"- {k}: {v}" for k, v in scope_names.items())
    return f"""You are an SAP S/4HANA Public Cloud functional consultant using SAP Activate.
Document type: {KINDS[kind]}

Scope items in this project:
{sl}

Read the document below and return ONLY JSON: {{"items": [ ... ]}} where each item has these keys:
 id (reference id from the document, else ""), title (one line), description (what is needed, 1-3 sentences),
 scope_item (one of the scope item IDs above that this belongs to, else ""), category (Configuration | Gap | Workflow | Report | Interface | Conversion | Enhancement | Form),
 priority (Must | Should | Could | Won't, or "" if not stated), decision (the decision taken, if any), rationale (why, if stated).
Use only facts in the document. Do not invent items. If there are none, return {{"items": []}}.

DOCUMENT:
{text}"""


def extract(kind, files, scope_names, use_llm, log, progress=None, ask_json=None):
    """files: list of paths. Returns list of items."""
    scope_ids = set(scope_names)
    items = []
    for f in files:
        f = Path(f)
        source = f.name
        got = []
        if f.suffix.lower() in (".xlsx", ".xlsm", ".csv", ".tsv"):
            try:
                got = _from_table(f, source, scope_ids)
            except Exception as e:
                log.append(f"{source}: table could not be parsed ({str(e)[:60]}).")
        if not got:
            text, note = readers.read_text(f)
            if note:
                log.append(note); continue
            if not use_llm:
                log.append(f"{source}: needs Claude to read this {kind} file (not a recognisable table) - sign in to Claude and run again.")
                continue
            parts = readers.chunks(text)[:10]
            if len(readers.chunks(text)) > 10:
                log.append(f"{source}: long document - first 10 sections read.")
            for n, part in enumerate(parts, 1):
                if progress: progress(f"Reading {source} ({n}/{len(parts)})")
                try:
                    d = ask_json(_prompt(kind, scope_names, part))
                    for it in d.get("items", []):
                        it = {k: str(it.get(k, "") or "").strip() for k in ("id", "title", "description", "scope_item", "category", "priority", "decision", "rationale")}
                        it["scope_item"] = it["scope_item"].upper() if it["scope_item"].upper() in scope_ids else ""
                        it["source"] = source
                        if it["title"]:
                            got.append(it)
                except Exception as e:
                    log.append(f"{source}: Claude could not read section {n} ({str(e)[:60]}).")
        # scope item by mention if still unassigned
        text_ids = None
        for it in got:
            if not it["scope_item"]:
                blob = f"{it['title']} {it['description']}".upper()
                hits = [s for s in scope_ids if re.search(rf"\b{re.escape(s)}\b", blob)]
                if len(hits) == 1:
                    it["scope_item"] = hits[0]
        items += got
    prefix = {"workshop": "WS", "kdd": "KDD", "gap": "GAP", "backlog": "BL"}[kind]
    for n, it in enumerate(items, 1):
        it["kind"] = kind
        it["id"] = it["id"] or f"{prefix}-{n:03d}"
    return items
