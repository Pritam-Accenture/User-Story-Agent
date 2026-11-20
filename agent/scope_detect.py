"""Auto-detect SAP scope item IDs in the user's uploaded files.

A 3-character token counts only if it is a known scope item ID (SSCUI list shipped with the app, the local Process Navigator
extract, or the Scope ID column of the user's L4 plan) and is written in capitals - so ordinary words never match.
"""
import json
import re
from pathlib import Path

import readers

STOP = {"ALL", "NA"}
TOKEN = re.compile(r"(?<![A-Za-z0-9_])[0-9][0-9A-Z]{2}(?![A-Za-z0-9_])|(?<![A-Za-z0-9_])[A-Z][0-9A-Z]{2}(?![A-Za-z0-9_])")
_known_cache = {}


def known_ids(sscui_path, cache_dir, l4_ids=()):
    key = (str(sscui_path), Path(sscui_path).stat().st_mtime if Path(sscui_path).exists() else 0)
    if key not in _known_cache:
        import build_plan
        try:
            _, by = build_plan.load_sscui(sscui_path)
        except Exception:
            by = {}
        _known_cache.clear()
        _known_cache[key] = {i for i in by if len(i) == 3 and i not in STOP}
    ids = set(_known_cache[key])
    d = Path(cache_dir) / "scope_items"
    if d.exists():
        ids |= {p.stem.upper() for p in d.glob("*.json") if len(p.stem) == 3}
    return ids | {i for i in l4_ids if i and len(i) == 3}


def all_known(sscui_path, cache_dir):
    """Every known scope item ID with its name (for the picker): [{id, name}]."""
    nm = names(cache_dir)
    return [dict(id=i, name=nm.get(i, "")) for i in sorted(known_ids(sscui_path, cache_dir))]


def names(cache_dir):
    out = {}
    d = Path(cache_dir) / "scope_items"
    for p in (d.glob("*.json") if d.exists() else []):
        try:
            j = json.loads(p.read_text(encoding="utf-8"))
            if str(j.get("found")).lower() == "true" and j.get("name"):
                out[p.stem.upper()] = re.sub(r"\s*\([0-9A-Z]{3}\)\s*$", "", j["name"])
        except Exception:
            pass
    return out


ALIASES = {"US": ["USA", "U.S.", "U.S.A.", "United States of America"], "GB": ["UK", "U.K.", "Great Britain", "England"],
           "AE": ["UAE"], "KR": ["Korea", "South Korea"], "CZ": ["Czech Republic"], "TR": ["Turkey", "Turkiye", "Türkiye"],
           "NL": ["Holland"], "RU": ["Russia", "Russian Federation"], "CN": ["PRC"]}


def country_patterns(countries):
    pats = {}
    for c in countries:
        names = [c["name"]] + ALIASES.get(c["code"], [])
        pats[c["code"]] = re.compile(r"(?<![A-Za-z])(?:" + "|".join(re.escape(n) for n in names) + r")(?![A-Za-z])")
    return pats


def detect(files_by_slot, sscui_path, cache_dir, countries_path=None):
    """files_by_slot: {slot: [Path]}. Returns {scope: [{id, name, count, sources}], countries: [{code, name, count, sources}]}."""
    import build_plan
    l4_ids = set()
    for p in files_by_slot.get("l4", []):
        try:
            l4_ids |= {r["sid"].upper() for r in build_plan.load_l4(p) if r.get("sid")}
        except Exception:
            pass
    known = known_ids(sscui_path, cache_dir, l4_ids)
    countries = json.loads(Path(countries_path).read_text(encoding="utf-8")) if countries_path else []
    cpats = country_patterns(countries)
    cfound = {}
    found = {}
    for slot, paths in files_by_slot.items():
        if slot == "template":
            continue
        for p in paths:
            try:
                text = readers.read_text(p)
            except Exception:
                continue
            if isinstance(text, tuple):
                text = text[0]
            for code, pat in cpats.items():
                n = len(pat.findall(text or ""))
                if n:
                    e = cfound.setdefault(code, dict(count=0, sources=[]))
                    e["count"] += n
                    if p.name not in e["sources"]:
                        e["sources"].append(p.name)
            for t in TOKEN.findall(text or ""):
                if t in known:
                    e = found.setdefault(t, dict(count=0, sources=[]))
                    e["count"] += 1
                    if p.name not in e["sources"]:
                        e["sources"].append(p.name)
    nm = names(cache_dir)
    cn = {c["code"]: c["name"] for c in countries}
    return dict(scope=[dict(id=i, name=nm.get(i, ""), count=v["count"], sources=v["sources"]) for i, v in sorted(found.items())],
                countries=[dict(code=i, name=cn[i], count=v["count"], sources=v["sources"]) for i, v in sorted(cfound.items(), key=lambda kv: -kv[1]["count"])])
