"""Local backend for the User Story Agent (runs on each user's own machine).

  python agent/server.py        ->  http://127.0.0.1:8000

Each user signs in to Claude (Claude Code login) and to SAP Process Navigator (Edge window) on their own
machine; nothing is shared between users.
"""
import datetime as dt
import json
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
ROOT = HERE.parent
WORK = HERE / "workspace"
FILES = WORK / "files"
OUT = ROOT / "Output"
CACHE = HERE / "cache"
REF = HERE / "reference"
UI_DIST = HERE / "ui" / "dist"
for d in (WORK, FILES, OUT, CACHE):
    d.mkdir(exist_ok=True)

import uvicorn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

import build_plan
import llm
import scope_detect

app = FastAPI(title="User Story Agent")
JOB = {"state": "idle", "message": "", "result": None}
_lock = threading.Lock()
EXTRACT_PROC = {"p": None}
CONFIG_FILE = WORK / "config.json"

# slot -> (label, mandatory, multiple files)
SLOTS = {
    "workshop": ("Fit-to-Standard workshop notes / minutes / MoM", False, True),
    "kdd": ("KDD register", True, True),
    "gap": ("Gap / WRICEF list", False, True),
    "trace": ("Requirements traceability matrix", False, True),
    "l4": ("L4 plan", True, False),
    "bpd": ("BPD", False, True),
    "backlog": ("Product backlog", False, True),
    "other": ("Any other input", False, True),
    "template": ("Output template", False, False),
}
DEFAULT_CONFIG = dict(scope_added=[], scope_excluded=[], country_added=[], country_excluded=[], free_text="", output_mode="default", model="")


def read_json(p, default=None):
    try:
        return json.loads(Path(p).read_text(encoding="utf-8"))
    except Exception:
        return default


def config():
    return {**DEFAULT_CONFIG, **(read_json(CONFIG_FILE, {}) or {})}


def parse_ids(text):
    return list(dict.fromkeys(t.upper() for t in re.split(r"[\s,;]+", text or "") if re.fullmatch(r"[0-9A-Za-z]{3}", t)))


def slot_files(slot):
    d = FILES / slot
    return sorted([p for p in d.iterdir() if p.is_file()], key=lambda p: p.stat().st_mtime) if d.exists() else []


def files_info():
    return {k: dict(label=v[0], mandatory=v[1], multiple=v[2],
                    files=[dict(name=p.name, size=p.stat().st_size) for p in slot_files(k)]) for k, v in SLOTS.items()}


_claude = {"t": 0.0, "v": None}


def claude_cached(force=False):
    """llm.status() launches the CLI (slow) - cache it; refresh on demand or after 30 s."""
    if force or _claude["v"] is None or time.time() - _claude["t"] > 30:
        _claude["v"], _claude["t"] = llm.status(), time.time()
    return _claude["v"]


_models = {"list": [], "loading": False, "t": 0.0}


def models_cached():
    """Load the user's available models in the background (starting the CLI takes several seconds)."""
    if not claude_cached()["loggedIn"]:
        _models["list"] = []
    elif (not _models["list"] and time.time() - _models["t"] > 20 or time.time() - _models["t"] > 600) and not _models["loading"]:
        def work():
            try:
                _models["list"] = llm.models() or _models["list"]
            finally:
                _models.update(loading=False, t=time.time())
        _models["loading"] = True
        threading.Thread(target=work, daemon=True).start()
    ms = _models["list"]
    return dict(list=ms, best=llm.strongest(ms), loading=_models["loading"])


_det = {"sig": None, "items": [], "countries": [], "loading": False}


def _files_sig():
    return tuple((k, p.name, p.stat().st_size, int(p.stat().st_mtime)) for k in SLOTS if k != "template" for p in slot_files(k))


def detect_cached():
    """Scope items found in the uploaded files. Re-runs in the background whenever the uploads change."""
    sig = _files_sig()
    if sig != _det["sig"] and not _det["loading"]:
        _det["loading"] = True

        def work():
            try:
                r = scope_detect.detect({k: slot_files(k) for k in SLOTS}, REF / "SSCUI_List.xlsm", CACHE, REF / "countries.json") if sig else {}
            except Exception:
                r = {}
            _det.update(items=r.get("scope", []), countries=r.get("countries", []), sig=sig, loading=False)
        threading.Thread(target=work, daemon=True).start()
    return dict(items=_det["items"], countries=_det["countries"], loading=_det["loading"] or sig != _det["sig"])


def effective_ids(cfg):
    """Detected scope item IDs the user has not deselected + IDs the user added."""
    off = set(cfg.get("scope_excluded") or [])
    return list(dict.fromkeys([d["id"] for d in _det["items"] if d["id"] not in off] + list(cfg.get("scope_added") or [])))


def effective_countries(cfg):
    off = set(cfg.get("country_excluded") or [])
    return list(dict.fromkeys([d["code"] for d in _det["countries"] if d["code"] not in off] + list(cfg.get("country_added") or [])))


_known = {"v": None}


def known_scope():
    if _known["v"] is None:
        _known["v"] = scope_detect.all_known(REF / "SSCUI_List.xlsm", CACHE)
    return _known["v"]


def extract_status():
    st = read_json(CACHE / "extract_status.json", {"state": "never", "message": ""})
    p = EXTRACT_PROC["p"]
    running = bool(p and p.poll() is None)
    if st.get("state") in ("waiting_login", "fetching", "downloading") and not running:
        st = dict(st, state="error", message="The extract process stopped unexpectedly.")
    return dict(status=st, meta=read_json(CACHE / "extract_meta.json"), running=running)


@app.get("/api/state")
def state():
    cfg = config()
    return dict(config=cfg, scope_ids=effective_ids(cfg), country_ids=effective_countries(cfg), known_scope=known_scope(), detected=detect_cached(), files=files_info(), claude=claude_cached(), models=models_cached(), extract=extract_status(),
                job=JOB, countries=read_json(REF / "countries.json", []))


# ---------------------------------------------------------------- Claude login
@app.get("/api/claude/status")
def claude_status():
    return claude_cached(force=True)


@app.post("/api/claude/login")
def claude_login():
    try:
        llm.login()
        return {"started": True, "message": "A browser tab opened - finish signing in to Claude, then press Check."}
    except Exception as e:
        raise HTTPException(500, str(e))


# ---------------------------------------------------------------- config + files
@app.post("/api/config")
def set_config(payload: dict):
    cfg = config()
    for k in DEFAULT_CONFIG:
        if k in payload:
            cfg[k] = payload[k]
    CONFIG_FILE.write_text(json.dumps(cfg), encoding="utf-8")
    return cfg


@app.post("/api/files/{slot}")
async def upload(slot: str, file: UploadFile = File(...)):
    if slot not in SLOTS:
        raise HTTPException(404, "unknown slot")
    d = FILES / slot
    d.mkdir(exist_ok=True)
    if not SLOTS[slot][2]:
        for p in d.iterdir():
            p.unlink()
    name = re.sub(r"[^\w.\- ()]", "_", Path(file.filename).name)
    (d / name).write_bytes(await file.read())
    return files_info()[slot]


@app.delete("/api/files/{slot}/{name}")
def delete_file(slot: str, name: str):
    p = (FILES / slot / name).resolve()
    if slot not in SLOTS or p.parent != (FILES / slot).resolve() or not p.exists():
        raise HTTPException(404, "not found")
    p.unlink()
    return files_info()[slot]


@app.get("/api/template/sample")
def sample_template():
    p = ROOT / "SAP_S4HANA_Public_Cloud_L4_Plan_and_User_Stories_Sample.xlsx"
    if not p.exists():
        raise HTTPException(404, "sample missing")
    return FileResponse(p, filename=p.name)


@app.post("/api/template/analyze")
def analyze_template():
    """Let Claude (or the keyword fallback) read the uploaded output template and show how it will be filled."""
    fs = slot_files("template")
    if not fs:
        raise HTTPException(400, "Upload an output template first.")
    import export_template
    log = []
    use = claude_cached(force=True)["loggedIn"]
    try:
        m, _, cached = export_template.get_mapping(fs[0], use, None, log, force=True)
    except Exception as e:
        raise HTTPException(400, str(e))
    return dict(mapping=m, notes=log, by_claude=use)


# ---------------------------------------------------------------- Process Navigator extract
@app.get("/api/extract/status")
def get_extract_status():
    return extract_status()


@app.post("/api/extract/refresh")
def refresh():
    p = EXTRACT_PROC["p"]
    if p and p.poll() is None:
        raise HTTPException(409, "An extract is already running.")
    cfg = config()
    ids = effective_ids(cfg)
    if not ids:
        raise HTTPException(400, "Select or enter scope item IDs first (the refresh covers the IDs you entered).")
    idf = WORK / "extract_ids.json"
    idf.write_text(json.dumps(ids))
    cmd = [sys.executable, "-I", str(HERE / "extract_process_navigator.py"), "--ids-file", str(idf),
           "--countries", ",".join(effective_countries(cfg) or ["US"]), "--force"]
    EXTRACT_PROC["p"] = subprocess.Popen(cmd, stdin=subprocess.DEVNULL, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    (CACHE / "extract_status.json").write_text(json.dumps(dict(state="waiting_login", message="Starting - an Edge window will open for SAP sign-in.", done=0, total=len(ids))))
    return dict(started=True, scope_items=len(ids))


# ---------------------------------------------------------------- generate
def run_job(cfg_run):
    def say(m):
        JOB["message"] = m
    try:
        res = build_plan.generate(cfg_run, say)
        res["file"] = Path(cfg_run["out"]).name
        JOB.update(state="done", message="Done", result=res)
    except Exception as e:
        JOB.update(state="error", message=f"{type(e).__name__}: {e}", result=dict(trace=traceback.format_exc()[-1500:]))


@app.post("/api/generate")
def generate(payload: dict | None = None):
    payload = payload or {}
    cfg = config()
    ids = effective_ids(cfg)
    problems = []
    if detect_cached()["loading"]:
        raise HTTPException(400, "Still detecting scope items in your files - try again in a moment.")
    if not ids:
        problems.append("scope item IDs")
    countries = effective_countries(cfg)
    if not countries:
        problems.append("at least one country")
    problems += [v[0] for k, v in SLOTS.items() if v[1] and not slot_files(k)]
    if cfg["output_mode"] == "template" and not slot_files("template"):
        problems.append("output template (or choose the sample format)")
    if problems:
        raise HTTPException(400, "Still needed: " + "; ".join(problems))
    ms = models_cached()
    chosen = cfg.get("model") or ""
    model = chosen if chosen in [m["value"] for m in ms["list"]] else ms["best"]   # blank = strongest available
    if model == "default":
        model = None
    tmpl = slot_files("template")[0] if cfg["output_mode"] == "template" else None
    ext = tmpl.suffix if tmpl else ".xlsx"
    name = f"SAP_S4HANA_Public_Cloud_L4_Plan_and_User_Stories_{dt.datetime.now():%Y%m%d_%H%M%S}{ext}"
    run = dict(scope_ids=ids, countries=countries, free_text=cfg["free_text"], use_llm=payload.get("use_claude", True), model=model,
               files={k: [str(p) for p in slot_files(k)] for k in SLOTS if k != "template"},
               l4=str(slot_files("l4")[0]), sscui=str(REF / "SSCUI_List.xlsm"), out=str(OUT / name),
               output=dict(mode="template", template=str(tmpl)) if tmpl else dict(mode="default"))
    with _lock:
        if JOB["state"] == "running":
            raise HTTPException(409, "A generation is already running.")
        JOB.update(state="running", message="Starting", result=None)
    threading.Thread(target=run_job, args=(run,), daemon=True).start()
    return {"started": True}


@app.get("/api/generate/status")
def generate_status():
    return JOB


@app.get("/api/download/{name}")
def download(name: str):
    p = (OUT / name).resolve()
    if p.parent != OUT.resolve() or not p.exists():
        raise HTTPException(404, "not found")
    return FileResponse(p, filename=p.name)


if UI_DIST.exists():
    app.mount("/", StaticFiles(directory=UI_DIST, html=True), name="ui")

def reset_workspace():
    """Every launch starts blank: no scope IDs, countries, uploads, free text or stale status messages."""
    CONFIG_FILE.unlink(missing_ok=True)
    shutil.rmtree(FILES, ignore_errors=True)
    FILES.mkdir(exist_ok=True)
    for f in ("extract_status.json", "extract_ids.json"):
        (CACHE / f).unlink(missing_ok=True)
        (WORK / f).unlink(missing_ok=True)


if __name__ == "__main__":
    reset_workspace()
    import os
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", 8000))
    print(f"User Story Agent running at http://127.0.0.1:{port}  (Ctrl+C to stop)")
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")
