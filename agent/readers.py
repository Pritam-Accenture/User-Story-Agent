"""Read text out of user attachments in (almost) any format.

Supported: docx, pptx, pdf, xlsx/xlsm/csv/tsv, txt/md/json/xml/html/eml/rtf-as-text.
Unreadable or unsupported files return a note instead of failing the run.
"""
import csv
import html
import io
import re
from pathlib import Path

MAX_CHARS = 400_000


def _clean(t):
    return re.sub(r"\n{3,}", "\n\n", re.sub(r"[ \t]+", " ", t or "")).strip()


def _docx(p):
    import docx
    d = docx.Document(str(p))
    out = [para.text for para in d.paragraphs if para.text.strip()]
    for t in d.tables:
        for row in t.rows:
            out.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
    return "\n".join(out)


def _pptx(p):
    from pptx import Presentation
    out = []
    for i, s in enumerate(Presentation(str(p)).slides, 1):
        out.append(f"--- Slide {i} ---")
        for sh in s.shapes:
            if sh.has_text_frame:
                out.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    out.append(" | ".join(c.text.strip() for c in row.cells))
        if s.has_notes_slide:
            out.append("Notes: " + s.notes_slide.notes_text_frame.text)
    return "\n".join(out)


def _pdf(p):
    from pypdf import PdfReader
    return "\n".join((pg.extract_text() or "") for pg in PdfReader(str(p)).pages)


def table_rows(p):
    """Return [(sheet_name, [row tuples])] for xlsx/xlsm/csv/tsv."""
    p = Path(p)
    if p.suffix.lower() in (".csv", ".tsv"):
        raw = p.read_bytes().decode("utf-8-sig", errors="replace")
        dlm = "\t" if p.suffix.lower() == ".tsv" else (";" if raw.count(";") > raw.count(",") else ",")
        return [(p.stem, [tuple(r) for r in csv.reader(io.StringIO(raw), delimiter=dlm)])]
    import openpyxl
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    out = [(ws.title, [tuple(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
    wb.close()
    return out


def _table_text(p):
    out = []
    for name, rows in table_rows(p):
        out.append(f"--- Sheet {name} ---")
        for r in rows:
            if any(c not in (None, "") for c in r):
                out.append(" | ".join("" if c is None else str(c).replace("\n", " ") for c in r))
    return "\n".join(out)


def read_text(path):
    """Return (text, note). note is '' on success."""
    p = Path(path)
    ext = p.suffix.lower()
    try:
        if ext == ".docx": t = _docx(p)
        elif ext == ".pptx": t = _pptx(p)
        elif ext == ".pdf": t = _pdf(p)
        elif ext in (".xlsx", ".xlsm", ".csv", ".tsv"): t = _table_text(p)
        elif ext in (".txt", ".md", ".json", ".xml", ".html", ".htm", ".eml", ".rtf", ".log", ""):
            t = p.read_bytes().decode("utf-8", errors="replace")
            if ext in (".html", ".htm", ".xml"):
                t = html.unescape(re.sub(r"<[^>]+>", " ", t))
        elif ext in (".doc", ".ppt", ".xls"):
            return "", f"{p.name}: old Office format ({ext}) - please save as {ext}x and upload again."
        else:
            return "", f"{p.name}: format {ext} not supported for reading."
    except Exception as e:
        return "", f"{p.name}: could not be read ({type(e).__name__}: {str(e)[:80]})."
    t = _clean(t)
    if not t:
        return "", f"{p.name}: no readable text found (scanned image?)."
    return t[:MAX_CHARS], ""


def chunks(text, size=12000):
    """Split on paragraph boundaries into pieces of about `size` characters."""
    out, cur = [], ""
    for para in text.split("\n"):
        if len(cur) + len(para) + 1 > size and cur:
            out.append(cur); cur = ""
        cur += para + "\n"
    if cur.strip():
        out.append(cur)
    return out
