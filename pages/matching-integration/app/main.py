"""
matching-integration dashboard — FastAPI app (same shape as the SAS dashboard).

Endpoints:
  GET  /                 -> the page
  GET  /plotly.min.js    -> the vendored Plotly bundle the page loads
  GET  /api/status       -> row count, last import
  POST /api/data         -> filtered chart payload
  POST /api/upload       -> load a raw workbook (Add / Start Over)
  POST /api/raw          -> the rows behind a clicked bar / cell / slice
  POST /api/raw_export   -> the same rows as an .xlsx download
  POST /api/feedback     -> a note from whoever is using the dashboard
  GET  /api/report       -> full report download

Everything the user sees is Integration + Matching only; charts.scope() is
applied inside charts._apply_filters, so the raw endpoints inherit it.
"""
from __future__ import annotations
import io
import re
import sqlite3
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import (HTMLResponse, StreamingResponse, FileResponse)

import parse
import charts

APP_DIR = Path(__file__).resolve().parent
DB = APP_DIR.parent / "dash" / "mi.db"
DB.parent.mkdir(parents=True, exist_ok=True)

app = FastAPI()

# The raw table: which columns, in which order, under which heading.
RAW_COLS = [
    ("notification", "NC"),
    ("title", "Title"),
    ("notif_type", "Type"),
    ("area", "Area"),
    ("project", "Project"),
    ("status", "Status"),
    ("opened", "Opened"),
    ("closed", "Closed"),
    ("leadtime", "Lead (d)"),
    ("disposition", "Disposition"),
    ("defect_class", "Class"),
    ("rc1", "Root cause L1"),
    ("rc2", "Root cause L2"),
    ("origin1", "Detection L1"),
    ("origin2", "Detection L2"),
    ("cause", "Cause"),
    ("material", "Material"),
    ("copq", "CoPQ (CHF)"),
]

PLACEHOLDERS = {"(blank)", "(not coded)", "(not recorded)", "Unassigned", ""}


# ---------------- storage ----------------

def _conn():
    c = sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS mi_nc (
        notification TEXT, notif_type TEXT, area TEXT, project TEXT, status TEXT,
        defect_class TEXT, disposition TEXT, cause TEXT, material TEXT,
        opened TEXT, closed TEXT, leadtime REAL, copq REAL,
        rc1 TEXT, rc2 TEXT, origin1 TEXT, origin2 TEXT, title TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS mi_import (
        ts TEXT, filename TEXT, rows INTEGER)""")
    c.execute("""CREATE TABLE IF NOT EXISTS feedback (
        ts TEXT, name TEXT, category TEXT, message TEXT, attachment TEXT)""")
    return c


def _load_df() -> pd.DataFrame:
    c = _conn()
    try:
        df = pd.read_sql("SELECT * FROM mi_nc", c)
    except Exception:
        df = pd.DataFrame(columns=parse.COLS)
    c.close()
    if not df.empty:
        for col in ("opened", "closed"):
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def _store_df(df: pd.DataFrame, filename: str, replace: bool):
    c = _conn()
    if replace:
        c.execute("DELETE FROM mi_nc")
        c.execute("DELETE FROM mi_import")
    out = df.copy()
    for col in ("opened", "closed"):
        out[col] = out[col].astype(str)
    out.to_sql("mi_nc", c, if_exists="append", index=False)
    c.execute("INSERT INTO mi_import (ts, filename, rows) VALUES (datetime('now'), ?, ?)",
              (filename, len(df)))
    c.commit()
    c.close()


# ---------------- slicing a click ----------------

def _match_col(df: pd.DataFrame, col: str, val) -> pd.DataFrame:
    """One column equals one value, with blanks matching the placeholders the
    charts display."""
    if not col or col not in df.columns:
        return df
    s = df[col].astype(str).str.strip()
    if str(val).strip() in PLACEHOLDERS:
        return df[df[col].isna() | s.isin(["", "nan", "None", "NaT", "-"])]
    return df[s == str(val)]


def _slice(df: pd.DataFrame, payload: dict) -> pd.DataFrame:
    """Apply everything the front end sent about the clicked point:
      click  {col, value}          one category
      scope  {col: value, ...}     the drill path / the other heat-map axis
      month  {col, value}          '2026-05' on opened or closed
      lead   [lo, hi]              lead-time bucket, hi may be null
      side   'Supplier'|'Production'   Z2 vs Z3
    """
    d = df
    click = payload.get("click") or {}
    if click:
        d = _match_col(d, click.get("col"), click.get("value"))

    for k, v in (payload.get("scope") or {}).items():
        if k == "__month_opened":
            d = _month(d, "opened", v)
        elif k == "__month_closed":
            d = _month(d, "closed", v)
        else:
            d = _match_col(d, k, v)

    m = payload.get("month") or {}
    if m.get("value"):
        d = _month(d, m.get("col", "opened"), m["value"])

    lead = payload.get("lead")
    if lead and isinstance(lead, (list, tuple)) and len(lead) == 2:
        s = pd.to_numeric(d["leadtime"], errors="coerce")
        lo, hi = lead
        if lo is not None:
            d = d[s >= float(lo)]
            s = pd.to_numeric(d["leadtime"], errors="coerce")
        if hi is not None:
            d = d[s < float(hi)]

    side = payload.get("side")
    if side in ("Supplier", "Production"):
        pref = "Z2" if side == "Supplier" else "Z3"
        d = d[d["notif_type"].astype(str).str.strip().str.startswith(pref)]

    return d


def _month(d: pd.DataFrame, col: str, val: str) -> pd.DataFrame:
    if col not in d.columns or not val:
        return d
    s = pd.to_datetime(d[col], errors="coerce").dt.to_period("M").astype(str)
    return d[s == str(val)]


def _rows(df: pd.DataFrame) -> dict:
    cols = [(k, lab) for k, lab in RAW_COLS if k in df.columns]
    keys = [k for k, _ in cols]
    out = df[keys].copy()
    for c in ("opened", "closed"):
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce").dt.strftime("%Y-%m-%d")
    if "copq" in out.columns:
        out["copq"] = pd.to_numeric(out["copq"], errors="coerce").round(0)
    if "leadtime" in out.columns:
        out["leadtime"] = pd.to_numeric(out["leadtime"], errors="coerce")
    if "notification" in out.columns:
        out = out.sort_values("notification")
    out = out.replace({np.nan: None})
    recs = out.to_dict("records")
    for r in recs:
        for k, v in list(r.items()):
            if isinstance(v, float):
                if v != v:              # NaN
                    r[k] = None
                elif v == int(v):
                    r[k] = int(v)
    return {"cols": keys, "labels": {k: lab for k, lab in cols},
            "rows": recs, "count": len(out),
            "numeric": [k for k in ("leadtime", "copq") if k in out.columns]}


# ---------------- routes ----------------

@app.get("/", response_class=HTMLResponse)
def index():
    return (APP_DIR / "page.html").read_text(encoding="utf-8")


@app.get("/plotly.min.js")
def plotly_js():
    """Serve the vendored Plotly bundle that sits next to this file.
    Without this route the page loads but every chart area stays blank."""
    return FileResponse(APP_DIR / "plotly.min.js",
                        media_type="application/javascript")


@app.get("/api/status")
def status():
    df = _load_df()
    c = _conn()
    imp = c.execute(
        "SELECT ts, filename, rows FROM mi_import ORDER BY ts DESC LIMIT 1"
    ).fetchone()
    c.close()
    return {
        "rows": len(charts.scope(df)),
        "rows_imported": len(df),
        "last_import": ({"ts": imp[0], "filename": imp[1], "rows": imp[2]}
                        if imp else None),
        "version": "mi v0.2",
    }


@app.post("/api/data")
async def data(payload: dict):
    df = _load_df()
    if df.empty:
        return {"empty": True}
    return charts.build_payload(df, payload.get("filters") or {})


@app.post("/api/upload")
async def upload(file: UploadFile = File(...), mode: str = Form("add")):
    raw = await file.read()
    tmp = APP_DIR.parent / "dash" / ("_upload_" + file.filename)
    tmp.write_bytes(raw)
    try:
        df = parse.load(str(tmp))
    finally:
        try:
            tmp.unlink()
        except Exception:
            pass
    _store_df(df, file.filename, replace=(mode == "replace"))
    total = len(charts.scope(_load_df()))
    return {"ok": True, "added": len(df), "total": total}


@app.post("/api/raw")
async def raw(payload: dict):
    df = _load_df()
    if df.empty:
        return {"rows": [], "count": 0, "cols": [], "labels": {}}
    d = charts._apply_filters(df, payload.get("filters") or {})
    return _rows(_slice(d, payload))


@app.post("/api/raw_export")
async def raw_export(payload: dict):
    df = _load_df()
    d = charts._apply_filters(df, payload.get("filters") or {})
    d = _slice(d, payload)
    r = _rows(d)
    out = pd.DataFrame(r["rows"], columns=r["cols"])
    out.columns = [r["labels"][c] for c in r["cols"]]
    name = payload.get("name") or "selection"
    name = re.sub(r"[^\w\-]+", "_", str(name))[:60] or "selection"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        out.to_excel(xl, index=False, sheet_name="NCs")
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=MI_{name}.xlsx"},
    )


@app.post("/api/feedback")
async def feedback(message: str = Form(""), name: str = Form(""),
                   category: str = Form(""), file: UploadFile = File(None)):
    """Every tool with a UI takes feedback. Notes land in the same database as
    the NCs, in a `feedback` table; attachments sit next to it in
    feedback_attachments/. Nobody using the dashboard can read other people's
    notes - submit only."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    saved = ""
    if file is not None and getattr(file, "filename", ""):
        folder = DB.parent / "feedback_attachments"
        folder.mkdir(parents=True, exist_ok=True)
        who = re.sub(r"[^\w\-]+", "_", (name or "anon"))[:20]
        safe = re.sub(r"[^\w\-.]+", "_", file.filename)[:80]
        saved = f"{stamp}_{who}_{safe}"
        (folder / saved).write_bytes(await file.read())
    c = _conn()
    c.execute("INSERT INTO feedback (ts, name, category, message, attachment) "
              "VALUES (datetime('now'), ?, ?, ?, ?)",
              (name or "", category or "", message or "", saved))
    c.commit()
    c.close()
    return {"ok": True}


@app.get("/api/report")
def report():
    df = charts.scope(_load_df())
    r = _rows(df)
    out = pd.DataFrame(r["rows"], columns=r["cols"])
    if len(r["cols"]):
        out.columns = [r["labels"][c] for c in r["cols"]]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        out.to_excel(xl, index=False, sheet_name="MI report")
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=MI_full_report.xlsx"},
    )