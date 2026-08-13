"""
dash_sas FastAPI backend.  Served at /sas behind nginx.

Endpoints
  GET  /sas/                 -> the page (page.html)
  GET  /sas/plotly.min.js    -> vendored Plotly (offline / air-gapped)
  POST /sas/api/data         -> {filters} -> full chart payload for the filtered set
  POST /sas/api/raw          -> {filters, click} -> raw rows for a clicked slice
  POST /sas/api/export       -> {filters, click} -> .xlsx of the current raw rows
  POST /sas/api/upload       -> multipart export file -> rebuild sas_nc
  POST /sas/api/feedback     -> store a note (+ optional file) in quality.db
  GET  /sas/api/status       -> last import info + row count
"""
from __future__ import annotations
import io
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import (HTMLResponse, JSONResponse, StreamingResponse,
                               FileResponse, Response)

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))

import parse            # noqa: E402
import build_sas        # noqa: E402
import charts           # noqa: E402
import ingest_sas       # noqa: E402

DB = build_sas.DEFAULT_DB
app = FastAPI(title="SAS NC Dashboard", root_path="/sas")


# ----------------------------- data access -----------------------------------
def _tracker_only_sas(con, existing: pd.DataFrame) -> pd.DataFrame:
    """SAS NCs that live only in the NC tracker (later batches, new PSP).

    The SAP export stops around Batch 19; the tracker already has Batch 21
    (W.IC248.Q.021) and will keep picking up successor WBS elements.
    """
    try:
        tr = pd.read_sql(
            "SELECT nc_id, sap_notif, nc_type, is_supplier_nc, detection_area, "
            "status, created_on, closure_date, material, psp_ref, nc_wbs, "
            "classification, defect_code_text, copq, leadtime, supplier_name, "
            "notif_year FROM nc WHERE project = 'SAS'", con)
    except Exception:
        return pd.DataFrame()
    if tr.empty:
        return pd.DataFrame()
    have = set()
    if existing is not None and not existing.empty and "notification" in existing.columns:
        have = set(existing["notification"].astype(str))
    tr["notification"] = tr["sap_notif"].fillna(tr["nc_id"]).astype(str)
    tr = tr[~tr["notification"].isin(have)].copy()
    if tr.empty:
        return pd.DataFrame()

    is_sup = tr["is_supplier_nc"].fillna(0).astype(int).eq(1) | \
        tr["nc_type"].astype(str).str.lower().str.startswith("supplier")
    opened = pd.to_datetime(tr["created_on"], errors="coerce")
    closed = pd.to_datetime(tr["closure_date"], errors="coerce")
    wbs = tr["psp_ref"].fillna(tr["nc_wbs"])
    cause = tr["detection_area"].map(_tracker_cause)
    rows = pd.DataFrame({
        "notification": tr["notification"],
        "notif_type": ["Z2 - Procurem. Complaint" if s else "Z3 - Production NC"
                       for s in is_sup],
        "notif_year": pd.to_numeric(tr["notif_year"], errors="coerce")
                      .fillna(opened.dt.year).astype("Int64"),
        "status": tr["status"].map(parse.status_label),
        "wbs": wbs.astype(str),
        "wbs_text": "",
        "batch": wbs.map(parse.batch_label),
        "defect_class": tr["classification"].astype(str),
        "defect_class_label": tr["classification"].map(_tracker_class),
        "defect_code": tr["defect_code_text"],
        "disposition": "Not assigned",
        "cause": cause,
        "notif_text": "",
        "material": tr["material"],
        "model": "",
        "vendor": tr["supplier_name"],
        "vendor_clean": tr["supplier_name"].map(parse.vendor_norm),
        "opened": opened,
        "closed": closed,
        "month": opened.dt.strftime("%Y-%m"),
        "leadtime": pd.to_numeric(tr["leadtime"], errors="coerce"),
        "copq": pd.to_numeric(tr["copq"], errors="coerce").fillna(0).abs(),
        "copq_booked": 0,
    })
    return rows.reset_index(drop=True)


def _tracker_cause(v) -> str:
    s = str(v or "").strip()
    low = s.lower()
    if low.startswith("s:") or low == "supplier":
        return "Supplier"
    if "customer complaint" in low or low.startswith("cc:"):
        return "Customer complaint"
    if low.startswith("ii:") or "incoming" in low:
        return "Incoming inspection"
    if "manufactur" in low:
        return "Manufacturing"
    if low.startswith("fi:") or "inspection" in low:
        return "Inspection"
    if "assembl" in low or low.startswith("i:"):
        return "Assembly"
    return s or "Not assigned"


def _tracker_class(v) -> str:
    s = str(v or "").strip()
    if s.lower().startswith("major"):
        return "Major"
    if s.lower().startswith("minor"):
        return "Minor"
    return parse.class_label(s)


def _read_all() -> pd.DataFrame:
    con = build_sas.connect(DB)
    try:
        try:
            df = pd.read_sql("SELECT * FROM sas_nc", con)
        except Exception:
            df = pd.DataFrame()
        extra = _tracker_only_sas(con, df)
    finally:
        con.close()
    if extra is not None and not extra.empty:
        df = pd.concat([df, extra], ignore_index=True) if not df.empty else extra
    if df.empty:
        return df
    return parse.tag_lanes(df)


def _apply_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    if df.empty or not f:
        return df
    d = df.copy()
    if f.get("date_from"):
        d = d[d["opened"] >= f["date_from"]]
    if f.get("date_to"):
        d = d[d["opened"] <= f["date_to"]]
    for col, key in (("notif_type", "types"), ("status", "statuses"),
                     ("batch", "batches"), ("disposition", "dispositions"),
                     ("defect_class_label", "classes"),
                     ("vendor_clean", "vendors"), ("defect_code", "defect_codes")):
        vals = f.get(key)
        if vals:
            d = d[d[col].isin(vals)]
    if f.get("months"):
        d = d[d["month"].isin(f["months"])]
    return d


def _apply_click(df: pd.DataFrame, click: dict) -> pd.DataFrame:
    """click = {col: value}. age_bucket / notification handled specially."""
    if not click:
        return df
    d = df.copy()
    col = click.get("col")
    val = click.get("value")
    if not col:
        return d
    if col == "age_bucket":
        d["opened_dt"] = pd.to_datetime(d["opened"], errors="coerce")
        today = pd.Timestamp.today().normalize()
        age = (today - d["opened_dt"]).dt.days
        d = d[d["status"] == "Open"]
        age = age.loc[d.index]
        rng = {"0-30": (0, 30), "31-60": (31, 60),
               "61-90": (61, 90), "90+": (91, 10**9)}.get(val)
        if rng:
            d = d[(age >= rng[0]) & (age <= rng[1])]
        return d.drop(columns=["opened_dt"], errors="ignore")
    if col in d.columns:
        d = d[d[col].astype(str) == str(val)]
    return d


# ------------------------------- routes ---------------------------------------
@app.get("/", response_class=HTMLResponse)
def index():
    html = (HERE / "page.html").read_text(encoding="utf-8")
    return Response(content=html.encode("utf-8"), media_type="text/html")


@app.get("/favicon.ico")
def favicon():
    return Response(status_code=204)


@app.get("/plotly.min.js")
def plotly():
    return FileResponse(HERE / "plotly.min.js", media_type="application/javascript")


@app.get("/api/status")
def status():
    con = build_sas.connect(DB)
    try:
        try:
            n = con.execute("SELECT COUNT(*) FROM sas_nc").fetchone()[0]
        except Exception:
            n = 0
        try:
            last = con.execute(
                "SELECT ts, file, rows, mode FROM sas_import "
                "ORDER BY id DESC LIMIT 1").fetchone()
        except Exception:
            last = None
    finally:
        con.close()
    return {"rows": n, "last_import": (
        {"ts": last[0], "file": last[1], "rows": last[2], "mode": last[3]}
        if last else None)}


@app.post("/api/data")
async def data(payload: dict):
    df = _read_all()
    if df.empty:
        return JSONResponse({"empty": True})
    d = _apply_filters(df, payload.get("filters", {}))
    out = charts.build_payload(d)
    out["empty"] = False
    out["filtered"] = int(len(d))
    out["all"] = int(len(df))
    # option lists for the filter controls (from the full set)
    out["options"] = {
        "types": sorted(df["notif_type"].dropna().unique().tolist()),
        "statuses": ["Closed", "Open", "Deleted"],
        "batches": sorted(df["batch"].dropna().unique().tolist(),
                          key=parse.batch_sort_key),
        "dispositions": sorted(df["disposition"].dropna().unique().tolist()),
        "classes": ["Minor", "Major", "Unclassified"],
        "vendors": list(df[df["lane"] == "Supplier"]
                        ["vendor_clean"].value_counts().index),
        "defect_codes": sorted(df["defect_code"].dropna().unique().tolist()),
        "date_min": str(pd.to_datetime(df["opened"], errors="coerce").min().date()),
        "date_max": str(pd.to_datetime(df["opened"], errors="coerce").max().date()),
    }
    return out


@app.post("/api/raw")
async def raw(payload: dict):
    df = _read_all()
    if df.empty:
        return {"cols": charts.RAW_COLS, "labels": charts.RAW_LABELS, "rows": []}
    d = _apply_filters(df, payload.get("filters", {}))
    scope = payload.get("scope") or {}
    for k, v in scope.items():
        if k in d.columns:
            d = d[d[k].astype(str) == str(v)]
    d = _apply_click(d, payload.get("click", {}))
    d = d.sort_values("opened", na_position="last")
    return {"cols": charts.RAW_COLS, "labels": charts.RAW_LABELS,
            "rows": charts._raw_rows(d), "count": int(len(d))}


@app.post("/api/export")
async def export(payload: dict):
    df = _read_all()
    d = _apply_filters(df, payload.get("filters", {}))
    scope = payload.get("scope") or {}
    for k, v in scope.items():
        if k in d.columns:
            d = d[d[k].astype(str) == str(v)]
    d = _apply_click(d, payload.get("click", {}))
    d = d.sort_values("opened", na_position="last")
    cols = charts.RAW_COLS
    out = d[cols].rename(columns=charts.RAW_LABELS)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="SAS NCs")
    buf.seek(0)
    fn = f"SAS_NCs_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return StreamingResponse(
        buf, media_type="application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fn}"'})


@app.post("/api/upload")
async def upload(file: UploadFile = File(...), mode: str = Form("rebuild")):
    raw_bytes = await file.read()
    try:
        con = build_sas.connect(DB)
        n = ingest_sas.ingest(con, io.BytesIO(raw_bytes), file.filename,
                              rebuild=(mode != "append"))
        con.close()
    except Exception as e:  # noqa: BLE001
        return JSONResponse({"ok": False, "error": str(e)}, status_code=400)
    return {"ok": True, "rows": n, "mode": mode, "file": file.filename}


@app.post("/api/feedback")
async def feedback(message: str = Form(...), name: str = Form(""),
                   category: str = Form("Idea"), file: UploadFile = File(None)):
    con = build_sas.connect(DB)
    con.execute(
        "CREATE TABLE IF NOT EXISTS feedback (id INTEGER PRIMARY KEY AUTOINCREMENT,"
        " ts TEXT, app TEXT, version TEXT, category TEXT, message TEXT,"
        " user TEXT, machine TEXT, context TEXT)")
    attach = ""
    if file is not None and file.filename:
        d = DB.parent / "feedback_attachments"
        d.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = d / f"{stamp}_{name or 'anon'}_{file.filename}"
        dest.write_bytes(await file.read())
        attach = dest.name
    con.execute(
        "INSERT INTO feedback (ts, app, version, category, message, user, machine,"
        " context) VALUES (?,?,?,?,?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"), "SAS Dashboard (web)",
         "web-2026-08", category, message, name or "web", "chbs4212",
         f'{{"submitted_by": "{name}", "attachment": "{attach}"}}'))
    con.commit()
    con.close()
    return {"ok": True}