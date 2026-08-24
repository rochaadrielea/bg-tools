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
import validate as val   # noqa: E402

DB = build_sas.DEFAULT_DB

def _version() -> str:
    """Whatever deploy.sh last stamped, so the page can show what is live."""
    try:
        return "v" + (ROOT / "VERSION").read_text().strip()
    except Exception:
        return "unversioned"
app = FastAPI(title="SAS NC Dashboard", root_path="/sas")


# ----------------------------- data access -----------------------------------
def _read_all() -> pd.DataFrame:
    con = build_sas.connect(DB)
    try:
        try:
            df = pd.read_sql("SELECT * FROM sas_nc", con)
        except Exception:
            df = pd.DataFrame()
    finally:
        con.close()
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
        # Display placeholders (from _clean_cat in charts.py) map to blank/nan
        # in the real data, not a literal string. Match emptiness for those so
        # clicking a "(not recorded)" slice returns the blank rows, not zero.
        PLACEHOLDERS = {"(not recorded)", "(blank)", "(not set)", "(none)",
                        "Unassigned", "Not recorded", ""}
        s = d[col].astype(str).str.strip()
        if str(val).strip() in PLACEHOLDERS:
            d = d[d[col].isna() | s.isin(["", "nan", "None", "-", "NaT"]) |
                  s.isin(PLACEHOLDERS)]
        else:
            d = d[s == str(val)]
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
    return {"version": _version(), "rows": n, "last_import": (
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
    d["opened"] = pd.to_datetime(d["opened"], errors="coerce")
    d = d.sort_values("opened", na_position="last")
    return {"cols": charts.RAW_COLS, "labels": charts.RAW_LABELS,
            "rows": charts._raw_rows(d), "count": int(len(d))}


@app.post("/api/export")
async def export(payload: dict):
    """Per-chart download. Same 22-column ingest format so it can be loaded
    back with Add/update. Plus a Source system column."""
    df = _read_all()
    d = _apply_filters(df, payload.get("filters", {}))
    scope = payload.get("scope") or {}
    for k, v in scope.items():
        if k in d.columns:
            d = d[d[k].astype(str) == str(v)]
    d = _apply_click(d, payload.get("click", {}))
    d["opened"] = pd.to_datetime(d["opened"], errors="coerce")
    d = d.sort_values("opened", na_position="last")

    # write in the full ingest shape so it round-trips
    inv = {v: k for k, v in parse.COLS.items()}
    out = pd.DataFrame()
    for internal, header in inv.items():
        if internal == "project":
            out[header] = parse.PROJECT
        elif internal == "status_raw":
            out[header] = d["status"].values if "status" in d.columns else None
        elif internal == "copq_raw":
            out[header] = (-pd.to_numeric(d["copq"], errors="coerce").fillna(0)
                           ).values if "copq" in d.columns else None
        elif internal == "copq_wbs":
            out[header] = ["-" if not b else "booked"
                           for b in pd.to_numeric(
                               d.get("copq_booked", 0), errors="coerce").fillna(0)] \
                          if "copq_booked" in d.columns else None
        elif internal == "batch_sap":
            out[header] = d["batch_sap"].values if "batch_sap" in d.columns else None
        elif internal in d.columns:
            out[header] = d[internal].values
        else:
            out[header] = None
    out["Plant"] = "Emmen"
    out["Source system"] = d.get("source_system",
        ["Teamcenter" if str(n).startswith("IR-") else "SAP"
         for n in d["notification"]]).values

    for c in ("Notification Date", "Closing date"):
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="SAS export")
    buf.seek(0)

    click = payload.get("click") or {}
    bit = "_" + str(click.get("value", "")).replace(" ", "_")[:30] if click.get("value") else ""
    fn = f"SAS_NCs{bit}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return StreamingResponse(
        buf, media_type="application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fn}"'})


@app.post("/api/full_export")
async def full_export(payload: dict):
    """Everything currently in the dashboard, in the exact 22-column shape the
    ingest reads. Edit it in Excel (add rows, remove rows, fix values) and load
    it straight back with 'Load export'. A Data quality sheet lists what is
    missing per column so the gaps are visible before anyone quotes a number."""
    df = _read_all()
    if df.empty:
        return JSONResponse({"ok": False, "error": "no data loaded"}, status_code=400)
    d = _apply_filters(df, payload.get("filters", {}))
    inv = {v: k for k, v in parse.COLS.items()}          # internal -> export header
    out = pd.DataFrame()
    for internal, header in inv.items():
        if internal == "project":
            out[header] = [parse.PROJECT] * len(d)
        elif internal == "status_raw":
            # stored as 'status' after parsing - without this the status column
            # comes out EMPTY and a re-upload wipes every status
            out[header] = d["status"].values if "status" in d.columns else None
        elif internal == "copq_raw":
            # stored as a positive magnitude; SAP books it negative, so write it
            # back the way SAP wrote it or the cost is lost on re-upload
            out[header] = (-pd.to_numeric(d["copq"], errors="coerce").fillna(0)
                           ).values if "copq" in d.columns else None
        elif internal == "copq_wbs":
            # '-' means nothing was booked; anything else means a cost element exists
            out[header] = ["-" if not b else "booked"
                           for b in pd.to_numeric(
                               d.get("copq_booked", 0), errors="coerce").fillna(0)] \
                          if "copq_booked" in d.columns else None
        elif internal == "batch_sap":
            out[header] = d["batch_sap"].values if "batch_sap" in d.columns else None
        elif internal in d.columns:
            out[header] = d[internal].values
        else:
            out[header] = None
    for c in ("opened", "closed"):
        h = inv.get(c)
        if h:
            out[h] = pd.to_datetime(d[c], errors="coerce").dt.strftime("%Y-%m-%d")
    if "source_system" in d.columns:
        out["Source system"] = d["source_system"].values

    # data quality: what is missing, per column
    dq = []
    for c in out.columns:
        v = out[c]
        filled = v.notna() & v.astype(str).str.strip().ne("") & \
            v.astype(str).str.strip().ne("-")
        dq.append({"Column": c, "Rows": len(out), "Filled": int(filled.sum()),
                   "Missing": int(len(out) - filled.sum()),
                   "Filled %": round(100 * filled.sum() / max(1, len(out)), 1)})

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="SAS export")
        pd.DataFrame(dq).to_excel(w, index=False, sheet_name="Data quality")
        pd.DataFrame([
            ["How to use this file"],
            ["1. Edit the 'SAS export' sheet - add rows, delete rows, correct values."],
            ["2. Keep the column headers EXACTLY as they are (note the double space "
             "in 'Project Text  (Notification)')."],
            ["3. Load it back with the dashboard's 'Load export' button."],
            ["4. 'Replace all' overwrites everything; 'Add / update' merges on the "
             "Notification number."],
            [],
            ["Derived by the dashboard, do not add columns for them:"],
            ["Batch", "from the 3-digit tail of WBS Element (Notification)"],
            ["Supplier vs Production", "Type Z2, or Cause = Supplier"],
            ["Minor / Major", "Defect Class 0-3 = Minor, 4-5 = Major"],
            ["Cost", "CoPQ is negative in SAP; shown as a positive cost"],
        ]).to_excel(w, index=False, header=False, sheet_name="Read me")
    buf.seek(0)
    fn = f"SAS_full_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return StreamingResponse(
        buf, media_type="application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fn}"'})


@app.post("/api/upload")
async def upload(file: UploadFile = File(...), mode: str = Form("append")):
    raw_bytes = await file.read()
    try:
        # read the right sheet
        import openpyxl, io as _io
        wb = openpyxl.load_workbook(_io.BytesIO(raw_bytes), read_only=True)
        sheet = "SAS export" if "SAS export" in wb.sheetnames else None
        wb.close()

        df_check = pd.read_excel(
            io.BytesIO(raw_bytes),
            sheet_name=sheet if sheet else 0)

        # current count for the partial-file guard
        con = build_sas.connect(DB)
        try:
            current = con.execute("SELECT COUNT(*) FROM sas_nc").fetchone()[0]
        except Exception:
            current = 0

        # validate BEFORE writing
        findings = val.validate(df_check, current_count=current)
        reds = val.reds(findings)

        # partial file + Start over guard
        if mode == "rebuild" and current > 0 and len(df_check) < current * 0.5:
            findings.append(("RED", None, None, None,
                f"This file has {len(df_check)} rows but the dashboard holds "
                f"{current}. Use 'Add to what is already here' instead, or "
                f"load the full report."))
            reds = val.reds(findings)

        if reds:
            con.close()
            return JSONResponse({
                "ok": False,
                "blocked": True,
                "error": val.summary(findings),
                "findings": [{"sev": s, "row": r, "nc": nc, "field": f, "msg": m}
                             for s, r, nc, f, m in findings],
            }, status_code=400)

        # ingest
        result = ingest_sas.ingest(
            con, io.BytesIO(raw_bytes), file.filename,
            rebuild=(mode == "rebuild"),
            current_count=current)
        con.close()

        if not result["ok"]:
            return JSONResponse({
                "ok": False,
                "blocked": True,
                "error": result["error"],
                "findings": [{"sev": s, "row": r, "nc": nc, "field": f, "msg": m}
                             for s, r, nc, f, m in result["findings"]],
            }, status_code=400)

        return {
            "ok": True,
            "rows": result["rows"],
            "mode": mode,
            "file": file.filename,
            "warnings": [{"sev": s, "row": r, "nc": nc, "field": f, "msg": m}
                         for s, r, nc, f, m in result["findings"]
                         if s != "INFO"],
            "new_refs": result.get("new_refs", {}),
        }
    except Exception as e:
        import traceback; traceback.print_exc()
        return JSONResponse({"ok": False, "error": str(e)}, status_code=400)


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