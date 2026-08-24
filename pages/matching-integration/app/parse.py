"""
matching-integration dashboard — data layer.

Reads a raw SAP NC export and (optionally) a raw CAPA export, joins the four
root-cause columns from CAPA onto the NCs by NC number, and returns a tidy
DataFrame the charts can use.

Two ways data arrives:
  1. A single "MI export" file already produced by joining (has the RC columns).
  2. Raw SAP + raw CAPA in the same workbook (Raw_Data_SAP + Raw_Data_Capa
     sheets) — parse joins them here (option A: join lives in parse.py).

Scope: Area I: / M: by default is NOT enforced here — the Area filter does that
in the UI. parse keeps every row so the filter can widen to all areas.
"""
from __future__ import annotations
import pandas as pd
import numpy as np


# ---- the columns the dashboard works with (the tidy schema) ----
COLS = [
    "notification", "notif_type", "area", "project", "status",
    "defect_class", "disposition", "cause", "material",
    "opened", "closed", "leadtime", "copq",
    "rc1", "rc2", "origin1", "origin2", "title",
]

# raw SAP header -> tidy name
SAP_MAP = {
    "Notification": "notification",
    "Notification Type": "notif_type",
    "Area": "area",
    "Project Text  (Notification)": "project",
    "Notif. Status": "status",
    "Defect Class": "defect_class",
    "Disposition Action TEXT": "disposition",
    "Notification Cause C TEXT": "cause",
    "Material": "material",
    "Notification Date": "opened",
    "Closing date": "closed",
    "Leadtime": "leadtime",
    "CoPQ (NC)": "copq",
    "Notification TEXT": "title",
}

# raw CAPA headers for the four root-cause columns (note leading spaces on some)
CAPA_NC = "NC/SCAR Number"
CAPA_RC1 = "RC Category L1"
CAPA_RC2 = "RC Category L2"
CAPA_OD1 = " (Real) Origin Area L1"
CAPA_OD2 = " (Real) Origin Area L2"


def _norm_nc(v) -> str:
    """NC numbers must match byte-for-byte between SAP and CAPA. Coerce both to a
    trimmed string; drop any trailing .0 pandas adds when it reads them as float."""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _build_capa_lookup(capa: pd.DataFrame) -> dict:
    """NC number -> {rc1, rc2, origin1, origin2}. When an NC has several CAPA
    rows, prefer the one whose RC Category L1 is a real value over N/A/blank."""
    look: dict[str, dict] = {}
    have = {c: (c in capa.columns) for c in
            (CAPA_NC, CAPA_RC1, CAPA_RC2, CAPA_OD1, CAPA_OD2)}
    if not have[CAPA_NC]:
        return look
    for _, r in capa.iterrows():
        nc = _norm_nc(r[CAPA_NC])
        if not nc or nc.lower() in ("nan", "none", ""):
            continue
        rc1 = str(r[CAPA_RC1]).strip() if have[CAPA_RC1] else ""
        rec = {
            "rc1": rc1,
            "rc2": str(r[CAPA_RC2]).strip() if have[CAPA_RC2] else "",
            "origin1": str(r[CAPA_OD1]).strip() if have[CAPA_OD1] else "",
            "origin2": str(r[CAPA_OD2]).strip() if have[CAPA_OD2] else "",
        }
        real = rc1 and rc1 not in ("N/A", "nan", "None", "", "0", "(blank)")
        if nc not in look or real:
            look[nc] = rec
    return look


def load(path: str) -> pd.DataFrame:
    """Load a workbook and return the tidy DataFrame.

    If the workbook has Raw_Data_SAP (+ optionally Raw_Data_Capa), it joins them.
    Otherwise it treats the first sheet as an already-joined MI export.
    """
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names

    if "Raw_Data_SAP" in sheets:
        sap = pd.read_excel(path, sheet_name="Raw_Data_SAP")
        capa = (pd.read_excel(path, sheet_name="Raw_Data_Capa")
                if "Raw_Data_Capa" in sheets else pd.DataFrame())
        return _from_raw(sap, capa)

    # already-joined export: first sheet, tidy or MI-export headers
    df = pd.read_excel(path, sheet_name=0)
    return _from_export(df)


def _from_raw(sap: pd.DataFrame, capa: pd.DataFrame) -> pd.DataFrame:
    look = _build_capa_lookup(capa) if not capa.empty else {}

    out = pd.DataFrame()
    for raw_col, tidy in SAP_MAP.items():
        out[tidy] = sap[raw_col] if raw_col in sap.columns else None

    key = sap["Notification"].map(_norm_nc) if "Notification" in sap.columns else ""
    out["notification"] = key
    out["rc1"] = key.map(lambda n: look.get(n, {}).get("rc1", ""))
    out["rc2"] = key.map(lambda n: look.get(n, {}).get("rc2", ""))
    out["origin1"] = key.map(lambda n: look.get(n, {}).get("origin1", ""))
    out["origin2"] = key.map(lambda n: look.get(n, {}).get("origin2", ""))

    return _finish(out)


def _from_export(df: pd.DataFrame) -> pd.DataFrame:
    # Accept either tidy names or the MI-export headers produced by the join step.
    ren = {
        "Notification": "notification", "Notification Type": "notif_type",
        "Area": "area", "Project Text": "project", "Notif. Status": "status",
        "Defect Class": "defect_class", "Disposition Action TEXT": "disposition",
        "Notification Cause C TEXT": "cause", "Material": "material",
        "Notification Date": "opened", "Closing date": "closed",
        "Leadtime": "leadtime", "CoPQ": "copq",
        "RC Category L1": "rc1", "RC Category L2": "rc2",
        "Origin Area L1": "origin1", "Origin Area L2": "origin2",
        "Notification TEXT": "title",
    }
    df = df.rename(columns={k: v for k, v in ren.items() if k in df.columns})
    for c in COLS:
        if c not in df.columns:
            df[c] = None
    if "notification" in df.columns:
        df["notification"] = df["notification"].map(_norm_nc)
    return _finish(df[COLS])


def _finish(df: pd.DataFrame) -> pd.DataFrame:
    """Common cleaning: types, blanks, area normalisation."""
    df = df.copy()

    # Area: "I:" / "M:" -> keep as "I" / "M" for clean labels
    df["area"] = (df["area"].astype(str).str.strip()
                  .str.rstrip(":").replace({"nan": "", "None": ""}))

    # Status tidy
    df["status"] = df["status"].astype(str).str.strip().replace(
        {"nan": "", "None": ""})

    # CoPQ numeric, absolute (raw values are negative)
    df["copq"] = pd.to_numeric(df["copq"], errors="coerce").fillna(0).abs()

    # leadtime numeric
    df["leadtime"] = pd.to_numeric(df["leadtime"], errors="coerce")

    # dates
    for c in ("opened", "closed"):
        df[c] = pd.to_datetime(df[c], errors="coerce")

    # root-cause blanks -> a single canonical empty so charts group them together
    for c in ("rc1", "rc2", "origin1", "origin2"):
        df[c] = (df[c].astype(str).str.strip()
                 .replace({"nan": "", "None": "", "0": "", "(blank)": ""}))

    # text fields
    for c in ("disposition", "cause", "material", "title", "project",
              "notif_type", "defect_class"):
        df[c] = df[c].astype(str).str.strip().replace({"nan": "", "None": ""})

    return df.reset_index(drop=True)


# ---- small helpers the charts/KPIs use ----

def area_label(a: str) -> str:
    return {"I": "Integration", "M": "Matching"}.get(str(a).strip(), str(a).strip())


def class_label(v) -> str:
    """Defect class 4-5 = Major, 0-3 = Minor, blank = Unclassified. Handles the
    same int/float/text mix as the SAS dashboard."""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if s in ("4", "5"):
        return "Major"
    if s in ("0", "1", "2", "3"):
        return "Minor"
    return "Unclassified"


def rc_or_none(v) -> str:
    s = str(v).strip()
    return s if s and s.lower() not in ("nan", "none", "") else "(not coded)"
