"""
Validation rules for SAS NC data. Applied identically by:
  - the dashboard's upload endpoint (Python, server-side)
  - the SAS export builder (JavaScript, browser-side — rules ported manually)

Every rule returns a list of findings. A finding is:
  (severity, row_index_or_None, nc_id, field, message)

Severities:
  RED    = data is impossible. The load is BLOCKED.
  AMBER  = data is suspicious. The load proceeds with a warning.
  INFO   = noteworthy but harmless.

These rules are the ONLY place validation logic lives in Python. If a check
needs adding, add it here and port it to the builder's JS.
"""
from __future__ import annotations
import re
import pandas as pd

RED = "RED"
AMBER = "AMBER"
INFO = "INFO"

# notification format: SAP = digits, Teamcenter = IR-digits
_SAP_NC = re.compile(r"^\d{5,10}$")
# short form IR-001579 OR full Teamcenter object string IR-001579/A;1-...
_TC_NC = re.compile(r"^IR-\d{4,8}(/.*)?$")
_VALID_STATUS = {"Closed", "Open", "Deleted"}


def _clean_nc(nc) -> str:
    """Extract a clean id from a raw notification value.
    IR-001579/A;1-NC_178...  ->  IR-001579
    221716                   ->  221716 (unchanged)
    """
    nc = str(nc)
    m = re.match(r'(IR-\d{4,8})', nc)
    if m:
        return m.group(1)
    return nc


def validate(df: pd.DataFrame, current_count: int = 0) -> list[tuple]:
    """Run every check on a DataFrame shaped like the ingest input.
    Returns a list of (severity, row, nc, field, message)."""
    findings = []

    if df.empty:
        findings.append((RED, None, None, None, "file contains no rows"))
        return findings

    nc_col = None
    for c in ("Notification", "notification"):
        if c in df.columns:
            nc_col = c
            break
    if nc_col is None:
        findings.append((RED, None, None, None,
                         "no 'Notification' column found"))
        return findings

    ncs_raw = df[nc_col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    ncs = ncs_raw.map(_clean_nc)

    # ---- 1. notification format ----
    for i, nc in enumerate(ncs_raw):
        nc = str(nc)
        if not nc or nc.lower() in ("nan", "none", ""):
            findings.append((RED, i, nc, nc_col, "blank notification number"))
        elif not _SAP_NC.match(nc) and not _TC_NC.match(nc):
            findings.append((RED, i, nc, nc_col,
                             f"'{nc}' is not a valid NC number "
                             f"(expected digits or IR-digits)"))

    # ---- 2. duplicate notification numbers ----
    dups = ncs[ncs.duplicated(keep=False)]
    if len(dups):
        dup_vals = sorted(dups.unique())
        for v in dup_vals[:10]:
            rows = list(dups[dups == v].index)
            findings.append((RED, rows[0], v, nc_col,
                             f"duplicate NC '{v}' on {len(rows)} rows"))

    # ---- 3. status values ----
    # FIX: coerce every cell to str() before calling .lower() / .title().
    # A blank status cell arrives as float NaN even after .astype(str) in
    # some pandas versions, and NaN has no .lower().
    status_col = None
    for c in ("Notif. Status", "status", "Status"):
        if c in df.columns:
            status_col = c
            break
    if status_col:
        for i, (nc, st_raw) in enumerate(zip(ncs, df[status_col])):
            st = str(st_raw).strip() if pd.notna(st_raw) else ""
            if st and st not in _VALID_STATUS and st.lower() not in ("nan", "none", ""):
                # title-case it and check again
                if st.title() not in _VALID_STATUS:
                    findings.append((AMBER, i, nc, status_col,
                                     f"status '{st}' is not Closed/Open/Deleted"))

    # ---- 4. Open with a closing date ----
    close_col = None
    for c in ("Closing date", "closed", "Closed"):
        if c in df.columns:
            close_col = c
            break
    if status_col and close_col:
        for i, (nc, st_raw, cd) in enumerate(zip(
                ncs, df[status_col], df[close_col])):
            st = str(st_raw).strip() if pd.notna(st_raw) else ""
            is_open = st.title() == "Open" or (not st or st.lower() in ("nan", "none"))
            has_close = pd.notna(cd) and str(cd).strip() not in ("", "nan", "None", "NaT")
            if is_open and has_close:
                findings.append((RED, i, nc, status_col,
                                 f"status is Open but has closing date '{cd}'"))

    # ---- 5. closed date before opened date ----
    open_col = None
    for c in ("Notification Date", "opened", "Opened"):
        if c in df.columns:
            open_col = c
            break
    if open_col and close_col:
        opened = pd.to_datetime(df[open_col], errors="coerce")
        closed = pd.to_datetime(df[close_col], errors="coerce")
        for i, (nc, o, c) in enumerate(zip(ncs, opened, closed)):
            if pd.notna(o) and pd.notna(c) and c < o:
                findings.append((AMBER, i, nc, close_col,
                                 f"closed {c.date()} before opened {o.date()}"))

    # ---- 6. sharp row-count drop (partial file loaded as Start over) ----
    if current_count > 0 and len(df) < current_count * 0.5:
        findings.append((AMBER, None, None, None,
                         f"file has {len(df)} rows but the dashboard currently "
                         f"holds {current_count}. If you use 'Start over', "
                         f"{current_count - len(df)} NCs will be deleted."))

    # ---- 7. leadtime vs dates ----
    lt_col = None
    for c in ("Leadtime", "leadtime"):
        if c in df.columns:
            lt_col = c
            break
    if lt_col and open_col and close_col:
        opened = pd.to_datetime(df[open_col], errors="coerce")
        closed = pd.to_datetime(df[close_col], errors="coerce")
        lead = pd.to_numeric(df[lt_col], errors="coerce")
        for i, (nc, o, c, lt) in enumerate(zip(ncs, opened, closed, lead)):
            if pd.notna(o) and pd.notna(c) and pd.notna(lt):
                expected = (c - o).days
                if abs(lt - expected) > 2:
                    findings.append((INFO, i, nc, lt_col,
                                     f"leadtime {int(lt)}d but dates say "
                                     f"{expected}d ({o.date()} to {c.date()})"))

    return findings


def reds(findings: list[tuple]) -> list[tuple]:
    return [f for f in findings if f[0] == RED]


def ambers(findings: list[tuple]) -> list[tuple]:
    return [f for f in findings if f[0] == AMBER]


def summary(findings: list[tuple]) -> str:
    r = len(reds(findings))
    a = len(ambers(findings))
    if r:
        return f"BLOCKED: {r} error(s) must be fixed before loading."
    if a:
        return f"WARNING: {a} issue(s) found. Review them before proceeding."
    return "OK: no issues found."