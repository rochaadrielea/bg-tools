"""
ingest.py
---------
Reads NCR_Cutover_Tracker.xlsx AND NC_s_Overview (SAP export),
cleans, merges (deduplicates overlapping NCs), loads to SQLite.

Two files in data/:
  - NCR_Cutover_Tracker*.xlsx  → active tracker (owner/disposition detail)
                                 .xlsm and .csv are accepted too; a CSV is read
                                 as a plain export of the NC_Tracker_Black_Out
                                 sheet
  - NC_s_Overview*.xlsx        → SAP full history (CoPQ/leadtime/defect)
  - *CAPA*.xls[xm]             → CAPA/RCA tracker (optional)

Merge logic:
  - NCs in both files    → enriched: tracker fields + SAP fields, source='both'
  - Tracker-only rows    → source='tracker'
  - SAP-only rows        → source='sap'
  On an overlapping NC the TRACKER wins on dates and status; SAP only fills a
  field the tracker left blank. The tracker is the live cutover record and the
  SAP export is a dated snapshot, so the live record takes precedence.

CAPA table:
  - Keyed on (nc_id, capa_type). NC/SCAR Number is the foreign key to nc.nc_id.
  - Keeps ALL three types (RCA / CA / PA) — one row per NC per type.
  - Origin Area / RC Category L1+L2 are only populated on RCA rows by design;
    CA and PA rows exist to record that the action exists, not a root cause.
  - Where the same (nc_id, capa_type) appears more than once, the row carrying
    the most content wins (an all-N/A duplicate never beats a filled row).

Launcher class (LLV / MLV / SLV / SAS):
  - Written onto both nc and capa. The CAPA tracker already encodes the class
    in "Affected Project" (LLV_A6, LLV_Vulcan, MLV_Vega, SLV_FlexLine, SAS), so
    that prefix is the source of truth. Vulcan is LLV_Vulcan and therefore sits
    inside LLV. SAS carries no prefix and stands alone as its own class.

Run daily/weekly to refresh. Idempotent (safe to re-run).

Usage:
    python ingest.py
"""

import os
import re
import sqlite3
from pathlib import Path
from datetime import datetime, date

import pandas as pd

# ---- CONFIG ----
DATA_DIR = Path("data")
DB_FILE = "quality.db"
TRACKER_SHEET = "NC_Tracker_Black_Out"
CAPA_GLOB = "*CAPA*.xls*"      # matches .xlsx AND .xlsm (the real file is macro-enabled)
CAPA_SHEET = "Requestor"
# RCA / CA / PA are the internal actions. EXT-8D is the external supplier 8D —
# it IS a corrective-action record, so an NC that only carries an Ext-8D still
# "has a CAPA". Dropping it (as this list used to) made 131 covered NCs read as
# 'CAPA open' on the burnout. It is kept as its own type so the RCA-only views
# (pizza, root-cause) are unaffected — they filter capa_type='RCA' explicitly.
CAPA_TYPES = ["RCA", "CA", "PA", "EXT-8D"]

# EZ1 / EZYone go-live. Ambiguous tracker dates (both parts <=12) on EZ1/Blackout
# rows that would land BEFORE this are read day-first instead of month-first, so
# a 3-August NC written "03/08/2026" stops appearing in March. If the real
# go-live is a different month, change this one line and re-run.
EZ1_GO_LIVE = date(2026, 5, 1)

# Column maps live at module level so validate.py can check the real mapping
# rather than a hand-kept copy of it. Two lists that can disagree WILL disagree:
# an earlier version of validate.py had its own list, and warned that 'NC Type',
# 'Failure' and 'Material' were unrecognised columns while load_tracker was
# reading all three quite happily.
TRACKER_COL_MAP = {
    "System": "system",
    "ID-Blackout": "nc_id",            # was "Title"
    "Title": "nc_id",                  # fallback for old snapshots
    "NC Type": "nc_type",
    "TC ID": "tc_id",
    "Migrated to EZ1": "migrated_to_ez1",
    "Project": "project",              # was "Project + Flight Unit"
    "Project + Flight Unit": "project",# fallback
    "Flight Unit": "flight_unit",      # new dedicated column
    "Detection": "detection_area",
    "Title and Problem Description": "description",  # was "Description"
    "Description": "description",      # fallback
    "Failure": "failure",
    "Material": "material",
    "Batch": "batch",
    "Issue Owner (QA/PA)": "owner",
    "Issue Owner": "owner",            # fallback
    "Created On": "created_on",
    "NRB disposition": "nrb_disposition",
    "Disposition Implemented Date": "disposition_date",
    "Classification": "classification",
    "PSP ref.": "psp_ref",
    "NC WBS (EzyOne)": "nc_wbs",
    "Status": "status",                # was "Notific. Status"
    "Notific. Status": "status",       # fallback
    "Closure date": "closure_date",
    "Supplier name": "supplier_name",
}

# Columns whose absence actually breaks something downstream. The rest of
# TRACKER_COL_MAP is either optional or has a fallback spelling.
TRACKER_REQUIRED_COLS = [
    "System", "ID-Blackout", "Issue Owner (QA/PA)", "Created On", "Status",
]


# ---- FILE FINDERS ----
def find_file(pattern):
    """Find the newest file matching pattern in data/.
    Skips Excel lock/temp files (~$...) which appear while a workbook is open."""
    DATA_DIR.mkdir(exist_ok=True)
    candidates = [p for p in DATA_DIR.glob(pattern) if not p.name.startswith("~$")]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


# ---- CLEANERS ----
def clean_status(val):
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    return str(val).strip().upper()


def clean_classification(val):
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    v = str(val).strip()
    v = re.sub(r"\s*-\s*", " ", v)
    v = re.sub(r"\s+", " ", v)
    v = v.title()
    # 'Minor Level 1' / 'Minor Level 2' are severities of Minor, not separate
    # classes. Collapsing them here keeps the Major vs Minor KPI to two bars.
    if v.startswith("Minor"):
        return "Minor"
    if v.startswith("Major"):
        return "Major"
    return v


def clean_owner(val):
    """Normalise 'Issue Owner (QA/PA)' to one canonical spelling per person.

    The tracker has no naming standard: the same person is entered as a bare
    first name on some rows and a full name on others, which splits their NC
    count across two owner values and undercounts everyone on the dashboard
    (verified against NCR_Cutover_Tracker_14_07_2026: Vitaly 34 + Vitaly Meshin
    11, Rikard 26 + Rikard Bjon 9, Tiziano 33 + Tiziano Casarico 3, and so on).

    Rule (Adriele, 04.08.2026): whenever only a first name is recorded, complete
    it with the surname. Most come from a bare first name matching exactly one
    full name elsewhere in the file. 'Güven' and 'Nikolaos' appear nowhere in
    full inside the tracker — their surnames come from the Quality team
    distribution list (Güven Bozcali, Nikolaos Madenas).

    Deliberately NOT mapped, and left exactly as typed:
      - 'E.Martins/ R. Bjon', 'Vitaly /Salvatore', 'Vitaly/ Domingos'
                             — two people on one NC; the data model has one
                               owner per NC, so these stay as their own value
                               rather than silently crediting one person
    """
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    v = str(val).strip()
    aliases = {
        # bare first name -> the single full name it maps to
        "Vitaly": "Vitaly Meshin",
        "Rikard": "Rikard Bjon",
        "Tiziano": "Tiziano Casarico",
        "Noel": "Noel Orwa",
        "Joaquin": "Joaquin Vera Rubio",
        "Domingos": "Domingos Moreira",
        "Salvatore": "Salvatore Varchetta",
        "Tiziano ": "Tiziano Casarico",
        # surnames taken from the Quality team distribution list, not the tracker
        "Güven": "Güven Bozcali",
        "Guven": "Güven Bozcali",
        "Nikolaos": "Nikolaos Madenas",
        # long form -> short canonical form
        "Domingos Manuel Ferreira Moreira": "Domingos Moreira",
        "Noel Kenoreny Orwa": "Noel Orwa",
        "S. Scampini": "Simone Scampini",
        "Nikolaos Madenas ": "Nikolaos Madenas",
    }
    return aliases.get(v, v)


def clean_project(val):
    """Normalize raw project text to the canonical validation set.

    Canonical: Ariane 5, Ariane 6, Vega, Vega-C, Atlas, Vulcan, Relativity,
    MHI_H3, Flexline, SAS, SLS. Anything that doesn't map cleanly -> None
    (shown as '(no project)').

    Order matters: specific before generic. 'VEGA-C' has to be tested before
    'VEGA', and 'ARIANE 6' before 'ARIANE', or the generic rule swallows both.

    An earlier version of this function had no Atlas rule at all, so every
    Atlas NC fell through to None: 1,035 NCs sat in '(no project)' and filled
    a third of the CAPA icicle with grey. It also collapsed Ariane 5 and
    Ariane 6 into a single 'Ariane' bucket of 1,620 NCs, and folded Vega-C
    into Vega.
    """
    if pd.isna(val):
        return None
    v = str(val).strip()
    if not v or v == '\xa0':
        return None

    up = v.upper()

    # Explicit non-projects / placeholders -> no project
    _blanks = ("BLANK", "(BLANK)", "NOT ASSIGNED", "TBD", "??", "?", "ALL",
               "RETURN TO FLY", "N/A", "-")
    if up in _blanks:
        return None
    if up.startswith("ADMIN") or up.startswith("PU-L") or up.startswith("PU-S") \
       or up.startswith("TESTING"):
        return None

    # Canonical keyword mapping (order matters: check specific before generic)
    # Flexline covers the small-launch customers: SpaceOne, Gilmour, Miura.
    if "FLEXLINE" in up or "SPACE ONE" in up or "SPACEONE" in up \
       or "GILMOUR" in up or "MIURA" in up:
        return "Flexline"
    if up.startswith("MHI") or up == "H3" or "H3" in re.split(r"[ ,/()-]+", up):
        return "MHI_H3"
    if "RELATIVITY" in up or up.startswith("RS "):
        return "Relativity"
    # VCN is the SAP short code for Vulcan; ACES and Kuiper are Vulcan shipsets.
    if "VULCAN" in up or up.startswith("VCN") or "VCN " in up \
       or "KUIPER" in up or "ACES" in up:
        return "Vulcan"
    if "SAS" in up or "KDS" in up:
        return "SAS"
    if "VEGA-C" in up or "VEGA C" in up:
        return "Vega-C"
    if "VEGA" in up or "VERTA" in up:
        return "Vega"
    if "ARIANE 6" in up or up.startswith("A6") or "ARIANE6" in up:
        return "Ariane 6"
    if "ARIANE 5" in up or up.startswith("A5") or "ARIANE5" in up:
        return "Ariane 5"
    if "ARIANE" in up:
        return "Ariane 6"
    # ISA and LPM are Atlas payload adapters; ATL is the SAP short code.
    if "ATLAS" in up or up.startswith("ATL ") or "ATL " in up \
       or "ISA" in up or "LPM" in up:
        return "Atlas"
    if "SLS" in up:
        return "SLS"

    # Unrecognized -> no project (kept out of the clean charts)
    return None


# ---- LAUNCHER CLASS ----
# Four classes, taken from the prefix the CAPA tracker already writes into
# "Affected Project": LLV_A6, MLV_Vega, SLV_FlexLine, SAS. Vulcan is written
# LLV_Vulcan, so Vulcan belongs inside LLV rather than standing on its own.
# SAS carries no prefix and is its own class.
LAUNCHER_CLASS = {
    "Ariane 5": "LLV", "Ariane 6": "LLV", "Atlas": "LLV", "Vulcan": "LLV",
    "Relativity": "LLV", "MHI_H3": "LLV", "SLS": "LLV",
    "Vega": "MLV", "Vega-C": "MLV",
    "Flexline": "SLV",
    "SAS": "SAS",
}


def launcher_class(project):
    """LLV / MLV / SLV / SAS from a canonical project name."""
    if pd.isna(project) or not project:
        return "(no class)"
    return LAUNCHER_CLASS.get(str(project).strip(), "(no class)")


def capa_launcher_class(affected_project):
    """LLV / MLV / SLV / SAS read straight off the CAPA 'Affected Project' prefix.

    Values with no prefix (All Projects, VCN Dome, Lager, ISA, 360_LPM, and the
    handful of part names people type in) fall back to the project keyword rules.
    """
    if pd.isna(affected_project):
        return "(no class)"
    up = str(affected_project).strip().upper()
    m = re.match(r"(LLV|MLV|SLV)[_ ]", up)
    if m:
        return m.group(1)
    if up == "SAS" or up.endswith("_SAS") or up.startswith("SAS"):
        return "SAS"
    return launcher_class(clean_project(affected_project))


def clean_text(val):
    if pd.isna(val):
        return None
    v = str(val).strip()
    if v in ('', '\xa0', '-'):
        return None
    return v


# ---- Detection-area consolidation (data treatment) --------------------------
# The same detection area is recorded two different ways: the old SAP free-text
# ("Incoming inspection") and the new NC-tracker letter-codes ("II: Incoming
# Inspection"). Left as-is, one real area shows up as two separate bars in
# "Open NCs by Area". This map collapses the pairs that are UNAMBIGUOUSLY the
# same area onto one canonical name. Matching is case-insensitive; anything not
# listed here is kept exactly as recorded (so we never merge things that only
# look similar — e.g. 'M: Machining' vs 'Manufacturing' are left separate).
DETECTION_AREA_CANON = {
    "ii: incoming inspection": "Incoming inspection",
    "incoming inspection":     "Incoming inspection",
    "cc: customer complaint":  "Customer complaint",
    "customer complaint":      "Customer complaint",
    "s: supplier":             "Supplier",
    "supplier":                "Supplier",
    "b: bonding":              "Bonding",
    "bonding":                 "Bonding",
    "t: testing":              "Test",
    "test":                    "Test",
    "p: post delivery":        "Post delivery",
    "post delivery":           "Post delivery",
}


def canon_detection_area(val):
    """Collapse detection-area synonyms to one canonical name (see map above)."""
    v = clean_text(val)
    if v is None:
        return None
    return DETECTION_AREA_CANON.get(v.strip().lower(), v)


def parse_date(val, dayfirst=True):
    """Parse a date cell to ISO. `dayfirst` must match the source's convention.

    The NC tracker writes US format (M/D/Y): values like 6/23/2026 and 3/25/2026
    have a day above 12, which only parses as month-first. Read day-first, an
    ambiguous 06/12/2026 became 6 December instead of 12 June — that pushed 29
    NCs into the future with a negative age, and emptied 'New since start' and
    'Avg new / week' because every post-go-live NC fell outside the To date.
    So the tracker loader passes dayfirst=False.
    """
    if pd.isna(val) or not str(val).strip() or str(val).strip() == '\xa0':
        return None
    try:
        return pd.to_datetime(val, dayfirst=dayfirst).date().isoformat()
    except Exception:
        return None


def parse_date_us(val):
    """Month-first date, for the NC tracker."""
    return parse_date(val, dayfirst=False)


def parse_ncr_date(val, not_before=None, not_after=None):
    """Parse a tracker date that may be M/D/Y OR D/M/Y — the source mixes both.

    The tracker is typed by several people: most rows are US month-first
    (6/23/2026), but some are European day-first (13/07/2026, 31/07/2026), and a
    few are genuinely ambiguous (03/08/2026 — 8 March or 3 August?).

    Rule:
      * a token > 12 fixes the format unambiguously;
      * an ambiguous date defaults to month-first (the file's dominant style),
        BUT flips to day-first if month-first would land before `not_before`
        (the EZ1 go-live), OR after `not_after` (a creation date cannot be in
        the future). That turns a July/August EZ1 written "03/08/2026" back into
        3 August instead of the phantom 8 March.

    NOTE: only the TEXT (CSV) path reaches this branch. In an .xlsx these cells
    arrive as real datetimes and return early below — Excel already picked an
    orientation and ingest must not second-guess it.

    Deterministic and explainable — it never relies on pandas guessing.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val.date().isoformat()
    s = str(val).strip()
    if not s or s == '\xa0':
        return None
    p = s.replace(".", "/").replace("-", "/").split("/")
    if len(p) == 3 and p[0].isdigit() and p[1].isdigit() and p[2].isdigit():
        a, b, y = int(p[0]), int(p[1]), int(p[2])
        if y < 100:
            y += 2000

        def mk(mo, da):
            try:
                return date(y, mo, da)
            except ValueError:
                return None

        if a > 12 and b <= 12:
            d = mk(b, a)                       # unambiguous day-first
        elif b > 12 and a <= 12:
            d = mk(a, b)                       # unambiguous month-first
        elif a <= 12 and b <= 12:             # ambiguous
            dm, dd = mk(a, b), mk(b, a)        # month-first, day-first
            # The tracker is European DD/MM/YYYY (e.g. '06.08.2026' = 6 August),
            # so an ambiguous date is read DAY-FIRST. Month-first mis-read
            # '06.08.2026' as 8 June and dropped a closure into the wrong month.
            if not_after and dd and dd > not_after and dm and dm <= not_after:
                d = dm      # day-first would be in the future -> use month-first
            elif not_before and dd and dd < not_before and dm and dm >= not_before:
                d = dm      # day-first before go-live -> use month-first
            else:
                d = dd or dm
        else:
            d = None
        if d:
            return d.isoformat()
    # Fallback for ISO or anything unusual — try month-first, then day-first.
    try:
        return pd.to_datetime(s, dayfirst=False).date().isoformat()
    except Exception:
        try:
            return pd.to_datetime(s, dayfirst=True).date().isoformat()
        except Exception:
            return None


def clean_level(val):
    """Clean an Origin/RC level value. '0', 'N/A', blank -> None (not recorded)."""
    if pd.isna(val):
        return None
    v = str(val).strip()
    if not v or v == '\xa0':
        return None
    if v.upper() in ("N/A", "NA", "N.A.", "-", "#N/A", "NONE"):
        return None
    if v == "0":          # '0' is the placeholder people type for "nothing"
        return None
    return v


def clean_capa_type(val):
    """Normalise the CAPA Type column to RCA / CA / PA. Anything else -> None."""
    if pd.isna(val):
        return None
    v = str(val).strip().upper()
    if v in ("", "NAN", "N/A", "NA", "NONE", "-"):
        return None
    return v if v in CAPA_TYPES else None


def clean_capa_status(val):
    """Normalise the Capa Board 'Status' to Closed / Overdue / Open.

    This is the CAPA tracker's OWN status (from the Capa Board tab) and is the
    source of truth for 'CAPA open / overdue' — NOT a rule invented downstream.
    Blank, '#VALUE!' (a broken sheet formula) and '0' -> None (unknown)."""
    if pd.isna(val):
        return None
    v = str(val).strip()
    if not v or v in ("\xa0", "#VALUE!", "#N/A", "0"):
        return None
    vl = v.lower()
    if vl.startswith("close"):
        return "Closed"
    if vl.startswith("overdue"):
        return "Overdue"
    if vl.startswith("open"):
        return "Open"
    return v.title()


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1: Load the CAPA / RCA tracker
# ══════════════════════════════════════════════════════════════════════════════
def load_capa():
    """Load the CAPA / RCA tracker. OPTIONAL — returns None if the file isn't in data/.

    Keyed on (nc_id, capa_type): every NC can carry an RCA row, a CA row and a
    PA row. All three are kept — the presence of a CA/PA row is itself the fact
    being reported, even though the root-cause columns on those rows are N/A by
    design.

    Header row is auto-detected so a banner row above it won't break the load.
    """
    path = find_file(CAPA_GLOB)
    if path is None:
        print(f"  No {CAPA_GLOB} in {DATA_DIR}/ — skipping CAPA (charts will hide).")
        return None
    print(f"  Reading {path.name}...")

    try:
        xl = pd.ExcelFile(path)
        # Prefer the 'Capa Board' tab: it is a superset of 'Requestor' and, unlike
        # it, carries the tracker's own Status + Due Date — the source of truth for
        # CAPA open/overdue (Adriele). Trailing space in the real tab name is real.
        _prefer = ["Capa Board ", "Capa Board", CAPA_SHEET]
        sheet = next((s for s in _prefer if s in xl.sheet_names), xl.sheet_names[0])
        if sheet not in ("Capa Board ", "Capa Board"):
            print(f"  ! 'Capa Board' tab not found, using '{sheet}'. Tabs: {xl.sheet_names}")

        # Auto-detect the header row: the row containing 'NC/SCAR'
        probe = pd.read_excel(path, sheet_name=sheet, header=None, nrows=12, dtype=str)
        header_row = 0
        for i in range(len(probe)):
            row_txt = " ".join(str(x) for x in probe.iloc[i].tolist())
            if "NC/SCAR" in row_txt:
                header_row = i
                break
        df = pd.read_excel(path, sheet_name=sheet, header=header_row, dtype=str)
    except Exception as e:
        print(f"  ! Could not read CAPA file: {e} — skipping.")
        return None

    df.columns = [str(c).strip() for c in df.columns]

    def pick(*names):
        for n in names:
            if n in df.columns:
                return n
        return None

    c_nc = pick("NC/SCAR Number", "NC/SCAR", "NC Number")
    c_type = pick("CAPA Type")
    if c_nc is None:
        print(f"  ! No 'NC/SCAR Number' column found. Columns: {list(df.columns)[:12]} — skipping.")
        return None
    if c_type is None:
        print(f"  ! No 'CAPA Type' column found. Columns: {list(df.columns)[:12]} — skipping.")
        return None

    def _clean_id(x):
        # Match the SAP loader's nc_id format exactly, or the join silently fails.
        # '0' is the placeholder people type for "nothing" (same as clean_level):
        # 78 board lines carry nc_id='0'. Without this guard they all collapse
        # into one phantom NC in the CAPA tab. Treat '0'/'' as no id.
        if pd.isna(x):
            return None
        if isinstance(x, (int, float)):
            try:
                r = str(int(x))
            except Exception:
                r = clean_text(x)
        else:
            s = clean_text(x)
            if s is None:
                return None
            try:                      # '213952.0' -> '213952'
                r = str(int(float(s)))
            except Exception:
                r = s
        return None if r in ("0", "", None) else r

    out = pd.DataFrame()
    out["nc_id"] = df[c_nc].apply(_clean_id)
    out["capa_type"] = df[c_type].apply(clean_capa_type)

    # Root-cause levels — populated on RCA rows only, by design.
    _level_cols = []
    for tgt, src in [
        ("origin_area_l1", pick("(Real) Origin Area L1", "Origin Area L1")),
        ("origin_area_l2", pick("(Real) Origin Area L2", "Origin Area L2")),
        ("rc_category_l1", pick("RC Category L1")),
        ("rc_category_l2", pick("RC Category L2")),
    ]:
        out[tgt] = df[src].apply(clean_level) if src else None
        _level_cols.append(tgt)

    # Context columns. 'responsible' differs per capa_type on the same NC —
    # the person who owns the analysis is not always the one who owns the action.
    _ctx_cols = []
    for tgt, src in [
        ("capa_project", pick("Affected Project")),
        ("requestor", pick("Requestor")),
        ("responsible", pick("Responsible")),
        ("capa_created_on", pick("Creation date requestor")),
        ("problem_description", pick("Problem Description")),
    ]:
        out[tgt] = df[src].apply(clean_text) if src else None
        _ctx_cols.append(tgt)

    # Launcher class straight off the 'Affected Project' prefix.
    _c_proj = pick("Affected Project")
    out["launcher_class"] = (df[_c_proj].apply(capa_launcher_class)
                             if _c_proj else "(no class)")

    # CAPA STATUS + dates from the Capa Board — the tracker's own truth for
    # open/overdue (Adriele: this tab is the source). If the file only has the
    # 'Requestor' sheet these columns won't exist and stay None.
    _c_status = pick("Status")
    out["capa_status"] = df[_c_status].apply(clean_capa_status) if _c_status else None
    for _tgt, _src in [("due_date", pick("Due Date")),
                       ("open_date", pick("Open Date")),
                       ("close_date", pick("Close Date"))]:
        out[_tgt] = df[_src].apply(parse_ncr_date) if _src else None

    n_raw = len(out)
    out = out[out["nc_id"].notna()].copy()
    n_no_id = n_raw - len(out)

    n_before_type = len(out)
    out = out[out["capa_type"].notna()].copy()
    n_bad_type = n_before_type - len(out)

    # NO dedup on (nc_id, capa_type). An NC legitimately has several actions of
    # the same type (the board numbers them 'ID Number 1, 2, 3…'), so every valid
    # CAPA line is kept — the CAPA count/status must match the Capa Board line for
    # line (Adriele: "take the data only from the Capa Board"). The RCA donuts in
    # capa_view.py de-duplicate per NC on their own so they are not double-counted.
    out = out.reset_index(drop=True)
    n_dupes = 0

    # ---- Reporting ----
    _by_type = out["capa_type"].value_counts().to_dict()
    n_ncs = out["nc_id"].nunique()
    n_o1 = out["origin_area_l1"].notna().sum()
    n_r1 = out["rc_category_l1"].notna().sum()
    n_rca = _by_type.get("RCA", 0)

    print(f"  {len(out)} CAPA rows across {n_ncs} NCs "
          f"(RCA: {_by_type.get('RCA', 0)} | CA: {_by_type.get('CA', 0)} | "
          f"PA: {_by_type.get('PA', 0)} | Ext-8D: {_by_type.get('EXT-8D', 0)})")
    if n_no_id or n_bad_type or n_dupes:
        print(f"    dropped: {n_no_id} no NC number | {n_bad_type} unrecognised CAPA type "
              f"| {n_dupes} duplicate (nc_id, capa_type)")
    print(f"    Origin L1 filled: {n_o1}/{n_rca} RCA rows | "
          f"RC Category L1 filled: {n_r1}/{n_rca} RCA rows")
    _by_class = out["launcher_class"].value_counts().to_dict()
    print(f"    Launcher class: " + " | ".join(f"{k}: {v}" for k, v in _by_class.items()))
    if "capa_status" in out.columns and out["capa_status"].notna().any():
        _by_status = out["capa_status"].value_counts(dropna=False).to_dict()
        print(f"    CAPA status (from Capa Board): "
              + " | ".join(f"{k}: {v}" for k, v in _by_status.items()))
    return out


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2: Load the NCR Cutover Tracker
# ══════════════════════════════════════════════════════════════════════════════
def read_tracker_file(f):
    """Read the tracker from .xlsx/.xlsm (named sheet) or from a .csv export.

    The workbook is the intended source, but the sheet is often exported to CSV
    from SharePoint and dropped into data/ instead. That used to be invisible:
    find_file matched *.xlsx only, load_tracker warned and returned empty, and
    the ingest carried on with SAP rows alone — the dashboard then showed 147
    open NCs (SAP only) instead of 221, with EZYone and Blackout flat at zero
    and no error anywhere. Accepting CSV removes that silent failure.

    CSV specifics: SharePoint writes UTF-8 with a BOM, so the first header comes
    back as '﻿System' unless the encoding says utf-8-sig. Every column is
    read as text and the cleaners do the typing, which matches how the Excel
    path already behaves.
    """
    if f.suffix.lower() == ".csv":
        print(f"  Reading {f.name} (CSV export of '{TRACKER_SHEET}')...")
        # Encoding varies by who exported it: SharePoint writes UTF-8 (BOM),
        # but an Excel "Save as CSV" on a German Windows writes cp1252/latin-1
        # (a stray non-breaking space is byte 0xA0, which is NOT valid UTF-8).
        # Try the likely encodings in order rather than crash on the wrong one.
        df = None
        for _enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                df = pd.read_csv(f, encoding=_enc, dtype=str)
                break
            except UnicodeDecodeError:
                continue
        if df is None:
            df = pd.read_csv(f, encoding="latin-1", dtype=str, engine="python")
        df.columns = [str(c).strip().lstrip("﻿") for c in df.columns]
        return df

    xl = pd.ExcelFile(f)
    sheet = TRACKER_SHEET if TRACKER_SHEET in xl.sheet_names else xl.sheet_names[0]
    if sheet != TRACKER_SHEET:
        print(f"  ! Sheet '{TRACKER_SHEET}' not found, using '{sheet}'. Tabs: {xl.sheet_names}")
    print(f"  Reading {f.name} sheet '{sheet}'...")
    df = pd.read_excel(f, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_tracker():
    """Load and clean the NCR Cutover Tracker (.xlsx, .xlsm or .csv)."""
    # Workbook first: if both are present the workbook is the live file and the
    # CSV is usually an older export sitting next to it.
    # Match the tracker by several real-world names people have saved it under:
    # NCR_Cutover_Tracker…, NCtracker…, …NC_Tracker…  (xls/xlsm/csv). The 'CAPA'
    # workbook also contains the word 'Tracker', so we never match a bare
    # '*Tracker*' — only names that carry NC / NCR, which the CAPA file does not.
    f = (find_file("NCR_Cutover_Tracker*.xlsx")
         or find_file("NCR_Cutover_Tracker*.xlsm")
         or find_file("NCtracker*.xlsx")
         or find_file("NCtracker*.xlsm")
         or find_file("*NC_Tracker*.xlsx")
         or find_file("*NC_Tracker*.xlsm")
         or find_file("NCR_Cutover_Tracker*.csv")
         or find_file("NCtracker*.csv")
         or find_file("*NC_Tracker*.csv"))
    if not f:
        print("  ⚠ No tracker file (NCR_Cutover_Tracker / NCtracker / NC_Tracker; "
              ".xlsx / .xlsm / .csv) found in data/ — skipping tracker. "
              "The dashboard will show SAP rows only.")
        return pd.DataFrame()

    df = read_tracker_file(f)

    if "System" not in df.columns:
        print(f"  ! No 'System' column in {f.name}. Columns: {list(df.columns)[:12]} "
              f"— skipping tracker.")
        return pd.DataFrame()

    # Drop blank/example rows
    df = df[df["System"].notna()].copy()
    df = df[df["System"].astype(str).str.strip() != '\xa0'].copy()
    # NC_EXAMPLE marker may live in the description column of the new schema
    _title_col = "Title and Problem Description" if "Title and Problem Description" in df.columns else (
        "Title" if "Title" in df.columns else None)
    if _title_col:
        df = df[df[_title_col].astype(str) != "NC_EXAMPLE"].copy()

    # Column mapping — see TRACKER_COL_MAP at module level. Kept there so
    # validate.py checks the same map the ingest actually uses.
    col_map = TRACKER_COL_MAP
    df = df.rename(columns=col_map)
    # Rename can create duplicate target names (old+new schema both mapped) —
    # keep the first non-empty occurrence of each.
    df = df.loc[:, ~df.columns.duplicated()].copy()
    # Keep only mapped columns (ignore extras), unique + order-preserving
    seen = set()
    keep = [c for c in col_map.values() if c in df.columns and not (c in seen or seen.add(c))]
    df = df[keep].copy()

    # Clean
    df["status"] = df["status"].apply(clean_status)
    df["classification"] = df["classification"].apply(clean_classification)
    df["owner"] = df["owner"].apply(clean_owner)
    df["project"] = df["project"].apply(clean_project)
    df["launcher_class"] = df["project"].apply(launcher_class)

    # Dates — the tracker MIXES M/D and D/M (typed by several people). Use the
    # per-token/plausibility parser: on EZ1/Blackout rows an ambiguous date that
    # would fall before EZ1 go-live is read day-first (fixes "EZ1 in March").
    # 'system' is still the raw string at this point, which is what we want.
    _nb = df["system"].apply(
        lambda s: EZ1_GO_LIVE if str(s).strip() in ("EZ1", "Blackout") else None)
    _today = date.today()   # a creation/closure date cannot be in the future
    df["created_on"] = [parse_ncr_date(v, nb, _today) for v, nb in zip(df["created_on"], _nb)]
    df["closure_date"] = [parse_ncr_date(v, nb, _today) for v, nb in zip(df["closure_date"], _nb)]
    if "disposition_date" in df.columns:
        df["disposition_date"] = [parse_ncr_date(v, nb, _today)
                                  for v, nb in zip(df["disposition_date"], _nb)]

    for col in df.columns:
        if col not in ("created_on", "closure_date", "disposition_date", "launcher_class"):
            df[col] = df[col].apply(clean_text)

    # An EZ1 row often has a blank ID-Blackout but carries a TC ID (NC_… / IR-…).
    # 34 rows in the 11.08 tracker are exactly this. Fall back to tc_id so those
    # NCs are not left without any identifier. (TRACKER_REQUIRED_COLS only checks
    # the column exists, not that every row has a value.)
    if "tc_id" in df.columns:
        df["nc_id"] = df["nc_id"].where(df["nc_id"].notna(), df["tc_id"])

    # Derived
    # Three states, not two. A blank Status is NOT 'closed' — treating it as
    # closed (which `is_open = status == 'OPEN'` silently did) removed those NCs
    # from every open-NC view, and is how owners' counts came out short. There
    # are 19 such rows in the 14/07 tracker across 8 owners. They are now
    # is_open = NULL: neither open nor closed, and visible as '(no status)'.
    def _state(s):
        # clean_status returns None, but pandas stores that as NaN (a float),
        # so test with isna() rather than `is None`.
        if pd.isna(s):
            return "(no status)"
        s = str(s)
        return "Open" if s == "OPEN" else ("Closed" if s == "CLOSED" else s.title())

    df["is_open"] = df["status"].map({"OPEN": 1, "CLOSED": 0})
    df["status_state"] = df["status"].apply(_state)
    # Supplier vs internal comes from the tracker's own 'NC Type' column, which
    # says Internal or Supplier outright. It used to be derived from whether
    # 'Supplier name' was filled — only 19 rows carry a supplier name while 60
    # are typed Supplier, so two thirds of the supplier NCs were being counted
    # as internal. This feeds Susanna's Supplier vs Internal KPI directly.
    if "nc_type" in df.columns:
        df["is_supplier_nc"] = (df["nc_type"].astype(str).str.strip().str.lower()
                                .str.startswith("supplier").astype(int))
    else:
        df["is_supplier_nc"] = df["supplier_name"].notna().astype(int)
    df["days_open"] = (
        pd.to_datetime("today").normalize() - pd.to_datetime(df["created_on"], errors="coerce")
    ).dt.days

    # For SAP-system rows, nc_id is the SAP notification number (for matching)
    df["sap_notif"] = None
    sap_mask = df["system"] == "SAP"
    df.loc[sap_mask, "sap_notif"] = df.loc[sap_mask, "nc_id"].apply(
        lambda x: str(int(float(x))) if x and str(x).replace('.', '').isdigit() else None
    )

    df["source"] = "tracker"
    print(f"  {len(df)} tracker rows loaded.")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3: Load the SAP NC Overview
# ══════════════════════════════════════════════════════════════════════════════
def load_sap_overview():
    """Load and clean the SAP NC Overview export."""
    f = find_file("NC_s_Overview*.xlsx")
    if not f:
        print("  ⚠ No NC_s_Overview*.xlsx found in data/ — skipping SAP overview.")
        return pd.DataFrame()

    print(f"  Reading {f.name}...")
    df = pd.read_excel(f)

    # Skip the totals row (row index 0 after header, where Plant='Totals')
    df = df[df.iloc[:, 0] != 'Totals'].copy()

    col_map = {
        "Plant": "plant_id",
        "Profit Center (WBS Element)": "profit_center",
        "Business Unit": "business_unit",
        "Project Text  (Notification)": "project",
        "Notification Type": "notification_type",
        "Notif. Year": "notif_year",
        "Notification": "nc_id",
        "Notification Date": "created_on",
        "Closing date": "closure_date",
        "Leadtime": "leadtime",
        "CoPQ (NC)": "copq",
        "Defect Class": "defect_class",
        "Defect Code TEXT": "defect_code_text",
        "Disposition Action TEXT": "nrb_disposition",
        "Notification Cause C TEXT": "detection_area",
        "Notification TEXT": "description",
        "Material Key": "material_key",
        "Material": "material",
        "Batch": "batch",
        "[Vendor NC]": "supplier_name",
        "System Status": "system_status",
    }

    # Only map columns that exist
    actual_map = {k: v for k, v in col_map.items() if k in df.columns}
    df = df.rename(columns=actual_map)
    keep = [c for c in actual_map.values() if c in df.columns]
    df = df[keep].copy()

    # Also check the second "Notif. Status" column (index 31 in original)
    # The first one is at index 8, the second at index 31
    # We need the one that says Open/Closed
    if "Notif. Status" in df.columns:
        df = df.rename(columns={"Notif. Status": "status"})
    else:
        # Try to get status from original dataframe
        orig = pd.read_excel(f)
        orig = orig[orig.iloc[:, 0] != 'Totals'].copy()
        if orig.shape[1] > 31:
            df["status"] = orig.iloc[:, 31].apply(clean_text)

    # Clean
    df["nc_id"] = df["nc_id"].apply(lambda x: str(int(x)) if isinstance(x, (int, float)) else clean_text(x))
    df["status"] = df["status"].apply(clean_status) if "status" in df.columns else None
    df["project"] = df["project"].apply(clean_project)
    df["launcher_class"] = df["project"].apply(launcher_class)
    df["detection_area"] = df["detection_area"].apply(clean_text)
    df["description"] = df["description"].apply(clean_text)
    df["nrb_disposition"] = df["nrb_disposition"].apply(clean_text)
    df["supplier_name"] = df["supplier_name"].apply(clean_text)
    df["material"] = df["material"].apply(clean_text)

    # Parse dates — SAP dates might be datetime objects or day-counts
    def parse_sap_date(val):
        if pd.isna(val) or val is None:
            return None
        if isinstance(val, datetime):
            return val.date().isoformat()
        if isinstance(val, (int, float)) and val > 30000:
            # Excel serial date
            try:
                return (datetime(1899, 12, 30) + pd.Timedelta(days=int(val))).date().isoformat()
            except Exception:
                return None
        return parse_date(val)

    df["created_on"] = df["created_on"].apply(parse_sap_date)
    df["closure_date"] = df["closure_date"].apply(parse_sap_date)

    # Derived — three states, same as the tracker (see load_tracker).
    if "status" in df.columns:
        def _state(s):
            if pd.isna(s):
                return "(no status)"
            s = str(s)
            return "Open" if s == "OPEN" else ("Closed" if s == "CLOSED" else s.title())
        df["is_open"] = df["status"].map({"OPEN": 1, "CLOSED": 0})
        df["status_state"] = df["status"].apply(_state)
    else:
        df["is_open"] = None
        df["status_state"] = "(no status)"
    df["is_supplier_nc"] = df.get("notification_type", pd.Series()).astype(str).str.contains("Z2", na=False).astype(int)

    # Classification from defect class
    def classify_defect(val):
        """SAP Defect Class -> Major / Minor.

        SAP changed coding scheme in 2023: MA / MI up to 2022, numeric 0-5 from
        2023 onwards. Not one 2026 NC carries MA or MI, so a Major/Minor KPI
        built on the letters alone reports zero Major.

        Rule (Adriele, 04.08.2026): 3, 4 and 5 are Major; 0, 1 and 2 are Minor.
        The raw code stays in `defect_class`, so the mapping can be re-cut later
        without re-reading SAP. The evidence for and against it sits on the Data
        quality sheet of the Major vs Minor export.
        """
        if pd.isna(val) or str(val).strip() in ('', '-', '\xa0'):
            return None
        v = str(val).strip().upper()
        if v == 'MA':
            return 'Major'
        if v == 'MI':
            return 'Minor'
        if v in ('3', '4', '5'):
            return 'Major'
        if v in ('0', '1', '2'):
            return 'Minor'
        return v
    df["classification"] = df.get("defect_class", pd.Series()).apply(classify_defect)

    df["days_open"] = (
        pd.to_datetime("today").normalize() - pd.to_datetime(df["created_on"], errors="coerce")
    ).dt.days

    # Leadtime and CoPQ as numeric
    if "leadtime" in df.columns:
        df["leadtime"] = pd.to_numeric(df["leadtime"], errors="coerce")
    if "copq" in df.columns:
        df["copq"] = pd.to_numeric(df["copq"], errors="coerce")

    df["sap_notif"] = df["nc_id"]
    df["system"] = "SAP"
    df["source"] = "sap"

    # Columns the tracker has but SAP doesn't
    for col in ["tc_id", "migrated_to_ez1", "failure", "psp_ref", "nc_wbs",
                "purchasing_doc", "po_item", "vendor_code", "disposition_date", "owner"]:
        if col not in df.columns:
            df[col] = None

    print(f"  {len(df)} SAP rows loaded.")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4: Merge
# ══════════════════════════════════════════════════════════════════════════════
def merge_data(tracker_df, sap_df):
    """Merge tracker + SAP, deduplicating on SAP notification number."""
    if tracker_df.empty and sap_df.empty:
        raise RuntimeError("No data loaded from either source.")
    if tracker_df.empty:
        return sap_df
    if sap_df.empty:
        return tracker_df

    # Find overlapping notification numbers
    tracker_sap_notifs = set(tracker_df["sap_notif"].dropna())
    sap_notifs = set(sap_df["sap_notif"].dropna())
    overlap = tracker_sap_notifs & sap_notifs

    print(f"  Overlap: {len(overlap)} NCs in both files")
    print(f"  Tracker-only: {len(tracker_df) - len(tracker_df[tracker_df['sap_notif'].isin(overlap)])} rows")
    print(f"  SAP-only: {len(sap_df) - len(sap_df[sap_df['sap_notif'].isin(overlap)])} rows")

    # For overlapping NCs: start with tracker row, enrich with SAP columns
    tracker_overlap = tracker_df[tracker_df["sap_notif"].isin(overlap)].copy()
    sap_overlap = sap_df[sap_df["sap_notif"].isin(overlap)].set_index("sap_notif")

    # Enrich tracker rows with SAP-only fields
    sap_only_cols = ["notification_type", "defect_class", "copq", "leadtime",
                     "notif_year", "profit_center", "business_unit",
                     "material_key", "system_status", "defect_code_text"]
    for col in sap_only_cols:
        if col in sap_overlap.columns:
            tracker_overlap[col] = tracker_overlap["sap_notif"].map(
                sap_overlap[col].to_dict()).astype(object)

    # Fill missing tracker fields from SAP where tracker is blank.
    # The tracker is the live cutover record, so it keeps precedence on dates
    # and status; SAP only fills a hole.
    # launcher_class is deliberately NOT in this list: it is never blank
    # (it falls back to '(no class)'), so it is recomputed from the final
    # project at the end of the merge instead.
    fill_cols = ["detection_area", "project", "classification", "closure_date"]
    for col in fill_cols:
        if col in sap_overlap.columns:
            # Written with .where() rather than .loc[mask, col] = ... on purpose.
            # pandas 3 refuses to assign a missing value into a str-dtype column
            # and raises TypeError; the server venv runs pandas 3.0.5. Casting to
            # object first and replacing the whole column avoids that entirely,
            # and behaves identically on pandas 2.
            _from_sap = tracker_overlap["sap_notif"].map(sap_overlap[col].to_dict())
            _cur = tracker_overlap[col].astype(object)
            tracker_overlap[col] = _cur.where(_cur.notna(), _from_sap)

    # An overlapping NC with no 'NC Type' typed in the tracker falls back to
    # SAP's notification type (Z2 = procurement complaint = supplier).
    if "notification_type" in tracker_overlap.columns and "nc_type" in tracker_overlap.columns:
        _blank_type = tracker_overlap["nc_type"].isna()
        _z2 = (tracker_overlap["notification_type"].astype(str)
               .str.contains("Z2", na=False).astype(int))
        tracker_overlap.loc[_blank_type, "is_supplier_nc"] = _z2[_blank_type]

    tracker_overlap["source"] = "both"

    # Tracker-only rows (EZ1, Blackout, TC, or SAP NCs not in the overview)
    tracker_only = tracker_df[~tracker_df["sap_notif"].isin(overlap) | tracker_df["sap_notif"].isna()].copy()

    # SAP-only rows (not in tracker)
    sap_only = sap_df[~sap_df["sap_notif"].isin(overlap)].copy()

    # Combine all three
    merged = pd.concat([tracker_overlap, tracker_only, sap_only], ignore_index=True)

    # Ensure consistent columns
    for col in ["notification_type", "defect_class", "copq", "leadtime",
                "notif_year", "profit_center", "business_unit",
                "material_key", "system_status", "defect_code_text"]:
        if col not in merged.columns:
            merged[col] = None

    # Recalculate is_open for all rows (in case SAP data filled closure_date).
    # Rows whose status is still blank keep is_open = NULL — they are neither
    # open nor closed, and must not be silently bucketed into either.
    merged.loc[merged["status"] == "CLOSED", "is_open"] = 0
    merged.loc[merged["status"] == "OPEN", "is_open"] = 1
    if "status_state" not in merged.columns:
        merged["status_state"] = merged["status"].apply(
            lambda s: "(no status)" if pd.isna(s) or not str(s).strip()
            else ("Open" if s == "OPEN" else ("Closed" if s == "CLOSED" else str(s).title())))
    else:
        merged["status_state"] = merged["status_state"].fillna("(no status)")

    # Recompute the class from the final project so a row that picked up its
    # project from SAP during the fill above gets the matching class.
    merged["launcher_class"] = merged["project"].apply(launcher_class)

    # Data treatment: collapse detection-area synonyms so one real area is not
    # counted twice under two spellings (SAP wording vs tracker letter-code).
    merged["detection_area"] = merged["detection_area"].apply(canon_detection_area)

    print(f"  Merged total: {len(merged)} rows")
    return merged


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4b: NEW-SYSTEM (Teamcenter / EZ1) NC export
# ══════════════════════════════════════════════════════════════════════════════
# A Teamcenter export of the new-system NCs. Every row is keyed by an IR- number
# in an 'Object' cell like "IR-001660/A;1-NC_1783..."; the NC_ part (when present)
# is the same TC ID the cutover tracker carries, which is how we avoid counting a
# migrated NC twice. The file name is a timestamp, so it is found by CONTENT:
# the only xls* in data/ (that is not the SAP overview or CAPA) whose first sheet
# has an 'Object' column full of IR- numbers.
import re as _re

def _parse_dt(v):
    """Creation Date in this export is a real datetime string; return ISO date."""
    if v is None or str(v).strip() == "" or str(v).lower() == "nan":
        return None
    try:
        return pd.to_datetime(v).date().isoformat()
    except Exception:
        return None


def load_new_system():
    import glob
    cands = [p for p in glob.glob(str(DATA_DIR / "*.xls*"))
             if not os.path.basename(p).startswith("~$")]
    path = None
    for p in sorted(cands, key=lambda p: os.path.getmtime(p), reverse=True):
        base = os.path.basename(p).upper()
        if any(k in base for k in ("NC_S_OVERVIEW", "OVERVIEW", "CAPA", "NCR_CUTOVER")):
            continue
        try:
            head = pd.read_excel(p, sheet_name=0, nrows=6, dtype=str)
        except Exception:
            continue
        cols = [str(c).strip() for c in head.columns]
        if "Object" in cols and head["Object"].astype(str).str.contains("IR-").any():
            path = p
            break
    if not path:
        print("  (no new-system NC export found in data/ — skipping)")
        return pd.DataFrame()

    print(f"  Reading new-system export {os.path.basename(path)} ...")
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    def col(name):
        return df[name] if name in df.columns else pd.Series([None] * len(df))

    out = pd.DataFrame()
    out["nc_id"] = df["Object"].apply(
        lambda o: (_re.search(r"(IR-\d+)", str(o)).group(1)
                   if _re.search(r"(IR-\d+)", str(o)) else None))
    out["tc_id"] = df["Object"].apply(
        lambda o: (_re.search(r"(NC_\d+)", str(o)).group(1)
                   if _re.search(r"(NC_\d+)", str(o)) else None))
    out["owner"] = col("Issue Owner").apply(clean_owner)
    out["created_on"] = col("Creation Date").apply(_parse_dt)
    out["description"] = df["Object"].apply(
        lambda o: _re.sub(r"^IR-\d+/[^-]*-", "", str(o)).strip() or None)

    out = out[out["nc_id"].notna()].copy()
    out = out.drop_duplicates(subset=["nc_id"], keep="first").reset_index(drop=True)

    # Fields the export does not carry. A brand-new NC with no closure recorded
    # is treated as OPEN (status is filled in later once the tracker catches it).
    out["system"] = "EZ1"
    out["sap_notif"] = None
    out["is_open"] = 1
    out["status"] = "OPEN"
    out["status_state"] = "Open"
    out["is_supplier_nc"] = 0            # internal unless the tracker says otherwise
    out["source"] = "ez1_new"
    out["project"] = None
    out["launcher_class"] = out["project"].apply(launcher_class)
    print(f"    {len(out)} new-system NCs "
          f"({out['tc_id'].notna().sum()} carry a TC ID, "
          f"{out['tc_id'].isna().sum()} IR-only)")
    return out


def _norm_nc(v):
    """A Teamcenter TC ID reduced to NC_<digits> (ignore any trailing _xxxx)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    m = _re.search(r"(NC_\d+)", str(v))
    return m.group(1) if m else None


def _norm_ir(v):
    """An IR number reduced to its integer, so IR-1393 == IR-001393 == IR-1393."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    m = _re.search(r"IR[-\s]?0*(\d+)", str(v))
    return f"IR-{int(m.group(1))}" if m else None


def merge_new_system(merged, new_df):
    """Fold the new-system export into the merged nc table WITHOUT double-counting.

    An NC that is in the new system AND in the tracker is ONE physical NC — it
    must appear once (Adriele). The tracker records a migrated NC by its TC ID
    (NC_…) and/or by its IR number (IR-001640), sometimes both in one cell and
    with different zero-padding (IR-1393 vs IR-001393). So overlap is decided on
    BOTH keys, normalised:

        new NC's TC ID  matches a tracker TC ID   -> same NC, or
        new NC's IR no. matches a tracker IR no.  -> same NC

    On a match we keep the detailed tracker row and just attach the IR number
    (column ir_number); we do NOT append a duplicate. Only NCs that match neither
    key are genuinely new and get appended (source='ez1_new').

    The earlier version compared only the RAW TC-ID string, so it missed every
    IR-number overlap and every NC_ with a differing suffix — it added ~42
    tracker NCs a second time and inflated the open count. See reconcile.py.
    """
    if "ir_number" not in merged.columns:
        merged["ir_number"] = None
    if new_df is None or new_df.empty:
        return merged

    # Map every tracker/merged row's normalised NC_ and IR keys -> its index.
    nc_to_idx, ir_to_idx = {}, {}
    if "tc_id" in merged.columns:
        for idx, raw in merged["tc_id"].items():
            nn, ni = _norm_nc(raw), _norm_ir(raw)
            if nn and nn not in nc_to_idx:
                nc_to_idx[nn] = idx
            if ni and ni not in ir_to_idx:
                ir_to_idx[ni] = idx

    add_rows = []
    n_by_nc = n_by_ir = 0
    for _, r in new_df.iterrows():
        nn = _norm_nc(r.get("tc_id"))
        ni = _norm_ir(r.get("nc_id"))          # new_df.nc_id IS the IR number
        idx = None
        if nn is not None and nn in nc_to_idx:
            idx = nc_to_idx[nn]; n_by_nc += 1
        elif ni is not None and ni in ir_to_idx:
            idx = ir_to_idx[ni]; n_by_ir += 1
        if idx is not None:
            cur = merged.at[idx, "ir_number"]
            if pd.isna(cur) or not str(cur).strip():
                merged.at[idx, "ir_number"] = r.get("nc_id")   # attach IR, no dup
        else:
            add_rows.append(r)

    added = pd.DataFrame(add_rows)
    if not added.empty:
        added["ir_number"] = added["nc_id"]        # for new rows, IR is the id
        for c in merged.columns:
            if c not in added.columns:
                added[c] = None
        added = added[merged.columns]
        merged = pd.concat([merged, added], ignore_index=True)

    print(f"  New-system merge: {n_by_nc + n_by_ir} overlapped the tracker "
          f"({n_by_nc} by TC ID, {n_by_ir} by IR number — IR attached, not "
          f"duplicated), {len(add_rows)} added as new NCs")
    return merged


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5: Write to SQLite
# ══════════════════════════════════════════════════════════════════════════════
def write_db(df, capa=None):
    """Write merged DataFrame to SQLite. Optionally write the CAPA table too."""
    conn = sqlite3.connect(DB_FILE)
    df.to_sql("nc", conn, if_exists="replace", index=False)

    cur = conn.cursor()
    # CAPA table is optional — only created when the file is present
    cur.execute("DROP TABLE IF EXISTS capa")
    if capa is not None and not capa.empty:
        capa.to_sql("capa", conn, if_exists="replace", index=False)
        # NON-unique: an NC legitimately has several CAPAs of the same type on
        # the board (numbered 1,2,3…), and we keep every line.
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_key ON capa(nc_id, capa_type)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_nc ON capa(nc_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_type ON capa(capa_type)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_o1 ON capa(origin_area_l1)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_rc1 ON capa(rc_category_l1)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_class ON capa(launcher_class)")
        if "capa_status" in capa.columns:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_capa_status ON capa(capa_status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_project ON nc(project)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_class ON nc(launcher_class)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_status ON nc(status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_owner ON nc(owner)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_detection ON nc(detection_area)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_source ON nc(source)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_created ON nc(created_on)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_nc_sap_notif ON nc(sap_notif)")
    conn.commit()
    return conn


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def copy_warnings_to_db(conn, warn_db="warnings.db"):
    """Mirror the warning log into quality.db as a read-only copy.

    warnings.db stays the master — quality.db is dropped and rebuilt on every
    ingest, so the log cannot live here or it would be wiped each run. This
    copy exists purely so the dashboard can query warnings alongside NCs
    without opening a second database.
    """
    if not Path(warn_db).exists():
        return 0
    try:
        with sqlite3.connect(warn_db) as wc:
            df = pd.read_sql("SELECT * FROM warnings", wc)
    except Exception as e:
        print(f"  ! Could not read {warn_db}: {e}")
        return 0
    if df.empty:
        return 0
    df.to_sql("warnings", conn, if_exists="replace", index=False)
    cur = conn.cursor()
    cur.execute("CREATE INDEX IF NOT EXISTS idx_warn_check ON warnings(check_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_warn_subject ON warnings(subject)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_warn_sev ON warnings(severity)")
    conn.commit()
    return len(df)


def main():
    print("=" * 60)
    print("Quality DB Ingest — Merge Tracker + SAP Overview")
    print("=" * 60)

    print("\n[1/6] Validating source files...")
    try:
        from validate import run_checks
        _t = (find_file("NCR_Cutover_Tracker*.xlsx")
              or find_file("NCR_Cutover_Tracker*.xlsm")
              or find_file("NCtracker*.xlsx")
              or find_file("NCtracker*.xlsm")
              or find_file("*NC_Tracker*.xlsx")
              or find_file("*NC_Tracker*.xlsm")
              or find_file("NCR_Cutover_Tracker*.csv")
              or find_file("NCtracker*.csv")
              or find_file("*NC_Tracker*.csv"))
        _s = find_file("NC_s_Overview*.xlsx")
        if _t:
            run_checks(_t, _s, clean_owner, TRACKER_SHEET)
        else:
            print("  No tracker found — skipping validation.")
    except ImportError:
        print("  validate.py not found — skipping checks.")
    except Exception as e:
        # Validation must never stop the ingest.
        print(f"  ! Validation error (ingest continues): {e}")

    print("\n[2/6] Loading NCR Cutover Tracker...")
    tracker = load_tracker()

    print("\n[3/6] Loading SAP NC Overview...")
    sap = load_sap_overview()

    print("\n[4/6] Merging (deduplicating overlaps)...")
    merged = merge_data(tracker, sap)

    print("\n[4b/6] Loading new-system (Teamcenter/EZ1) NC export (optional)...")
    new_sys = load_new_system()
    merged = merge_new_system(merged, new_sys)

    print("\n[5/6] Loading CAPA / RCA tracker (optional)...")
    capa = load_capa()

    print("\n[6/6] Writing to SQLite...")
    conn = write_db(merged, capa)

    # Mirror the warning log into quality.db so the dashboard can read it.
    # Runs after write_db because that call replaces the whole file.
    # The CAPA orphan check has to run here too: it compares the capa table to
    # the nc table, so it needs the database that was just written — it cannot
    # be answered from the source files alone.
    try:
        from validate import _connect, check_capa_orphans
        _wc = _connect()
        _no = check_capa_orphans(_wc, DB_FILE)
        _wc.commit(); _wc.close()
        if _no:
            print(f"  ! {_no} CAPA NC number(s) match no NC — logged, left unmatched")
    except Exception as e:
        print(f"  ! CAPA orphan check skipped: {e}")

    _nw = copy_warnings_to_db(conn)
    if _nw:
        print(f"  {_nw} warning(s) copied into quality.db (master stays warnings.db)")

    # Summary. is_open is now three-state (1 / 0 / NULL), so the counts must use
    # explicit IS NULL rather than 1 - is_open, which returns NULL for blanks.
    stats = pd.read_sql("""
        SELECT
            COUNT(*) AS total,
            SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS open_count,
            SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) AS closed_count,
            SUM(CASE WHEN is_open IS NULL THEN 1 ELSE 0 END) AS no_status_count,
            SUM(CASE WHEN source = 'tracker' THEN 1 ELSE 0 END) AS tracker_only,
            SUM(CASE WHEN source = 'sap' THEN 1 ELSE 0 END) AS sap_only,
            SUM(CASE WHEN source = 'both' THEN 1 ELSE 0 END) AS in_both,
            SUM(CASE WHEN project IS NULL THEN 1 ELSE 0 END) AS blank_project,
            SUM(CASE WHEN detection_area IS NULL THEN 1 ELSE 0 END) AS blank_detection,
            SUM(CASE WHEN owner IS NULL THEN 1 ELSE 0 END) AS blank_owner
        FROM nc
    """, conn)

    print("\n" + "=" * 60)
    print("Summary:")
    print("=" * 60)
    print(stats.T.to_string(header=False))

    # ---- Date-parsing sanity: EZ1/Blackout by month, so a phantom pre-go-live
    # month is obvious right in the ingest log (this is the check that catches a
    # future "EZ1 in March"-style date mix). ----
    try:
        dchk = pd.read_sql(
            "SELECT substr(created_on,1,7) AS month, system, COUNT(*) AS n "
            "FROM nc WHERE system IN ('EZ1','Blackout') AND created_on IS NOT NULL "
            "GROUP BY month, system ORDER BY month", conn)
        _early = dchk[dchk["month"] < EZ1_GO_LIVE.strftime("%Y-%m")]
        print("\n" + "=" * 60)
        print(f"EZ1 / Blackout by month (go-live anchor {EZ1_GO_LIVE:%Y-%m}):")
        print("=" * 60)
        print(dchk.to_string(index=False))
        if not _early.empty:
            print(f"  ⚠ {int(_early['n'].sum())} EZ1/Blackout NC(s) dated BEFORE go-live "
                  f"({EZ1_GO_LIVE:%Y-%m}) — check these dates in the tracker:")
            print(_early.to_string(index=False))
    except Exception as e:
        print(f"  ! Date sanity check skipped: {e}")

    # ---- Future creation dates: a created_on after today is a data-entry error
    # in the workbook (Excel already resolved the day/month at entry time, so
    # ingest cannot and must not silently "fix" it). Surface the exact NCs so
    # Elisa/Tiziano can correct them in the tracker. ----
    try:
        _fut = pd.read_sql(
            "SELECT nc_id, system, created_on FROM nc "
            "WHERE created_on > ? ORDER BY created_on",
            conn, params=[date.today().isoformat()])
        if not _fut.empty:
            print("\n" + "=" * 60)
            print(f"  ⚠ {len(_fut)} NC(s) with a creation date IN THE FUTURE "
                  f"(fix in the tracker — ingest does not guess):")
            print("=" * 60)
            print(_fut.to_string(index=False))
    except Exception as e:
        print(f"  ! Future-date check skipped: {e}")

    # Launcher class split — the four classes the CAPA tab draws from.
    cls = pd.read_sql("""
        SELECT COALESCE(launcher_class,'(no class)') AS launcher_class,
               COUNT(*) AS ncs,
               SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS open
        FROM nc GROUP BY launcher_class ORDER BY ncs DESC
    """, conn)
    print("\n" + "=" * 60)
    print("NCs per launcher class (LLV includes Vulcan; SAS stands alone):")
    print("=" * 60)
    print(cls.to_string(index=False))

    proj = pd.read_sql("""
        SELECT COALESCE(project,'(no project)') AS project, COUNT(*) AS ncs
        FROM nc GROUP BY project ORDER BY ncs DESC
    """, conn)
    print("\nNCs per project:")
    print(proj.to_string(index=False))

    # Owner counts — the number people actually check. Printed every run so a
    # split or a drop is visible immediately rather than reported by an owner.
    own = pd.read_sql("""
        SELECT COALESCE(owner,'(no owner)') AS owner,
               COUNT(*) AS total,
               SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS open,
               SUM(CASE WHEN is_open IS NULL THEN 1 ELSE 0 END) AS no_status
        FROM nc WHERE source IN ('tracker','both')
        GROUP BY owner ORDER BY total DESC
    """, conn)
    print("\n" + "=" * 60)
    print("NCs per owner (from the tracker — SAP has no owner field):")
    print("=" * 60)
    print(own.to_string(index=False))

    # CAPA coverage — how many NCs in the DB actually carry each action type.
    # This is the number the dashboard's coverage KPI reports on.
    if capa is not None and not capa.empty:
        cov = pd.read_sql("""
            SELECT
                (SELECT COUNT(*) FROM nc) AS nc_total,
                COUNT(DISTINCT c.nc_id) AS ncs_with_any_capa,
                COUNT(DISTINCT CASE WHEN c.capa_type='RCA' THEN c.nc_id END) AS ncs_with_rca,
                COUNT(DISTINCT CASE WHEN c.capa_type='CA'  THEN c.nc_id END) AS ncs_with_ca,
                COUNT(DISTINCT CASE WHEN c.capa_type='PA'  THEN c.nc_id END) AS ncs_with_pa
            FROM capa c JOIN nc n ON n.nc_id = c.nc_id
        """, conn)
        print("\n" + "=" * 60)
        print("CAPA coverage (joined to nc):")
        print("=" * 60)
        print(cov.T.to_string(header=False))

        # Orphans: CAPA rows whose NC number doesn't exist in nc. If this is
        # large the foreign key isn't matching and the join is silently lossy.
        orph = pd.read_sql("""
            SELECT COUNT(DISTINCT c.nc_id) AS n
            FROM capa c LEFT JOIN nc n ON n.nc_id = c.nc_id
            WHERE n.nc_id IS NULL
        """, conn).iloc[0]["n"]
        if orph:
            print(f"\n  ⚠ {orph} NC number(s) in the CAPA tracker have no match in `nc`.")
            ex = pd.read_sql("""
                SELECT DISTINCT c.nc_id
                FROM capa c LEFT JOIN nc n ON n.nc_id = c.nc_id
                WHERE n.nc_id IS NULL LIMIT 10
            """, conn)["nc_id"].tolist()
            print(f"    e.g. {ex}")
            print("    Check the nc_id format matches on both sides before trusting coverage.")

    conn.close()
    print(f"\n✓ Database written to: {DB_FILE}")
    print(f"  Tracker: {len(tracker)} rows | SAP: {len(sap)} rows | Merged: {len(merged)} rows")
    if capa is not None and not capa.empty:
        print(f"  CAPA: {len(capa)} rows across {capa['nc_id'].nunique()} NCs")


if __name__ == "__main__":
    main()