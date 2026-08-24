"""
SAS NC export -> clean DataFrame.

One source: the "NC's Overview" SAP export (35 columns, same schema as
SAP_Raw_DATA in the reference Dashboard.xlsx). Row 1 is a 'Totals' band and is
dropped. The page is scoped to Project Text = 'KDS SAS Emmen'.

All derivations (batch label, status, defect class, cost) live here so the
ingest, the API and any future page share exactly one definition.
"""
from __future__ import annotations
import re
import pandas as pd

PROJECT = "KDS SAS Emmen"

# canonical column names we rely on downstream (raw export header -> our name)
COLS = {
    "Notification": "notification",
    "Notification Type": "notif_type",
    "Notif. Year": "notif_year",
    "Notif. Status": "status_raw",
    "WBS Element (Notification)": "wbs",
    "WBS Text (Notification)": "wbs_text",
    "Defect Class": "defect_class",
    "Defect Code TEXT": "defect_code",
    "Disposition Action TEXT": "disposition",
    "Notification Cause C TEXT": "cause",
    "Notification TEXT": "notif_text",
    "Material": "material",
    "Batch": "batch_sap",
    "Model": "model",
    "[Vendor NC]": "vendor",
    "Vendor Key NC": "vendor_key",
    "Notification Date": "opened",
    "Closing date": "closed",
    "Leadtime": "leadtime",
    "CoPQ (NC)": "copq_raw",
    "WBS Element Id (CoPQ)": "copq_wbs",
    "Project Text  (Notification)": "project",
}

# --- how a batch is identified, per system -------------------------------
# SAP  : the 3-digit tail of the PSP element ...Q.NNN  (900 = Lager, 901 = Springs)
#        or a successor WBS ending -NNN
# Teamcenter : the segment AFTER the '-90' level, e.g. R-3011-00045-90-15 = Batch 15
#        Confirmed by the SAS owner 2026-08 and cross-checked against the NC
#        titles: WBS and title agree on batches 4, 22 and 23. Where they
#        disagree the WBS is correct - a title can mention a PREVIOUS batch
#        (IR-001601 says "batch 13" in the text but sits on -90-15), and
#        R-3011-00045-90 with nothing after it is the level, NOT batch 90.
_BATCH_RE = re.compile(r"\.Q\.(\d{3})(?:\b|$)", re.I)
_BATCH_TAIL_RE = re.compile(r"[-.](\d{3})$")
_BATCH_TC_RE = re.compile(r"-90-(\d{1,2})(?:-|$)")   # first number after -90-, sub-levels may follow
_TC_ID_RE = re.compile(r"(IR-\d{4,8})")
_NC_CODE_RE = re.compile(r"(NC_\d{10,16}_\d{6,12})")


def _clean_tc_id(nc: str) -> str:
    """IR-001579/A;1-NC_178... -> IR-001579. SAP ids pass through unchanged."""
    m = _TC_ID_RE.match(str(nc))
    return m.group(1) if m else str(nc)


def _extract_nc_code(obj: str) -> str:
    """Extract the NC_<epoch>_<number> from a TC object string.
    Returns the full NC code, or empty string if not present."""
    m = _NC_CODE_RE.search(str(obj))
    return m.group(1) if m else ""


def _extract_s4_notif(obj: str) -> str:
    """Extract the S/4HANA notification number (trailing digits of the NC code).
    IR-001660/A;1-NC_1786015699989_170000019477 -> 170000019477"""
    m = _NC_CODE_RE.search(str(obj))
    if m:
        parts = m.group(1).split("_")
        return parts[-1] if len(parts) >= 3 else ""
    return ""


def batch_label(wbs) -> str:
    """WBS / PSP element -> Batch N / Lager / Springs / Unassigned.

    Nothing is ever read from free text. If the WBS does not carry a batch the
    answer is 'Unassigned', not a guess.
    """
    s = str(wbs).strip()
    if not s or s.lower() in ("nan", "none", "-"):
        return "Unassigned"

    m = _BATCH_RE.search(s)                      # SAP  W.IC248.Q.NNN
    if m:
        n = int(m.group(1))
        if n == 900:
            return "Lager"
        if n == 901:
            return "Springs"
        return f"Batch {n}"

    m = _BATCH_TC_RE.search(s)                   # Teamcenter  R-...-90-NN
    if m:
        return f"Batch {int(m.group(1))}"

    m = _BATCH_TAIL_RE.search(s)                 # successor WBS ending -NNN
    if m:
        n = int(m.group(1))
        if n == 900:
            return "Lager"
        if n == 901:
            return "Springs"
        return f"Batch {n}"

    return "Unassigned"


def origin_label(cause, notif_type) -> str:
    """Supplier vs Production. Binary: anything not supplier is production."""
    if str(notif_type).strip().startswith("Z2"):
        return "Supplier"
    if str(cause).strip().lower() == "supplier":
        return "Supplier"
    return "Production"


def is_customer_complaint(cause) -> bool:
    return str(cause).strip().lower() == "customer complaint"


def tag_lanes(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["origin"] = [origin_label(c, t)
                     for c, t in zip(out["cause"], out["notif_type"])]
    out["is_cc"] = out["cause"].map(is_customer_complaint)
    out["lane"] = ["Customer complaint" if cc else o
                   for cc, o in zip(out["is_cc"], out["origin"])]
    return out


def batch_sort_key(label: str):
    """Order: Batch 1,2,... then Springs, Lager, Unassigned last."""
    m = re.match(r"Batch (\d+)", label)
    if m:
        return (0, int(m.group(1)))
    return {"Springs": (1, 0), "Lager": (2, 0)}.get(label, (3, 0))


def status_label(raw) -> str:
    """SAP raw -> display. Open shows as WIP in the batch/supplier charts;
    here we keep the three real states and let the UI relabel if wanted."""
    v = str(raw).strip()
    if v in ("Closed", "Open", "Deleted"):
        return v
    return "Open"


def class_label(v) -> str:
    """Major = defect class 4 or 5, or Teamcenter *-Major severity.
    Minor = 0-3, or TC *-Minor / *-Observation.
    '-'/blank = Unclassified."""
    s = str(v).strip()
    # Excel/pandas turns 2 into '2.0' — strip the trailing .0 so it matches
    if s.endswith(".0"):
        s = s[:-2]
    if s in ("4", "5"):
        return "Major"
    if s in ("0", "1", "2", "3"):
        return "Minor"
    # Teamcenter severity text, in case a file arrives unconverted
    low = s.lower()
    import re as _re
    if _re.match(r"^(low|medium|high)-major$", low):
        return "Major"
    if _re.match(r"^(low|medium|high)-(minor|observation)$", low):
        return "Minor"
    return "Unclassified"


# ---------------------------------------------------------------------------
# Defect-code vocabulary: SAP and Teamcenter word the SAME defect differently.
# Left untranslated they split into two bars on every defect chart. Only
# unambiguous pairs are merged here; anything needing a judgement call is left
# alone so a human decides it, not this file.
#   Adriele confirmed 2026-08: merge the three below.
#   NOT merged (needs Quality to rule): Teamcenter 'Surface finish and
#   appearance defects' vs SAP 'Inaccurate surface properties'.
# ---------------------------------------------------------------------------
DEFECT_SYNONYMS = {
    "material not compliant with specification": "Material not compliant with spec.",
    "manufacturing / assembly execution errors": "Manufacturing / assembly",
    "function / performance insufficient or incomplete":
        "Function/performance insuf. or incompl.",
}


def defect_norm(v) -> str:
    """Collapse known cross-system wordings of the same defect onto one label."""
    s = str(v or "").strip()
    if not s or s.lower() in ("nan", "none"):
        return s
    return DEFECT_SYNONYMS.get(re.sub(r"\s+", " ", s).lower(), s)


# Cause values arrive with inconsistent capitalization from different systems
# ("Customer Complaint" vs "Customer complaint"). Merge them to one canonical
# spelling so a single cause does not split into two rows on every chart.
CAUSE_CANON = {
    "customer complaint": "Customer Complaint",
    "incoming inspection": "Incoming Inspection",
    "final inspection": "Final Inspection",
    "not assigned": "Not assigned",
}


def cause_norm(v) -> str:
    s = str(v or "").strip()
    if not s or s.lower() in ("nan", "none"):
        return s
    return CAUSE_CANON.get(s.lower(), s)


def vendor_norm(v) -> str:
    """Merge case/whitespace duplicates (e.g. the two 'Beyond Gravity Sweden')."""
    s = str(v).strip()
    if s in ("", "nan", "None", "ES", "-"):
        return "Not recorded"
    return re.sub(r"\s+", " ", s).upper()


def _to_num(series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _to_date(series) -> pd.Series:
    """Handle both real datetimes and Excel serial numbers (e.g. 41988).
    Serials are days since the 1899-12-30 Excel epoch."""
    s = series.copy()
    num = pd.to_numeric(s, errors="coerce")
    # values that are plausible Excel serials (roughly 1990..2050 => 32000..55000)
    is_serial = num.notna() & (num > 20000) & (num < 80000)
    out = pd.to_datetime(s, errors="coerce")
    if is_serial.any():
        serial_dates = pd.to_datetime(num[is_serial], unit="D",
                                      origin="1899-12-30")
        out.loc[is_serial] = serial_dates
    return out


def load(path_or_buffer, project: str = PROJECT) -> pd.DataFrame:
    """Read the export, drop the Totals band, scope to the project, derive."""
    raw = pd.read_excel(path_or_buffer, sheet_name=0)

    # the export sometimes names the status column twice -> 'Notif. Status.1'
    if "Notif. Status" not in raw.columns and "Notif. Status.1" in raw.columns:
        raw = raw.rename(columns={"Notif. Status.1": "Notif. Status"})

    # drop the Totals band (first data row where Plant == 'Totals')
    if "Plant" in raw.columns:
        raw = raw[raw["Plant"].astype(str).str.strip() != "Totals"]

    missing = [c for c in COLS if c not in raw.columns]
    if missing:
        raise ValueError(f"export missing columns: {missing}")
    # extra columns (e.g. Source from the builder) are kept in raw for
    # later use but not included in the rename — they pass through intact.

    df = raw[list(COLS)].rename(columns=COLS).copy()
    df = df[df["project"].astype(str).str.strip() == project].copy()
    if df.empty:
        raise ValueError(f"no rows for project '{project}' in this export")

    # derivations
    df["batch"] = df["wbs"].map(batch_label)
    df["status"] = df["status_raw"].map(status_label)
    df["defect_class_label"] = df["defect_class"].map(class_label)
    df["vendor_clean"] = df["vendor"].map(vendor_norm)
    df["defect_code"] = df["defect_code"].map(defect_norm)
    df["cause"] = df["cause"].map(cause_norm)
    df["origin"] = [origin_label(c, t) for c, t in zip(df["cause"], df["notif_type"])]
    df["is_cc"] = df["cause"].map(is_customer_complaint)

    df["opened"] = _to_date(df["opened"])
    df["closed"] = _to_date(df["closed"])
    df["leadtime"] = _to_num(df["leadtime"])
    # CoPQ is booked negative in SAP; store the positive cost magnitude
    df["copq"] = _to_num(df["copq_raw"]).fillna(0).abs()
    df["copq_booked"] = df["copq_wbs"].astype(str).str.strip().ne("-") & df["copq"].gt(0)

    df["month"] = df["opened"].dt.strftime("%Y-%m")
    df["notification"] = df["notification"].astype(str).str.replace(r"\.0$", "", regex=True)
    # Clean TC object strings: IR-001579/A;1-NC_178... -> IR-001579
    df["notification"] = df["notification"].map(_clean_tc_id)

    # If the builder included a Source column, carry it through as source_file
    # so the original document names survive instead of being overwritten by
    # the upload filename.
    if "Source" in raw.columns:
        df["source_file"] = raw.loc[df.index, "Source"].values

    # Status Note documents special status handling (e.g. TC return-to-work)
    if "Status Note" in raw.columns:
        df["status_note"] = raw.loc[df.index, "Status Note"].values

    # Extract the NC code and S/4HANA notification from TC object strings.
    # These come from the ORIGINAL Notification column (before cleaning),
    # which is in the raw frame, not df.
    raw_notif = raw.loc[df.index, "Notification"].astype(str)
    nc_codes = raw_notif.map(_extract_nc_code)
    s4_notifs = raw_notif.map(_extract_s4_notif)
    if nc_codes.any():
        df["nc_code"] = nc_codes.values
    if s4_notifs.any():
        df["s4_notification"] = s4_notifs.values

    return df.reset_index(drop=True)