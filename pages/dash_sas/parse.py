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

# Later SAS batches can sit on a different PSP prefix; the batch id is still
# the .Q.NNN tail (e.g. W.IC248.Q.021 = Batch 21, or a successor WBS with .Q.022).
_BATCH_RE = re.compile(r"\.Q\.(\d{3})(?:\b|$)", re.I)


def batch_label(wbs) -> str:
    """Any WBS whose tail is .Q.NNN -> 'Batch N'; .900 = Lager; .901 = Springs."""
    m = _BATCH_RE.search(str(wbs))
    if not m:
        return "Unassigned"
    n = int(m.group(1))
    if n == 900:
        return "Lager"
    if n == 901:
        return "Springs"
    return f"Batch {n}"


def origin_label(cause, notif_type) -> str:
    """Supplier vs Production. Binary: anything that is not supplier is production.

    Supplier = SAP Z2 (procurement complaint) OR cause 'Supplier' (a Z3 can
    still be a supplier problem). Production = everything else.
    """
    if str(notif_type).strip().startswith("Z2"):
        return "Supplier"
    if str(cause).strip().lower() == "supplier":
        return "Supplier"
    return "Production"


def is_customer_complaint(cause) -> bool:
    return str(cause).strip().lower() == "customer complaint"


def tag_lanes(df: pd.DataFrame) -> pd.DataFrame:
    """Stamp origin / lane on every row. lane is what the charts filter on:

    Supplier            — Z2 or cause Supplier
    Customer complaint  — cause Customer complaint (own chart)
    Production          — everything else
    """
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
    """SAP / tracker raw -> Closed / Open / Deleted. Case-insensitive so
    tracker OPEN/CLOSED land in the same buckets as the SAP export."""
    v = str(raw or "").strip()
    if not v or v.lower() in ("nan", "none"):
        return "Open"
    key = v.title()
    if key in ("Closed", "Open", "Deleted"):
        return key
    return "Open"


def class_label(v) -> str:
    """Major = defect class 4 or 5. Minor = 0-3. '-'/blank = Unclassified."""
    s = str(v).strip()
    if s in ("4", "5"):
        return "Major"
    if s in ("0", "1", "2", "3"):
        return "Minor"
    return "Unclassified"


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

    df = raw[list(COLS)].rename(columns=COLS).copy()
    df = df[df["project"].astype(str).str.strip() == project].copy()
    if df.empty:
        raise ValueError(f"no rows for project '{project}' in this export")

    # derivations
    df["batch"] = df["wbs"].map(batch_label)
    df["status"] = df["status_raw"].map(status_label)
    df["defect_class_label"] = df["defect_class"].map(class_label)
    df["vendor_clean"] = df["vendor"].map(vendor_norm)
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

    return df.reset_index(drop=True)