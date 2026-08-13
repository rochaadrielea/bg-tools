"""
dashboard_body.py — the Quality BRM dashboard body (rendered inside the
"Quality Dashboard" tab by app.py). Page config + auth live in app.py.

Interactive charts with tooltips explaining every measurement and data source.
Compatible with both old (tracker-only) and new (merged) quality.db schemas.

2026-08 changes (Adriele):
  * Default window = THIS YEAR -> today (auto), with a 'This year' preset.
  * New section "Where NCs happen" (by project & class): bar, open-only bar,
    class->project treemap, project x detection-area heatmap.
  * EVERY '📥 Excel' now downloads the ORIGINAL raw NC rows for that view — not a
    chart summary — plus a drill so you can pull the raw rows for a sub-slice.
"""
from datetime import datetime, date
from io import BytesIO
from pathlib import Path
import sqlite3
import math

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# auth is handled by app.py (this body renders inside the Dashboard tab)

DB_FILE = "quality.db"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1400px; }
    [data-testid="stMetricValue"] { font-size: 1.8rem; }
    [data-testid="stMetricLabel"] { font-size: 0.85rem; }
    .main-header {
        background: linear-gradient(90deg, #1E2761 0%, #F26E21 100%);
        padding: 0.8rem 1.5rem; border-radius: 8px; color: white; margin-bottom: 1rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.5rem; }
    .main-header p { color: #e0e0e0; margin: 0; font-size: 0.9rem; }
    .nav-link {
        display: block; padding: 0.4rem 0.6rem; margin: 0.2rem 0;
        border-radius: 4px; text-decoration: none; color: #1E2761;
        border-left: 3px solid transparent;
    }
    .nav-link:hover { background: #f0f2f6; border-left-color: #F26E21; }
    .filter-active {
        background: #FFF3E0; border: 1px solid #F26E21; border-radius: 6px;
        padding: 0.3rem 0.6rem; font-size: 0.8rem; color: #E65100; margin-top: 0.5rem;
    }
    /* Highlight the reach-zero calculator section so it draws attention */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div.reach-zero-anchor) {
        background: #EAF2FB;
        border: 2px solid #2a78d6 !important;
        border-radius: 12px;
        padding: 0.5rem 0.75rem;
    }
    /* Month picker section box */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div.month-anchor) {
        background: #FFF6EC;
        border: 2px solid #F26E21 !important;
        border-radius: 12px;
        padding: 0.5rem 0.75rem;
    }
    /* Drill / raw-data section box */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(div.drill-anchor) {
        background: #EEF6EE;
        border: 2px solid #4CAF50 !important;
        border-radius: 12px;
        padding: 0.5rem 0.75rem;
    }
    /* KPI cards with visible descriptions */
    .kpi-card {
        padding: 0.4rem 0.2rem 0.9rem 0;
    }
    .kpi-label {
        font-size: 0.82rem; color: #5B6B78; font-weight: 600; margin-bottom: 0.1rem;
    }
    .kpi-value {
        font-size: 1.9rem; color: #1E2761; font-weight: 700; line-height: 1.1;
    }
    .kpi-desc {
        font-size: 0.78rem; color: #7A8894; margin-top: 0.25rem; line-height: 1.35;
    }
    /* Filter bar (the only expander on the page): dark-blue header, bold, bigger,
       so everyone reads it as THE filter. */
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] > details > summary,
    .streamlit-expanderHeader {
        background: #1E2761 !important;
        border-radius: 8px !important;
        padding: 0.7rem 1rem !important;
    }
    [data-testid="stExpander"] summary:hover,
    .streamlit-expanderHeader:hover { background: #2a3170 !important; }
    [data-testid="stExpander"] summary p,
    [data-testid="stExpander"] summary span,
    [data-testid="stExpander"] summary div,
    .streamlit-expanderHeader p {
        color: #ffffff !important;
        font-weight: 800 !important;
        font-size: 1.12rem !important;
    }
    [data-testid="stExpander"] summary svg,
    .streamlit-expanderHeader svg { fill: #ffffff !important; color: #ffffff !important; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="main-header">
        <h1>BU Launchers Switzerland — Quality</h1>
        <p>Business Review Dashboard</p>
    </div>
""", unsafe_allow_html=True)

if not Path(DB_FILE).exists():
    st.error(f"`{DB_FILE}` not found. Run `python ingest.py` first.")
    st.stop()


def _q(sql, params=None):
    with sqlite3.connect(DB_FILE) as conn:
        return pd.read_sql(sql, conn, params=params or [])


def _has_column(col_name, table="nc"):
    """Check if a column exists in a table."""
    try:
        cols = _q(f"PRAGMA table_info({table})")
        return col_name in cols["name"].values
    except Exception:
        return False


def _filter_context_rows(include_since=False):
    """Context block describing when the report was taken and which filters were active.
    Reads the committed filter values from session_state, so every export self-documents."""
    ss = st.session_state
    rows = [
        ("Report generated", datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Data source", DB_FILE),
        ("Data last modified",
         datetime.fromtimestamp(Path(DB_FILE).stat().st_mtime).strftime("%d.%m.%Y %H:%M")
         if Path(DB_FILE).exists() else "n/a"),
        ("Period From", str(ss.get("date_from", ""))),
        ("Period To", str(ss.get("date_to", ""))),
    ]
    if include_since:
        rows.append(("Since (burndown anchor)", str(ss.get("since_date", ""))))
    rows += [
        ("NC type", str(ss.get("nc_type", "All"))),
        ("Status", str(ss.get("nc_status", "All"))),
        ("Project", ", ".join(ss.get("nc_project", [])) or "All"),
        ("Owner", ", ".join(ss.get("nc_owner", [])) or "All"),
        ("Data source filter", str(ss.get("nc_source", "All"))),
    ]
    return pd.DataFrame(rows, columns=["Field", "Value"])


def to_excel_bytes(df, sheet_name="Data", include_since=False):
    """Export a dataframe with a filter-context header block above the data."""
    buf = BytesIO()
    ctx = _filter_context_rows(include_since=include_since)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Context block at the top, data below it
        ctx.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=len(ctx) + 2)
        ws = writer.sheets[sheet_name]
        for _c in ("A", "B", "C", "D", "E", "F"):
            ws.column_dimensions[_c].width = 24
    return buf.getvalue()


def build_full_report_bytes(datasets, include_since=True):
    """One workbook: a Filters/Context sheet + one sheet per chart dataset."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        _filter_context_rows(include_since=include_since).to_excel(
            writer, sheet_name="Filters", index=False)
        writer.sheets["Filters"].column_dimensions["A"].width = 26
        writer.sheets["Filters"].column_dimensions["B"].width = 40
        for name, d in datasets.items():
            if d is None or (hasattr(d, "empty") and d.empty):
                continue
            sheet = name[:31]
            d.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for _c in ("A", "B", "C", "D", "E", "F"):
                ws.column_dimensions[_c].width = 22
    return buf.getvalue()


CHART_CONFIG = {
    # 'hover' keeps the zoom/pan toolbar out of the plot area until you point at
    # the chart — as 'True' it renders on top of the tallest bars and hides them.
    "displayModeBar": "hover",
    "toImageButtonOptions": {
        "format": "png",
        "height": 600,
        "width": 1000,
        "scale": 2,
    },
    "displaylogo": False,
}

# Top margin reserved for the Plotly modebar so it never overlaps the top bar.
MODEBAR_T = 40

# Check schema once at startup
HAS_SOURCE = _has_column("source")
# Three-state status: blank Status is neither open nor closed (is_open IS NULL).
# Older DBs built before this change have no status_state column.
HAS_STATUS_STATE = _has_column("status_state")


def _has_table(name):
    try:
        t = _q("SELECT name FROM sqlite_master WHERE type='table' AND name=?", [name])
        return not t.empty
    except Exception:
        return False


HAS_CAPA = _has_table("capa")
HAS_COPQ = _has_column("copq")
# capa_type is what makes RCA / CA / PA reporting possible. Older DBs built by
# the RCA-only ingest have no such column — coverage hides itself in that case.
HAS_CAPA_TYPE = HAS_CAPA and _has_column("capa_type", "capa")


# ------------------------------------------------------------------------------
# FILTER SETUP (no UI here). The on-page filter itself is rendered lower down,
# right above "Open NCs by Area", and only governs the sections BELOW it.
# Sections 1–4 (backlog, burn-down, YTD, monthly) are a fixed overview: they
# always show THIS YEAR → today and are NOT touched by the filter.
# ------------------------------------------------------------------------------
# Earliest NC in the data — used as the lower bound of the date pickers.
_min_row = _q("SELECT MIN(created_on) AS m FROM nc WHERE created_on IS NOT NULL")
_earliest = None
if not _min_row.empty and _min_row.iloc[0]["m"]:
    try:
        _earliest = pd.to_datetime(_min_row.iloc[0]["m"]).date()
    except Exception:
        _earliest = None
_pick_min = _earliest or date(2015, 1, 1)
_pick_max = date(date.today().year + 2, 12, 31)
_default_from = max(_pick_min, date(date.today().year, 1, 1))

# ---- Defaults for the filter (committed values live in session_state) ----
_filter_defaults = {
    "date_from": _default_from,
    "date_to": date.today(),
    "since_date": date(2026, 6, 1),
    "nc_type": "All",
    "nc_status": "All",
    "nc_project": [],
    "nc_owner": [],
    "nc_source": "All",
}
for _k, _v in _filter_defaults.items():
    st.session_state.setdefault(_k, _v)

# Keep the filter's To date rolling with the calendar (unchanged behaviour).
_to_anchor = st.session_state.get("_to_anchor_day")
if (_to_anchor is not None and _to_anchor != date.today().isoformat()
        and not st.session_state.get("_to_manual")
        and st.session_state.get("date_to") != date.today()):
    st.session_state["date_to"] = date.today()
    st.session_state.pop("w_date_to", None)
st.session_state["_to_anchor_day"] = date.today().isoformat()

_projects = _q("SELECT DISTINCT COALESCE(project,'(no project)') AS p FROM nc ORDER BY p")["p"].tolist()
_owners = _q("SELECT DISTINCT COALESCE(owner,'(no owner)') AS o FROM nc ORDER BY o")["o"].tolist()
_sources = []
if HAS_SOURCE:
    _sources = _q("SELECT DISTINCT source FROM nc WHERE source IS NOT NULL ORDER BY source")["source"].tolist()

# ---- FIXED overview values for sections 1–4 (the filter does not reach them) ----
date_from = _default_from
date_to = date.today()
if date_from > date_to:
    date_from, date_to = date_to, date_from
nc_type = "All"
nc_status = "All"
nc_project = []
nc_owner = []
nc_source = "All"
exercise_start = st.session_state["since_date"]
exercise_start = max(date_from, min(exercise_start, date_to))


# ------------------------------------------------------------------------------
# Filter builder
# ------------------------------------------------------------------------------
def build_where(extra=None, date_col="created_on", use_date=True):
    cl, pr = [], []
    if use_date and date_from:
        cl.append(f"{date_col} >= ?"); pr.append(str(date_from))
    if use_date and date_to:
        cl.append(f"{date_col} <= ?"); pr.append(str(date_to) + " 23:59:59")
    if nc_type == "Production": cl.append("is_supplier_nc = 0")
    elif nc_type == "Supplier": cl.append("is_supplier_nc = 1")
    if nc_status == "Open": cl.append("is_open = 1")
    elif nc_status == "Closed": cl.append("is_open = 0")
    elif nc_status == "(no status)": cl.append("is_open IS NULL")
    if nc_project:
        cl.append(f"COALESCE(project,'(no project)') IN ({','.join(['?']*len(nc_project))})")
        pr.extend(nc_project)
    if nc_owner:
        cl.append(f"COALESCE(owner,'(no owner)') IN ({','.join(['?']*len(nc_owner))})")
        pr.extend(nc_owner)
    if HAS_SOURCE and nc_source != "All":
        cl.append("source = ?"); pr.append(nc_source)
    if extra:
        for s, p in extra: cl.append(s); pr.extend(p)
    return ("WHERE " + " AND ".join(cl)) if cl else "", pr


def _qf(tmpl, extra=None, date_col="created_on", use_date=True):
    w, p = build_where(extra, date_col, use_date)
    return _q(tmpl.replace("{WHERE}", w), p)


def raw_nc(extra=None, use_date=True):
    """The ORIGINAL NC rows (all columns) for the active filter (+ optional extra).
    Every '📥 Excel' hands over these raw records — never a chart summary
    (Adriele: the person wants the data, not the picture's numbers)."""
    w, p = build_where(extra, use_date=use_date)
    return _q(f"SELECT * FROM nc {w} ORDER BY created_on DESC", p)


_DATE_COLS = {"created_on", "closure_date", "disposition_date",
              "open_date", "close_date", "due_date"}


def _fmt_dates_ddmmyyyy(df):
    """Show every date column as DD/MM/YYYY (Adriele's standard). Applies to the
    on-screen table AND the download, so both read the same European format."""
    d = df.copy()
    for c in d.columns:
        if c in _DATE_COLS:
            d[c] = pd.to_datetime(d[c], errors="coerce").dt.strftime("%d/%m/%Y")
            d[c] = d[c].where(d[c].notna(), "")
    return d


def raw_block(df, key, fname="nc_rows.xlsx", sheet="NC_rows", include_since=False,
              dl_label="Download (Excel)", note=None, preview_rows=200):
    """End-of-chart controls: a '👁 View raw data' toggle that shows the rows
    right on the dashboard, plus a '📥 Download' button — both from the SAME df,
    i.e. the raw NC rows for THIS chart's current filter/slice. What you see is
    exactly what you download."""
    df = _fmt_dates_ddmmyyyy(df)   # DD/MM/YYYY everywhere in the raw output
    n = len(df)
    # No st.columns here on purpose: this block is dropped inside chart columns
    # too, and columns-in-columns is fragile. Toggle then button, stacked, always
    # renders — whether the block sits at page width or inside a half-width chart.
    show = st.toggle(f"👁 View raw data ({n})", key=f"{key}__view",
                     help=f"Show the {n} raw NC rows behind this chart, on screen.")
    st.download_button(f"📥 {dl_label} ({n})",
                       to_excel_bytes(df, sheet, include_since=include_since),
                       fname, XLSX_MIME, key=f"{key}__dl")
    if show:
        if note:
            st.caption(note)
        if n:
            # Show empty cells as blank, not the literal text "None"/"NaN". A
            # missing value (no flight unit, no owner, …) is blank, not data.
            _disp = df.head(preview_rows).fillna("")
            st.dataframe(_disp, width='stretch', hide_index=True,
                         height=min(400, 80 + 28 * min(n, 11)))
            if n > preview_rows:
                st.caption(f"Showing the first {preview_rows} of {n} rows — the download has all {n}.")
        else:
            st.info("No rows for this selection.")


def _raw_dl(label, extra=None, use_date=True, fname="nc_rows.xlsx", key=None,
            include_since=False):
    """View + download the raw NC rows for a view/slice (uses the active filter)."""
    df = raw_nc(extra, use_date=use_date)
    raw_block(df, key=key or "raw", fname=fname, include_since=include_since,
              dl_label=label.replace("Excel — ", "").replace("Excel —", "").strip() or "Download (Excel)")


def _headroom(fig, values, horizontal=False, pad=0.18):
    """Leave room for the value labels that sit outside the bars.

    Plotly ends the axis at the tallest bar, so a label written outside that bar
    has nowhere to go and is clipped — the biggest number on the chart, the one
    people actually read, is the one that gets cut in half. Extending the axis
    past the maximum fixes it for every bar at once.
    """
    try:
        _max = float(max(values)) if len(values) else 0.0
    except (TypeError, ValueError):
        return fig
    if _max <= 0:
        return fig
    _rng = [0, _max * (1 + pad)]
    fig.update_layout(xaxis=dict(range=_rng)) if horizontal else \
        fig.update_layout(yaxis=dict(range=_rng))
    return fig


def _kpi_card(label, value, desc):
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-desc">{desc}</div>
    </div>
    """

# ------------------------------------------------------------------------------
# Shared filter clauses
# ------------------------------------------------------------------------------
# ---- Burndown filter: applies Project/Owner/NC-type/Status/Data-source ----
def _bd_filter(include_dates=True):
    cl, pr = [], []
    if include_dates and date_from:
        cl.append("created_on >= ?"); pr.append(str(date_from))
    if include_dates and date_to:
        cl.append("created_on <= ?"); pr.append(str(date_to) + " 23:59:59")
    if nc_type == "Production": cl.append("is_supplier_nc = 0")
    elif nc_type == "Supplier": cl.append("is_supplier_nc = 1")
    if nc_status == "Open": cl.append("is_open = 1")
    elif nc_status == "Closed": cl.append("is_open = 0")
    elif nc_status == "(no status)": cl.append("is_open IS NULL")
    if nc_project:
        cl.append(f"COALESCE(project,'(no project)') IN ({','.join(['?']*len(nc_project))})")
        pr.extend(nc_project)
    if nc_owner:
        cl.append(f"COALESCE(owner,'(no owner)') IN ({','.join(['?']*len(nc_owner))})")
        pr.extend(nc_owner)
    if HAS_SOURCE and nc_source != "All":
        cl.append("source = ?"); pr.append(nc_source)
    return cl, pr

# Clauses for the OVERVIEW sections 1–4 (fixed This-year window; the on-page
# filter lower down recomputes these for sections 5–10).
_BF_CL, _BF_PR = _bd_filter(include_dates=True)
# Non-date filter — for the BACKLOG KPIs, which describe NCs created BEFORE the
# 'Since' anchor.
_NB_CL, _NB_PR = _bd_filter(include_dates=False)
_filter_sig = f"{date_from}|{date_to}|{nc_type}|{nc_status}|{'-'.join(sorted(nc_project))}|{'-'.join(sorted(nc_owner))}|{nc_source}"


# ------------------------------------------------------------------------------
# 1 - OPEN BACKLOG: SAP vs the new system
# ------------------------------------------------------------------------------
st.markdown('<div id="backlog-system"></div>', unsafe_allow_html=True)
st.subheader("Open backlog — SAP vs new system")
st.caption("Where the open NCs actually sit. SAP ECC is the old system being burnt down, "
           "EZYone is the new one filling up, and Blackout NCs are waiting for their work "
           "order to be recreated in EZYone before they can migrate. Ignores the From/To "
           "dates — this is the position today.")

_SYS_COLORS = {"SAP": "#1E2761", "EZ1": "#F26E21", "Blackout": "#4CAF50"}
_SYS_LABEL = {"SAP": "SAP", "EZ1": "EZYone", "Blackout": "Blackout"}
_SF = (" AND " + " AND ".join(_NB_CL)) if _NB_CL else ""

df_sysopen = _q("SELECT COALESCE(system,'(unknown)') AS system, COUNT(*) AS open_ncs "
                "FROM nc WHERE is_open=1" + _SF + " GROUP BY system", _NB_PR)
_sys_now = dict(zip(df_sysopen["system"], df_sysopen["open_ncs"])) if not df_sysopen.empty else {}

_sys_cards = [
    ("Open in SAP", int(_sys_now.get("SAP", 0)), "The burn-down target — NCs still open in SAP ECC."),
    ("Open in EZYone", int(_sys_now.get("EZ1", 0)), "Created in the new system since go-live."),
    ("Open in Blackout", int(_sys_now.get("Blackout", 0)),
     "Raised during the blackout period. Cannot migrate until the work order exists in EZYone."),
    ("Total open", int(df_sysopen["open_ncs"].sum()) if not df_sysopen.empty else 0,
     "All three together."),
]
for _c, (_l, _v, _d) in zip(st.columns(4), _sys_cards):
    _c.markdown(_kpi_card(_l, _v, _d), unsafe_allow_html=True)

st.markdown("**Open at month-end, by system**")
_bl_months = pd.date_range(max(date_from, date(2025, 12, 1)).replace(day=1),
                           min(date_to, date.today()), freq="ME")
_bl_rows = []
for _m in _bl_months:
    _me = _m.strftime("%Y-%m-%d")
    _d = _q("SELECT COALESCE(system,'(unknown)') AS system, COUNT(*) AS n FROM nc "
            "WHERE created_on <= ? AND (is_open=1 OR closure_date > ?)" + _SF
            + " GROUP BY 1",
            [_me, _me] + _NB_PR)
    _row = {"month": _m.strftime("%Y-%m")}
    _row.update(dict(zip(_d["system"], _d["n"])))
    _bl_rows.append(_row)
df_backlog_sys = pd.DataFrame(_bl_rows).fillna(0)

if df_backlog_sys.empty or len(df_backlog_sys) < 2:
    st.info("Not enough months in the selected range to draw the trend.")
else:
    figb = go.Figure()
    for _sys in ["SAP", "EZ1", "Blackout"]:
        if _sys not in df_backlog_sys.columns:
            continue
        figb.add_trace(go.Bar(
            x=df_backlog_sys["month"], y=df_backlog_sys[_sys].astype(int),
            name=_SYS_LABEL[_sys], marker_color=_SYS_COLORS[_sys],
            hovertemplate="<b>%{x}</b><br>" + _SYS_LABEL[_sys] + " open: %{y}<extra></extra>"))
    figb.update_layout(height=330, margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                       barmode="stack", legend=dict(orientation="h", y=-0.15),
                       xaxis_title="", yaxis_title="Open NCs")
    st.plotly_chart(figb, width='stretch', config=CHART_CONFIG, key="backlog_by_system")

    _first, _last = df_backlog_sys.iloc[0], df_backlog_sys.iloc[-1]
    def _mv(row, col):
        return int(row[col]) if col in df_backlog_sys.columns else 0
    _sap_d = _mv(_last, "SAP") - _mv(_first, "SAP")
    st.caption(
        f"Stacked, so the top of each bar is the total open that month. "
        f"SAP went from {_mv(_first,'SAP')} to {_mv(_last,'SAP')} ({_sap_d:+d}) over this window, "
        f"while EZYone went from {_mv(_first,'EZ1')} to {_mv(_last,'EZ1')} and Blackout from "
        f"{_mv(_first,'Blackout')} to {_mv(_last,'Blackout')}. "
        + ("The SAP pile is shrinking, but the new system fills faster than SAP empties, "
           "so the total still climbs." if _sap_d < 0 else
           "Both the old and the new pile are growing."))
    # RAW download: every open NC behind the backlog (ignores From/To, like the chart)
    _raw_dl("Excel — raw open NCs", extra=[("is_open=1", [])], use_date=False,
            fname="open_backlog_raw.xlsx", key="dl_backlog_sys")

st.markdown("")


# ------------------------------------------------------------------------------
# BURNDOWN TRACKER
# ------------------------------------------------------------------------------
st.markdown('<div id="burndown"></div>', unsafe_allow_html=True)

# ---- Independent date control for THIS section only ---------------------------
# The burndown is its own tool: it needs to be able to look back before January
# (there is data from earlier), so it gets its own From date. The end is always
# today. It does NOT use the page filter and does NOT affect any other section.
_ov_from, _ov_to = date_from, date_to               # remember the page window
_ov_BF_CL, _ov_BF_PR = _BF_CL, _BF_PR
_bd_from_default = date(2025, 6, 1)   # default start of the burndown history
_bd_from_val = st.session_state.get("bd_from", _bd_from_default)
_bd_from_val = min(max(_bd_from_val, _pick_min), date.today())  # keep in range
st.markdown('<div class="month-anchor"></div>', unsafe_allow_html=True)
with st.container(border=True):
    _bc1, _bc2, _bc3 = st.columns([1.2, 1, 3])
    bd_from = _bc1.date_input(
        "From (this section only)", value=_bd_from_val, min_value=_pick_min,
        max_value=date.today(), key="bd_from", format="DD/MM/YYYY",
        help="Start date for the burndown only — defaults to 01/06/2025, and "
             "you can go back further to include older NCs. The end is always today.")
    _bc2.markdown("<div style='height:1.8rem'></div>", unsafe_allow_html=True)
    if _bc2.button("Reset", key="bd_reset",
                   disabled=(bd_from == _bd_from_default),
                   help="Snap back to 01 January 2025."):
        st.session_state.pop("bd_from", None)
        st.rerun()
    _bc3.markdown(
        "<div style='height:0.4rem'></div>"
        f"<span style='font-size:0.85rem;color:#5B6B78'>Burndown window: "
        f"<b>{bd_from.strftime('%d %b %Y')}</b> → <b>today</b>. Independent of the "
        "page filter.</span>", unsafe_allow_html=True)

# apply the burndown window locally (rebind the globals its queries read)
date_from = bd_from
date_to = date.today()
_BF_CL, _BF_PR = _bd_filter(include_dates=True)

st.subheader(f"NC Burndown Tracker · since {exercise_start.strftime('%d %B %Y')}")
st.caption(f"⚓ **Since {exercise_start.strftime('%d %b %Y')}** = the date the burndown measures from. "
           f"It drives **only the KPI numbers below**. The three *backlog* figures (Backlog at freeze, "
           f"Closed since start, Still open) describe NCs created **before** that date, so they are not "
           f"clipped by the From date — but they do follow the Project / Owner / NC-type filters. "
           f"Every chart below follows this section's window (**{date_from}** → **today**).")

es = str(exercise_start)

def _bd(where_parts, params, use_dates=True):
    """Run a burndown count query with the active filter AND-ed in."""
    _cl, _pr = (_BF_CL, _BF_PR) if use_dates else (_NB_CL, _NB_PR)
    parts = list(where_parts) + _cl
    prm = list(params) + _pr
    sql = "SELECT COUNT(*) AS n FROM nc" + (" WHERE " + " AND ".join(parts) if parts else "")
    return _q(sql, prm).iloc[0]["n"]

backlog_at_start = _bd(["created_on < ?", "(is_open=1 OR closure_date >= ?)"], [es, es], use_dates=False)
closed_from_backlog = _bd(["created_on < ?", "is_open=0", "closure_date >= ?"], [es, es], use_dates=False)
still_open_backlog = _bd(["created_on < ?", "is_open=1"], [es], use_dates=False)
new_since_start = _bd(["created_on >= ?"], [es])
new_still_open = _bd(["created_on >= ?", "is_open=1"], [es])
total_open = _bd(["is_open=1"], [])

today = date.today()
weeks_elapsed = max(1, (today - exercise_start).days / 7)
avg_new_wk = round(new_since_start / weeks_elapsed, 1)
team_size = 11

_es_str = exercise_start.strftime('%d %b %Y')
_kpis_row1 = [
    ("Backlog at freeze", int(backlog_at_start), f"NCs already open on {_es_str} — the starting pile."),
    ("Closed since start", int(closed_from_backlog), f"How many of that original {int(backlog_at_start)} you've since closed."),
    ("Still open (backlog)", int(still_open_backlog), f"Of the original {int(backlog_at_start)}, how many are still open."),
    ("Total open now", int(total_open), "All currently open NCs (old backlog + everything new)."),
]
_cols = st.columns(4)
for _c, (_lab, _val, _desc) in zip(_cols, _kpis_row1):
    _c.markdown(_kpi_card(_lab, _val, _desc), unsafe_allow_html=True)

st.markdown('<div class="reach-zero-anchor"></div>', unsafe_allow_html=True)
with st.container(border=True):
    st.markdown("##### 🎯 Closure rate to reach target")
    sc1, sc2, sc3, sc4 = st.columns([2, 1, 1, 1])
    with sc1:
        target_deadline = st.slider(
            "Deadline", min_value=today,
            max_value=date(today.year + 2, 12, 31),
            value=date(2026, 9, 30), format="DD MMM YYYY",
            help="Drag to change the target date. The per-week / per-month numbers update live.")
    with sc2:
        inflow = st.number_input(
            "New / week (inflow)", min_value=0.0, value=float(avg_new_wk), step=0.1,
            key=f"inflow_{_filter_sig}",
            help=f"Assumed new NCs arriving per week. Default is this selection's average ({avg_new_wk}). Adjust to stress-test.")
    with sc3:
        open_now_input = st.number_input(
            "Open now", min_value=0, value=int(total_open), step=1,
            key=f"opennow_{_filter_sig}",
            help="Current open NC count for this selection. Editable to test scenarios.")
    with sc4:
        target_open = st.number_input(
            "Target open", min_value=0, value=65, step=1,
            key="target_open",
            help="The acceptable number of open NCs to get down to by the deadline (not necessarily zero). "
                 "Change it here if the goal changes.")

    weeks_left = max(1, (target_deadline - today).days / 7)
    _gap = open_now_input - target_open
    _at_target = _gap <= 0
    required_per_week = (max(0, _gap) / weeks_left) + inflow
    required_per_month = required_per_week * 4.33
    weekly_target = math.ceil(required_per_week)
    breakeven = math.ceil(inflow)

    zc1, zc2, zc3 = st.columns(3)
    if _at_target:
        zc1.metric(f"Close / week → {target_open} by {target_deadline.strftime('%d %b %Y')}",
                   f"{breakeven} NCs",
                   help=f"Already at or below the target ({open_now_input} open = {target_open}). "
                        f"Closing {breakeven}/week just matches inflow and holds the line.")
    else:
        zc1.metric(f"Close / week → {target_open} by {target_deadline.strftime('%d %b %Y')}",
                   f"{math.ceil(required_per_week)} NCs",
                   help=f"({open_now_input} open - {target_open} target) ÷ {weeks_left:.0f} weeks "
                        f"+ {inflow:.1f}/wk inflow. With {team_size} people ≈ "
                        f"{max(1, round(required_per_week/team_size,1))} NC/person/week.")
    zc2.metric("Close / month", f"{math.ceil(required_per_month)} NCs",
               help="Weekly required × 4.33 weeks per month.")
    zc3.metric("Break-even (hold the line)", f"{breakeven} NCs/wk",
               help="Closing only this many just matches inflow — open count stays flat. Below this, it grows.")

    if _at_target:
        st.success(f"✅ Already at target — {open_now_input} open is at or below the target of {target_open}. "
                   f"Keep closing ~{breakeven}/week to hold it.")
    else:
        st.caption(f"Need to clear **{_gap}** NCs ({open_now_input} open → {target_open} target) "
                   f"in **{weeks_left:.0f}** weeks, while ~{inflow:.1f} new NCs arrive each week.")

    progress = closed_from_backlog / max(1, backlog_at_start)
    st.progress(min(progress, 1.0), text=f"Backlog closure: {int(closed_from_backlog)} / {int(backlog_at_start)} ({progress:.0%})")
    st.caption(f"Progress = closures from original backlog ÷ backlog at freeze. Does not include new NCs opened after {exercise_start.strftime('%d.%m.%Y')}.")

    _hist_months = pd.date_range(date_from.replace(day=1), min(date_to, today), freq="ME")
    _flt = (" AND " + " AND ".join(_BF_CL)) if _BF_CL else ""
    _rows = []
    for _m in _hist_months:
        _me = _m.strftime("%Y-%m-%d")
        _n = _q("SELECT COUNT(*) AS n FROM nc WHERE created_on <= ? AND (is_open=1 OR closure_date > ?)" + _flt,
                [_me, _me] + _BF_PR).iloc[0]["n"]
        _rows.append({"month": _m.strftime("%Y-%m"), "open": int(_n)})
    actual = pd.DataFrame(_rows)

    flow = _q(f"""
        WITH o AS (SELECT substr(created_on,1,7) AS month, COUNT(*) AS opened
                   FROM nc WHERE created_on IS NOT NULL{_flt} GROUP BY month),
             c AS (SELECT substr(closure_date,1,7) AS month, COUNT(*) AS closed
                   FROM nc WHERE closure_date IS NOT NULL{_flt} GROUP BY month)
        SELECT COALESCE(o.month,c.month) AS month,
               COALESCE(o.opened,0) AS opened, COALESCE(c.closed,0) AS closed
        FROM o LEFT JOIN c ON o.month=c.month
        WHERE COALESCE(o.month,c.month) IS NOT NULL
        ORDER BY month
    """, _BF_PR + _BF_PR)

    if actual.empty:
        st.info("No data in the selected From/To range.")
    else:
        actual.loc[actual.index[-1], "open"] = int(open_now_input)
        last_month = actual["month"].iloc[-1]
        start_open = int(open_now_input)

        _cut = (today - pd.Timedelta(weeks=12)).isoformat()
        _closed_recent = _q("SELECT COUNT(*) n FROM nc WHERE is_open=0 AND closure_date>=?" + _flt,
                            [_cut] + _BF_PR).iloc[0]["n"]
        close_rate_wk = round(_closed_recent / 12, 1)

        st.markdown("**Open NCs per month** — what actually happened")
        st.caption(f"📅 Showing **{date_from}** → **{date_to}**")
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=flow["month"], y=flow["opened"], name="Opened", marker_color="#F26E21",
                              hovertemplate="<b>%{x}</b><br>Opened: %{y}<extra></extra>"))
        fig1.add_trace(go.Bar(x=flow["month"], y=flow["closed"], name="Closed", marker_color="#4CAF50",
                              hovertemplate="<b>%{x}</b><br>Closed: %{y}<extra></extra>"))
        fig1.add_trace(go.Scatter(x=actual["month"], y=actual["open"], name="Open at month-end",
                                  mode="lines+markers", line=dict(color="#1E2761", width=3), marker=dict(size=6),
                                  hovertemplate="<b>%{x}</b><br>Open: %{y}<extra></extra>"))
        fig1.update_layout(height=300, margin=dict(l=0, r=0, t=MODEBAR_T, b=0), barmode="group",
                           legend=dict(orientation="h", y=-0.15), xaxis_title="", yaxis_title="NCs")
        st.plotly_chart(fig1, width='stretch', config=CHART_CONFIG)
        st.caption(f"Orange = NCs opened that month. Green = NCs closed that month. Navy line = total open at month-end. "
                   f"Recent pace: ~{inflow:.1f} opened/week vs ~{close_rate_wk:.1f} closed/week — "
                   f"{'closing faster than opening (backlog shrinks)' if close_rate_wk >= inflow else 'opening faster than closing (backlog grows)'}.")

        # RAW view + download: the actual NC rows in the burndown window
        _raw_dl("Excel — raw NCs in window", fname="burndown_raw.xlsx",
                key="dl_burndown_raw", include_since=True)

# restore the page overview window so the sections below are unaffected
date_from, date_to = _ov_from, _ov_to
_BF_CL, _BF_PR = _ov_BF_CL, _ov_BF_PR

st.markdown("")

# ------------------------------------------------------------------------------
# 3 - YTD
# ------------------------------------------------------------------------------
st.markdown('<div id="ytd"></div>', unsafe_allow_html=True)

_today = date.today()
_ytd_start_default = date(_today.year, 1, 1)

# Auto-roll the "To" date to today each day, but only while the person has not
# manually chosen a custom period (start still on 1 Jan AND end still on today).
_anchor = st.session_state.get("_ytd_anchor_day")
if _anchor is not None and _anchor != _today.isoformat() \
        and not st.session_state.get("_ytd_period_manual"):
    st.session_state.pop("ytd_asof", None)
    st.session_state.pop("ytd_from_pick", None)
st.session_state["_ytd_anchor_day"] = _today.isoformat()

_ac1, _ac2, _ac3, _ac4 = st.columns([1.2, 1.2, 1, 2.6])
with _ac1:
    ytd_start = st.date_input(
        "From", value=_ytd_start_default, min_value=date(_today.year - 5, 1, 1),
        max_value=date(_today.year + 1, 12, 31), key="ytd_from_pick",
        format="DD/MM/YYYY",
        help="Start of the period. Leave it on 1 January for a normal year-to-date "
             "view, or move it to report a custom range (a quarter, a single month, "
             "last week…).")
with _ac2:
    ytd_asof = st.date_input(
        "To", value=_today, min_value=date(_today.year - 5, 1, 1),
        max_value=date(_today.year + 1, 12, 31), key="ytd_asof",
        format="DD/MM/YYYY",
        help="End of the period. Defaults to today and follows the calendar "
             "automatically. Change it to report a month-end or rerun last week's "
             "numbers.")
with _ac3:
    st.markdown("<div style='height:1.8rem'></div>", unsafe_allow_html=True)
    _is_default_ytd = (ytd_start == _ytd_start_default and ytd_asof == _today)
    if st.button("This year", key="ytd_today", disabled=_is_default_ytd,
                 help="Snap back to 1 January → today."):
        st.session_state.pop("ytd_asof", None)
        st.session_state.pop("ytd_from_pick", None)
        st.session_state.pop("_ytd_period_manual", None)
        st.rerun()

# Guard: never let From be after To (swap so queries stay valid).
if ytd_start > ytd_asof:
    ytd_start, ytd_asof = ytd_asof, ytd_start

_is_ytd_view = (ytd_start == date(ytd_asof.year, 1, 1) and ytd_asof == _today)
if _is_ytd_view:
    st.session_state.pop("_ytd_period_manual", None)
else:
    st.session_state["_ytd_period_manual"] = True

_ytd_year = ytd_asof.year
_ytd_from = ytd_start.isoformat()
_ytd_to = ytd_asof.isoformat()
_ytd_from_txt = ytd_start.strftime('%d %b %Y')
_ytd_to_txt = ytd_asof.strftime('%d %b %Y')

if _is_ytd_view:
    st.subheader(f"Year to Date (YTD) {_ytd_year}")
    st.caption(f"Everything opened or closed between **1 January {_ytd_year}** and **today**. "
               f"Change the From / To dates above to report any custom period. It still "
               f"follows Project / Owner / NC type / Status / Data source.")
else:
    st.subheader(f"Selected period · {_ytd_from_txt} → {_ytd_to_txt}")
    st.caption(f"Everything opened or closed between **{_ytd_from_txt}** and **{_ytd_to_txt}**. "
               f"Press **This year** to snap back to 1 January → today. It still "
               f"follows Project / Owner / NC type / Status / Data source.")

_YF = (" AND " + " AND ".join(_NB_CL)) if _NB_CL else ""
_YTD_WIN = " AND created_on BETWEEN ? AND ?"
_YTD_CLOSE_WIN = " AND closure_date BETWEEN ? AND ?"
_YP = [_ytd_from, _ytd_to]

ytd_opened = _q("SELECT COUNT(*) n FROM nc WHERE 1=1" + _YTD_WIN + _YF,
                _YP + _NB_PR).iloc[0]["n"]
ytd_closed = _q("SELECT COUNT(*) n FROM nc WHERE 1=1" + _YTD_CLOSE_WIN + _YF,
                _YP + _NB_PR).iloc[0]["n"]
ytd_open_now = _q("SELECT COUNT(*) n FROM nc WHERE is_open=1" + _YF, _NB_PR).iloc[0]["n"]
ytd_cc = _q("SELECT COUNT(*) n FROM nc WHERE ("
            "detection_area LIKE 'CC%' OR LOWER(detection_area) LIKE '%customer%complaint%')"
            + _YTD_WIN + _YF, _YP + _NB_PR).iloc[0]["n"]

_y1 = [
    (f"NCs opened in {_ytd_year}", int(ytd_opened),
     f"Created between {_ytd_from_txt} and {_ytd_to_txt}."),
    (f"NCs closed in {_ytd_year}", int(ytd_closed),
     "Closed inside that window, including backlog from earlier years."),
    ("Open now", int(ytd_open_now),
     "Currently unresolved, any year. This one is live status, not a snapshot — "
     "it does not move with the As of date."),
    ("Customer complaints", int(ytd_cc),
     "NCs in the window whose detection area is a customer complaint (CC)."),
]
for _c, (_l, _v, _d) in zip(st.columns(4), _y1):
    _c.markdown(_kpi_card(_l, _v, _d), unsafe_allow_html=True)

# RAW view + download: the actual NC rows opened in the YTD window
_ytd_raw = _q("SELECT * FROM nc WHERE 1=1" + _YTD_WIN + _YF + " ORDER BY created_on DESC",
              _YP + _NB_PR)
raw_block(_ytd_raw, key="ytd_raw", fname="ytd_raw_ncs.xlsx", sheet="YTD_raw")

yc1, yc2 = st.columns(2)

with yc1:
    st.markdown(f"**Supplier vs Internal** · :grey[NCs opened in {_ytd_year} to {ytd_asof.strftime('%d %b')}]")
    df_ytd_si = _q("""
        SELECT CASE WHEN is_supplier_nc=1 THEN 'Supplier' ELSE 'Internal' END AS kind,
               COUNT(*) AS n
        FROM nc WHERE 1=1""" + _YTD_WIN + _YF + " GROUP BY kind", _YP + _NB_PR)
    if not df_ytd_si.empty:
        fig = px.bar(df_ytd_si, x="kind", y="n", text="n",
                     color="kind", color_discrete_map={"Internal": "#1E2761",
                                                       "Supplier": "#F26E21"})
        fig.update_layout(height=250, margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                          showlegend=False, xaxis_title="", yaxis_title="NCs")
        fig.update_traces(textposition="outside",
                          hovertemplate="<b>%{x}</b><br>NCs: %{y}<extra></extra>")
        _headroom(fig, df_ytd_si["n"])
        st.plotly_chart(fig, width='stretch', config=CHART_CONFIG, key="ytd_si")
        st.caption("In SAP, Z2 = procurement complaint (supplier) and Z3 = production NC "
                   "(internal). NCs from the new system take the NC Type recorded in the "
                   "NC tracker instead.")
    else:
        st.info(f"No {_ytd_year} NCs in this selection.")

with yc2:
    st.markdown(f"**Major vs Minor** · :grey[NCs opened in {_ytd_year} to {ytd_asof.strftime('%d %b')}]")
    df_ytd_mm = _q("""
        SELECT CASE WHEN classification LIKE 'Major%' THEN 'Major'
                    WHEN classification LIKE 'Minor%' THEN 'Minor'
                    WHEN classification IS NULL THEN '(not classified)'
                    ELSE classification END AS kind,
               COUNT(*) AS n
        FROM nc WHERE 1=1""" + _YTD_WIN + _YF +
        " GROUP BY kind ORDER BY n DESC", _YP + _NB_PR)
    if not df_ytd_mm.empty:
        _mm_colors = ["#E53E3E" if k == "Major" else
                      ("#B0B7C3" if k.startswith("(") else "#4CAF50")
                      for k in df_ytd_mm["kind"]]
        fig = px.bar(df_ytd_mm, x="kind", y="n", text="n", color="kind",
                     color_discrete_sequence=_mm_colors)
        fig.update_layout(height=250, margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                          showlegend=False, xaxis_title="", yaxis_title="NCs")
        fig.update_traces(textposition="outside",
                          hovertemplate="<b>%{x}</b><br>NCs: %{y}<extra></extra>")
        _headroom(fig, df_ytd_mm["n"])
        st.plotly_chart(fig, width='stretch', config=CHART_CONFIG, key="ytd_mm")
        st.caption("Major NCs need an NRB disposition and carry the higher risk. "
                   "'(not classified)' means the classification cell is empty in the source.")
    else:
        st.info(f"No {_ytd_year} NCs in this selection.")

st.markdown(f"**Top recurring issues** · :grey[NCs opened in {_ytd_year} to {ytd_asof.strftime('%d %b')}]")
# Pick the issue field that actually carries data. defect_code_text lived only in
# the SAP export (now dropped) — if it is empty, use the tracker's Failure field,
# which is the equivalent "what kind of issue"; detection_area is the last resort.
def _has_data(col):
    return _has_column(col) and _q(
        f"SELECT COUNT(*) AS n FROM nc WHERE {col} IS NOT NULL AND TRIM({col})<>''"
    ).iloc[0]["n"] > 0
_issue_col = ("defect_code_text" if _has_data("defect_code_text")
              else "failure" if _has_data("failure")
              else "detection_area")
df_ytd_top = _q(f"""
    SELECT COALESCE({_issue_col},'(not recorded)') AS issue, COUNT(*) AS n
    FROM nc WHERE 1=1{_YTD_WIN}{_YF}
    GROUP BY issue ORDER BY n DESC LIMIT 10""", _YP + _NB_PR)
if not df_ytd_top.empty:
    fig = px.bar(df_ytd_top, x="n", y="issue", orientation="h",
                 color_discrete_sequence=["#1E2761"], text="n")
    fig.update_layout(height=max(260, 34 * len(df_ytd_top)),
                      margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                      yaxis=dict(categoryorder="total ascending"),
                      xaxis_title="NCs", yaxis_title="")
    fig.update_traces(textposition="outside",
                      hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
    _headroom(fig, df_ytd_top["n"], horizontal=True)
    st.plotly_chart(fig, width='stretch', config=CHART_CONFIG, key="ytd_top")
    st.caption(f"The ten most frequent values of `{_issue_col}` among {_ytd_year} NCs. "
               "A tall bar repeating year on year is the case for a preventive action.")
else:
    st.info(f"No {_ytd_year} NCs in this selection.")

st.markdown("")


# ------------------------------------------------------------------------------
# 4 - MONTHLY PERFORMANCE
# ------------------------------------------------------------------------------
st.markdown('<div id="monthly"></div>', unsafe_allow_html=True)

_cm_cl, _cm_pr = _bd_filter(include_dates=True)
_cm_flt = (" AND " + " AND ".join(_cm_cl)) if _cm_cl else ""

st.markdown('<div class="month-anchor"></div>', unsafe_allow_html=True)
with st.container(border=True):
    _months_df = _q(
        "SELECT DISTINCT substr(created_on,1,7) AS m FROM nc WHERE created_on IS NOT NULL"
        + _cm_flt + " ORDER BY m DESC", _cm_pr)
    _month_opts = [m for m in _months_df["m"].tolist() if m]
    _this_month = datetime.now().strftime("%Y-%m")
    if _this_month not in _month_opts:
        _month_opts = [_this_month] + _month_opts
    _default_idx = _month_opts.index(_this_month) if _this_month in _month_opts else 0

    hcol, scol = st.columns([2, 1])
    with scol:
        cm = st.selectbox("Month", _month_opts, index=_default_idx, key="month_pick",
                          help="Pick any month inside the From/To window. Defaults to the current month.")
    cm_label = datetime.strptime(cm, "%Y-%m").strftime("%B %Y")
    with hcol:
        _is_current = (cm == _this_month)
        st.subheader(("Monthly Performance — " if _is_current else "Month — ") + cm_label)

    cm_opened = _q("SELECT COUNT(*) AS n FROM nc WHERE substr(created_on,1,7)=?" + _cm_flt,
                   [cm] + _cm_pr).iloc[0]["n"]
    cm_closed = _q("SELECT COUNT(*) AS n FROM nc WHERE substr(closure_date,1,7)=?" + _cm_flt,
                   [cm] + _cm_pr).iloc[0]["n"]
    cm_wip = _q("SELECT COUNT(*) AS n FROM nc WHERE is_open=1" + (
        (" AND " + " AND ".join(_NB_CL)) if _NB_CL else ""), _NB_PR).iloc[0]["n"]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Opened", int(cm_opened),
              help=f"NCs created in {cm_label}. New quality issues entering the queue.")
    m2.metric("Closed", int(cm_closed),
              help=f"NCs closed in {cm_label}. The team's closure output.")
    m3.metric("WIP", int(cm_wip),
              help="Every NC still open right now — this year's plus the older backlog. "
                   "Not limited to the selected month.")
    m4.metric("Net (closed − opened)", int(cm_closed - cm_opened),
              delta_color="normal" if cm_closed >= cm_opened else "inverse",
              help="Closed minus opened. Positive = backlog shrinking. Negative = backlog growing.")

    # RAW view + download: NC rows opened in the chosen month
    _cm_raw = _q("SELECT * FROM nc WHERE substr(created_on,1,7)=?" + _cm_flt
                 + " ORDER BY created_on DESC", [cm] + _cm_pr)
    raw_block(_cm_raw, key="month_raw", fname=f"ncs_{cm}.xlsx", sheet="Month_raw",
              dl_label=f"Download {cm_label}")

st.markdown("")


# ------------------------------------------------------------------------------
# 5 - OPEN NCs BY AREA
# ------------------------------------------------------------------------------
st.markdown('<div id="open-by-area"></div>', unsafe_allow_html=True)

# ============================================================================
# ON-PAGE FILTER — sits here, above "Open NCs by Area", and governs ONLY the
# graphs BELOW it (sections 5–10). Sections 1–4 above are untouched.
# Compact & collapsible: one line by default, the header shows the active
# window/filters, click to open the controls.
# ============================================================================
_sf_from = st.session_state["date_from"]
_sf_to = st.session_state["date_to"]
_sum_bits = [f"{_sf_from.strftime('%d %b %Y')} → {_sf_to.strftime('%d %b %Y')}"]
if st.session_state["nc_type"] != "All": _sum_bits.append(st.session_state["nc_type"])
if st.session_state["nc_status"] != "All": _sum_bits.append(st.session_state["nc_status"])
if st.session_state["nc_project"]: _sum_bits.append(f"{len(st.session_state['nc_project'])} project(s)")
if st.session_state["nc_owner"]: _sum_bits.append(f"{len(st.session_state['nc_owner'])} owner(s)")
if st.session_state["nc_source"] != "All": _sum_bits.append(st.session_state["nc_source"])
_flabel = "🔍  Filter (this section and everything below)  ·  " + "   ·   ".join(_sum_bits)

with st.expander(_flabel, expanded=False):
    st.caption("Quick range")
    _pcol1, _pcol2, _pcol3 = st.columns(3)
    if _pcol1.button("This year", width='stretch', key="preset_year",
                     help="1 Jan of the current year → today."):
        st.session_state["date_from"] = max(_pick_min, date(date.today().year, 1, 1))
        st.session_state["date_to"] = date.today()
        for _wk in ("w_date_from", "w_date_to"):
            st.session_state.pop(_wk, None)
        st.session_state.pop("_to_manual", None)
        st.rerun()
    if _pcol2.button("Last 12 mo", width='stretch', key="preset_12m",
                     help="Rolling 12 months ending today."):
        st.session_state["date_from"] = max(_pick_min,
                                            (pd.Timestamp(date.today()) - pd.DateOffset(months=12)).date())
        st.session_state["date_to"] = date.today()
        for _wk in ("w_date_from", "w_date_to"):
            st.session_state.pop(_wk, None)
        st.session_state.pop("_to_manual", None)
        st.rerun()
    if _pcol3.button("All", width='stretch', key="preset_all",
                     help="Everything from the earliest NC to today."):
        st.session_state["date_from"] = _pick_min
        st.session_state["date_to"] = date.today()
        for _wk in ("w_date_from", "w_date_to"):
            st.session_state.pop(_wk, None)
        st.session_state.pop("_to_manual", None)
        st.rerun()

    with st.form("general_filters", clear_on_submit=False):
        r1c1, r1c2, r1c3, r1c4 = st.columns(4)
        with r1c1:
            f_date_from = st.date_input("From", value=st.session_state["date_from"],
                                        min_value=_pick_min, max_value=_pick_max,
                                        key="w_date_from", format="DD/MM/YYYY")
        with r1c2:
            f_date_to = st.date_input("To", value=st.session_state["date_to"],
                                      min_value=_pick_min, max_value=_pick_max,
                                      key="w_date_to", format="DD/MM/YYYY")
        with r1c3:
            f_nc_type = st.selectbox("NC type", ["All", "Production", "Supplier"],
                                     index=["All", "Production", "Supplier"].index(st.session_state["nc_type"]),
                                     key="w_nc_type")
        with r1c4:
            _status_opts = ["All", "Open", "Closed"] + (["(no status)"] if HAS_STATUS_STATE else [])
            f_nc_status = st.selectbox(
                "Status", _status_opts,
                index=_status_opts.index(st.session_state["nc_status"])
                if st.session_state["nc_status"] in _status_opts else 0,
                key="w_nc_status",
                help="'(no status)' = the Status cell is blank in the source. These NCs are "
                     "neither open nor closed.")

        r2c1, r2c2, r2c3 = st.columns(3)
        with r2c1:
            f_nc_project = st.multiselect("Project", _projects,
                                          default=[p for p in st.session_state["nc_project"] if p in _projects],
                                          key="w_nc_project")
        with r2c2:
            f_nc_owner = st.multiselect("Owner", _owners,
                                        default=[o for o in st.session_state["nc_owner"] if o in _owners],
                                        key="w_nc_owner")
        with r2c3:
            if _sources:
                f_nc_source = st.selectbox("Data source", ["All"] + _sources,
                                           index=(["All"] + _sources).index(st.session_state["nc_source"])
                                           if st.session_state["nc_source"] in (["All"] + _sources) else 0,
                                           key="w_nc_source")
            else:
                f_nc_source = "All"

        c_apply, c_reset, _cgap = st.columns([1, 1, 4])
        applied = c_apply.form_submit_button("✓ Apply filters", width='stretch', type="primary")
        reset = c_reset.form_submit_button("Reset", width='stretch')

    if applied:
        f_date_from = st.session_state["w_date_from"]
        f_date_to = st.session_state["w_date_to"]
        f_nc_type = st.session_state["w_nc_type"]
        f_nc_status = st.session_state["w_nc_status"]
        f_nc_project = st.session_state["w_nc_project"]
        f_nc_owner = st.session_state["w_nc_owner"]
        f_nc_source = st.session_state.get("w_nc_source", "All")
        if f_date_from > f_date_to:
            st.error(f"⚠️ 'From' ({f_date_from}) is after 'To' ({f_date_to}). "
                     "Pick a From date on or before the To date — filters not applied.")
        else:
            st.session_state["date_from"] = f_date_from
            st.session_state["date_to"] = f_date_to
            st.session_state["_to_manual"] = (f_date_to != date.today())
            st.session_state["nc_type"] = f_nc_type
            st.session_state["nc_status"] = f_nc_status
            st.session_state["nc_project"] = f_nc_project
            st.session_state["nc_owner"] = f_nc_owner
            st.session_state["nc_source"] = f_nc_source
            st.rerun()

    if reset:
        for _k, _v in _filter_defaults.items():
            st.session_state[_k] = _v
        for _wk in ["w_date_from", "w_date_to", "w_nc_type", "w_nc_status",
                    "w_nc_project", "w_nc_owner", "w_nc_source"]:
            st.session_state.pop(_wk, None)
        st.session_state.pop("_to_manual", None)
        st.rerun()

# Re-point the globals to the FILTER's committed values so every section from
# here down (5–10) runs on them. build_where / _qf / raw_nc read these at call
# time, so recomputing the shared clause lists switches the lower sections onto
# the filter's window in one place. (_ND_* stays date-free for the two
# management trend numbers, which must span their own fixed period.)
date_from = st.session_state["date_from"]
date_to = st.session_state["date_to"]
if date_from > date_to:
    date_from, date_to = date_to, date_from
nc_type = st.session_state["nc_type"]
nc_status = st.session_state["nc_status"]
nc_project = st.session_state["nc_project"]
nc_owner = st.session_state["nc_owner"]
nc_source = st.session_state["nc_source"]
_BF_CL, _BF_PR = _bd_filter(include_dates=True)
_NB_CL, _NB_PR = _bd_filter(include_dates=True)
_ND_CL, _ND_PR = _bd_filter(include_dates=False)

st.subheader("Open NCs by Area")
st.caption("Open NCs only, grouped by the detection area code recorded in the NC tracker "
           "(B: Bonding, I: Integration, CC: Customer Complaint, and so on). Closed NCs "
           "are excluded, and only NCs created inside the date window above are counted.")

df_area = _q("""
    SELECT COALESCE(detection_area,'BLANK - to clean') AS area, COUNT(*) AS open_ncs
    FROM nc WHERE is_open=1""" + ((" AND " + " AND ".join(_NB_CL)) if _NB_CL else "") + """
    GROUP BY area ORDER BY open_ncs DESC""", _NB_PR)
if not df_area.empty:
    _colors = ["#E53E3E" if a.startswith("BLANK") else "#F26E21" for a in df_area["area"]]
    fig = px.bar(df_area, x="open_ncs", y="area", orientation="h",
                 color=df_area["area"], color_discrete_sequence=_colors, text="open_ncs")
    fig.update_layout(height=max(300, 30 * len(df_area)),
                      margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                      yaxis=dict(categoryorder="total ascending"),
                      xaxis_title="Open NCs", yaxis_title="")
    fig.update_traces(textposition="outside",
                      hovertemplate="<b>%{y}</b><br>Open NCs: %{x}<extra></extra>")
    _headroom(fig, df_area["open_ncs"], horizontal=True)
    st.plotly_chart(fig, width='stretch', config=CHART_CONFIG, key="open_by_area")
    _blank = int(df_area.loc[df_area["area"].str.startswith("BLANK"), "open_ncs"].sum())
    st.caption(f"Red = no detection area recorded ({_blank} open NCs). Detection area is "
               "captured in the NC tracker only, so an SAP-only NC has no area by design.")
    # RAW download: the actual open NC rows (created inside the date window)
    _raw_dl("Excel — raw open NCs", extra=[("is_open=1", [])], use_date=True,
            fname="open_ncs_by_area_raw.xlsx", key="dl_open_area")
else:
    st.info("No open NCs in this selection.")

st.markdown("")


# ------------------------------------------------------------------------------
# 6 - WHERE NCs HAPPEN — by project & class
# ------------------------------------------------------------------------------
st.markdown('<div id="by-project"></div>', unsafe_allow_html=True)
st.subheader("Where NCs happen — by project & class")
st.caption("Follows the date filter above (plus type, status…). Total = all NCs created in "
           "the window; Open = still unresolved. Shows which programmes carry the most NCs "
           "and where the live workload sits now. Every download below is the RAW NC rows.")

# Project-breakdown charts use ONLY NCs with a real project (Adriele): a NO
# PROJECT / EZ1 NC is never counted under a real programme. Open-now metrics
# above still count everything; this filter is only for the by-project views.
_PROJ_ONLY = [("COALESCE(project,'NO PROJECT') <> 'NO PROJECT'", [])]

df_proj = _qf("""SELECT COALESCE(project,'(no project)') AS project,
                        COUNT(*) AS total,
                        SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS open
                 FROM nc {WHERE} GROUP BY project ORDER BY total DESC""",
              extra=_PROJ_ONLY)
if df_proj.empty:
    st.info("No NCs in this selection.")
else:
    pc1, pc2 = st.columns(2)
    with pc1:
        st.markdown("**All NCs by project**")
        figpt = go.Figure(go.Bar(
            y=df_proj["project"], x=df_proj["total"], orientation="h",
            marker_color="#1E2761", text=df_proj["total"],
            hovertemplate="<b>%{y}</b><br>Total NCs: %{x}<extra></extra>"))
        figpt.update_layout(height=max(300, 32*len(df_proj)),
                            margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                            yaxis=dict(categoryorder="total ascending"),
                            xaxis_title="NCs (total)", yaxis_title="")
        figpt.update_traces(textposition="outside")
        _headroom(figpt, df_proj["total"], horizontal=True)
        st.plotly_chart(figpt, width='stretch', config=CHART_CONFIG, key="by_project_total")
        # raw data right under THIS chart
        raw_block(raw_nc(_PROJ_ONLY), key="proj_total_raw", fname="all_ncs_by_project.xlsx",
                  dl_label="All NCs by project",
                  note="Every NC in the current filter (real projects only).")
    with pc2:
        st.markdown("**Open NCs by project** — current workload")
        dpo = df_proj[df_proj["open"] > 0].sort_values("open")
        if dpo.empty:
            st.info("No open NCs in this selection.")
        else:
            figpo = go.Figure(go.Bar(
                y=dpo["project"], x=dpo["open"], orientation="h",
                marker_color="#F26E21", text=dpo["open"],
                hovertemplate="<b>%{y}</b><br>Open NCs: %{x}<extra></extra>"))
            figpo.update_layout(height=max(300, 32*len(dpo)),
                                margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                                yaxis=dict(categoryorder="total ascending"),
                                xaxis_title="Open NCs", yaxis_title="")
            figpo.update_traces(textposition="outside")
            _headroom(figpo, dpo["open"], horizontal=True)
            st.plotly_chart(figpo, width='stretch', config=CHART_CONFIG, key="by_project_open")
        # raw data right under THIS chart
        raw_block(raw_nc(_PROJ_ONLY + [("is_open=1", [])]), key="proj_open_raw",
                  fname="open_ncs_by_project.xlsx", dl_label="Open NCs by project",
                  note="Open NCs only (real projects), in the current filter.")

st.markdown("**Class → project** — biggest to smallest")
df_tm = _qf("""SELECT COALESCE(launcher_class,'(no class)') AS cls,
                      COALESCE(project,'(no project)') AS project, COUNT(*) AS n
               FROM nc {WHERE} GROUP BY cls, project""", extra=_PROJ_ONLY)
if not df_tm.empty:
    figtm = px.treemap(df_tm, path=[px.Constant("All NCs"), "cls", "project"], values="n",
                       color="cls",
                       color_discrete_map={"LLV": "#1E2761", "MLV": "#F26E21",
                                           "SLV": "#4CAF50", "SAS": "#1C7293",
                                           "(no class)": "#B0B7C3"})
    figtm.update_traces(root_color="#EEF1F7",
                        texttemplate="%{label}<br><b>%{value}</b> NCs",
                        textfont_size=14,
                        hovertemplate="<b>%{label}</b><br>NCs: %{value}<extra></extra>")
    figtm.update_layout(height=430, margin=dict(l=0, r=0, t=MODEBAR_T, b=0))
    _tm_ev = st.plotly_chart(figtm, width='stretch', config=CHART_CONFIG,
                             key="by_project_treemap", on_select="rerun")
    st.caption("Top blocks are launcher classes (LLV includes Vulcan; SAS stands alone) — "
               "click a class or project; the raw data below follows your click.")

    # Which node was clicked? Treemap point ids look like 'All NCs/LLV/Ariane'.
    # depth 1 = a class, depth 2 = a project. Use it to scope the raw rows so
    # 'View raw data' shows exactly what the selected block represents.
    _sel_cls = _sel_proj = None
    try:
        _pts = _tm_ev["selection"]["points"]
    except Exception:
        _pts = []
    if _pts:
        _parts = [p for p in str(_pts[-1].get("id", "")).split("/")
                  if p and p != "All NCs"]
        if len(_parts) == 1:
            _sel_cls = _parts[0]
        elif len(_parts) >= 2:
            _sel_cls, _sel_proj = _parts[0], _parts[1]

    _tm_extra, _tm_note = [], ("Every NC in the current filter "
                               "(each row carries its class & project).")
    if _sel_proj:
        _tm_extra = [("COALESCE(project,'(no project)') = ?", [_sel_proj])]
        _tm_note = f"Filtered to your click: project **{_sel_proj}**."
    elif _sel_cls:
        _tm_extra = [("COALESCE(launcher_class,'(no class)') = ?", [_sel_cls])]
        _tm_note = f"Filtered to your click: class **{_sel_cls}**."
    raw_block(raw_nc(_PROJ_ONLY + _tm_extra), key="treemap_raw",
              fname="class_project_raw.xlsx",
              dl_label="Class → project rows", note=_tm_note)

st.markdown("**Where + what** — NCs by project × detection area")
df_hm = _qf("""SELECT COALESCE(project,'(no project)') AS project,
                      COALESCE(detection_area,'(none)') AS area, COUNT(*) AS n
               FROM nc {WHERE} GROUP BY project, area""", extra=_PROJ_ONLY)
if not df_hm.empty:
    _topP = df_hm.groupby("project")["n"].sum().sort_values(ascending=False).head(10).index.tolist()
    _topA = df_hm.groupby("area")["n"].sum().sort_values(ascending=False).head(12).index.tolist()
    _h = df_hm[df_hm["project"].isin(_topP) & df_hm["area"].isin(_topA)]
    piv = _h.pivot_table(index="area", columns="project", values="n", aggfunc="sum", fill_value=0)
    if not piv.empty:
        fighm = go.Figure(go.Heatmap(
            z=piv.values, x=list(piv.columns), y=list(piv.index), colorscale="Oranges",
            hovertemplate="Project %{x}<br>Area %{y}<br>NCs: %{z}<extra></extra>"))
        fighm.update_layout(height=max(320, 26*len(piv.index)),
                            margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                            xaxis_title="", yaxis_title="")
        st.plotly_chart(fighm, width='stretch', config=CHART_CONFIG, key="by_project_heatmap")
        st.caption("Darker = more NCs. Top 10 projects × top 12 detection areas — the hotspots "
                   "to attack first (a dark cell = that problem in that programme).")
        # raw scoped to exactly what the heatmap draws (top-10 projects × top-12 areas)
        _hm_extra = []
        if _topP:
            _hm_extra.append((f"COALESCE(project,'(no project)') IN ({','.join(['?']*len(_topP))})", _topP))
        if _topA:
            _hm_extra.append((f"COALESCE(detection_area,'(none)') IN ({','.join(['?']*len(_topA))})", _topA))
        raw_block(raw_nc(_PROJ_ONLY + _hm_extra), key="heatmap_raw", fname="project_area_raw.xlsx",
                  dl_label="Project × area rows",
                  note="The NCs shown in the heatmap (top-10 projects × top-12 detection areas).")

st.markdown("")


# ------------------------------------------------------------------------------
# 7 - ROOT CAUSE ANALYSIS (headline only; full analysis on the CAPA tab)
# ------------------------------------------------------------------------------
st.markdown('<div id="root-cause"></div>', unsafe_allow_html=True)
st.subheader("Root Cause Analysis")

df_rc_top = None
df_oa_top = None
if HAS_CAPA:
    _rca_only = " AND c.capa_type='RCA'" if HAS_CAPA_TYPE else ""
    _rc_flt = (" AND " + " AND ".join(_NB_CL)) if _NB_CL else ""
    df_oa_top = _q(f"""SELECT c.origin_area_l1 AS area, COUNT(*) AS n
                       FROM capa c JOIN nc USING(nc_id)
                       WHERE c.origin_area_l1 IS NOT NULL{_rca_only}{_rc_flt}
                       GROUP BY area ORDER BY n DESC LIMIT 8""", _NB_PR)
    df_rc_top = _q(f"""SELECT c.rc_category_l1 AS cat, COUNT(*) AS n
                       FROM capa c JOIN nc USING(nc_id)
                       WHERE c.rc_category_l1 IS NOT NULL{_rca_only}{_rc_flt}
                       GROUP BY cat ORDER BY n DESC LIMIT 8""", _NB_PR)

    rcc1, rcc2 = st.columns(2)
    with rcc1:
        st.markdown("**Top Real Origin Area (L1)**")
        if df_oa_top is not None and not df_oa_top.empty:
            fig = px.bar(df_oa_top, x="n", y="area", orientation="h",
                         color_discrete_sequence=["#1C7293"], text="n")
            fig.update_layout(height=280, margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                              showlegend=False, yaxis=dict(categoryorder="total ascending"),
                              xaxis_title="", yaxis_title="")
            fig.update_traces(textposition="outside",
                              hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
            _headroom(fig, df_oa_top["n"], horizontal=True)
            st.plotly_chart(fig, width='stretch', config=CHART_CONFIG, key="rc_area")
        else:
            st.info("No origin area recorded for this selection.")
    with rcc2:
        st.markdown("**Top RC Category (L1)**")
        if df_rc_top is not None and not df_rc_top.empty:
            fig = px.bar(df_rc_top, x="n", y="cat", orientation="h",
                         color_discrete_sequence=["#21295C"], text="n")
            fig.update_layout(height=280, margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                              showlegend=False, yaxis=dict(categoryorder="total ascending"),
                              xaxis_title="", yaxis_title="")
            fig.update_traces(textposition="outside",
                              hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
            _headroom(fig, df_rc_top["n"], horizontal=True)
            st.plotly_chart(fig, width='stretch', config=CHART_CONFIG, key="rc_cat")
        else:
            st.info("No RC category recorded for this selection.")

    # RAW view + download: the RCA rows behind these charts (raw CAPA joined to its NC)
    _rc_raw = _q(f"""SELECT c.*, n.project, n.system, n.owner, n.created_on AS nc_created_on
                     FROM capa c JOIN nc n USING(nc_id)
                     WHERE 1=1{_rca_only}{_rc_flt}""", _NB_PR)
    raw_block(_rc_raw, key="rc_raw", fname="root_cause_raw.xlsx", sheet="RCA_raw",
              dl_label="Download RCA rows")

    st.caption("Where the problem started and what kind of cause it was, taken from the RCA "
               "rows of the CAPA tracker. The full analysis — drill-downs by launcher class, "
               "CAPA coverage and the L2 breakdowns — is on the **CAPA** tab.")
else:
    st.info("No CAPA data in `quality.db`. Run the ingest with the CAPA tracker present.")

st.markdown("")


# ------------------------------------------------------------------------------
# 8 - MANAGEMENT VIEW
# ------------------------------------------------------------------------------
st.markdown('<div id="mgmt-view"></div>', unsafe_allow_html=True)
st.subheader("Management View")
st.caption("The five management questions, answered from the data above. Every answer is "
           "computed — nothing here is typed by hand. The two trend questions (backlog "
           "reducing? / closures vs openings?) use the full history so a narrow date "
           "window can't distort them; the rest follow the date filter.")

# _MF  = full filter INCLUDING the inline date window (for "open now" snapshots).
# _MF_ND = same filter but WITHOUT the date window — used by the two trend
#          questions (6-month backlog change, opened-vs-closed this year) so a
#          narrow window can't distort a measurement that is about a fixed span.
_MF = (" AND " + " AND ".join(_NB_CL)) if _NB_CL else ""
_MF_ND = (" AND " + " AND ".join(_ND_CL)) if _ND_CL else ""

_now_m = date.today().strftime("%Y-%m-%d")
_six_ago = (pd.Timestamp(date.today()) - pd.DateOffset(months=6)).strftime("%Y-%m-%d")
_open_now = _q("SELECT COUNT(*) n FROM nc WHERE created_on <= ? AND "
               "(is_open=1 OR closure_date > ?)" + _MF_ND, [_now_m, _now_m] + _ND_PR).iloc[0]["n"]
_open_then = _q("SELECT COUNT(*) n FROM nc WHERE created_on <= ? AND "
                "(is_open=1 OR closure_date > ?)" + _MF_ND, [_six_ago, _six_ago] + _ND_PR).iloc[0]["n"]
_delta = int(_open_now) - int(_open_then)

_yr = date.today().strftime("%Y")
_op_yr = _q("SELECT COUNT(*) n FROM nc WHERE substr(created_on,1,4)=?" + _MF_ND,
            [_yr] + _ND_PR).iloc[0]["n"]
_cl_yr = _q("SELECT COUNT(*) n FROM nc WHERE substr(closure_date,1,4)=?" + _MF_ND,
            [_yr] + _ND_PR).iloc[0]["n"]

_top_area = _q("SELECT COALESCE(detection_area,'(not recorded)') a, COUNT(*) n FROM nc "
               "WHERE is_open=1" + _MF + " GROUP BY a ORDER BY n DESC LIMIT 1", _NB_PR)
_major_open = _q("SELECT COUNT(*) n FROM nc WHERE is_open=1 AND classification LIKE 'Major%'"
                 + _MF, _NB_PR).iloc[0]["n"]
_no_area = _q("SELECT COUNT(*) n FROM nc WHERE is_open=1 AND detection_area IS NULL" + _MF,
              _NB_PR).iloc[0]["n"]

_top_rc = "not recorded"
if df_rc_top is not None and not df_rc_top.empty:
    _top_rc = ", ".join(df_rc_top["cat"].head(3).tolist())

_rows = [
    ("Is the backlog reducing?",
     "Yes" if _delta < 0 else ("Flat" if _delta == 0 else "No"),
     f"Open NCs went from {int(_open_then)} six months ago to {int(_open_now)} today "
     f"({_delta:+d})."),
    ("Are closures exceeding openings?",
     "Yes" if _cl_yr > _op_yr else ("Level" if _cl_yr == _op_yr else "No"),
     f"In {_yr}: {int(_op_yr)} opened vs {int(_cl_yr)} closed."),
    ("Where is the backlog accumulating?",
     _top_area.iloc[0]["a"] if not _top_area.empty else "—",
     f"{int(_top_area.iloc[0]['n'])} open NCs in that area — the largest single pile."
     if not _top_area.empty else "No open NCs in this selection."),
    ("What are the main root causes?",
     _top_rc.split(",")[0].strip() if _top_rc != "not recorded" else "—",
     f"Top three RC categories from the CAPA tracker: {_top_rc}."),
    ("Which areas need management attention?",
     f"{int(_major_open)} Major open",
     f"{int(_major_open)} open NCs are classified Major. {int(_no_area)} open NCs have no "
     f"detection area recorded and cannot be routed."),
]

df_mgmt = pd.DataFrame(_rows, columns=["Question", "Answer", "Evidence"])
st.dataframe(df_mgmt, width='stretch', hide_index=True,
             column_config={"Question": st.column_config.TextColumn(width="medium"),
                            "Answer": st.column_config.TextColumn(width="small"),
                            "Evidence": st.column_config.TextColumn(width="large")})
# RAW download: the open NCs the management view is talking about (in the window)
_raw_dl("Excel — raw open NCs", extra=[("is_open=1", [])], use_date=True,
        fname="management_view_raw_open.xlsx", key="dl_mgmt")

st.caption(f"Data source: `quality.db` · Last modified: "
           f"{datetime.fromtimestamp(Path(DB_FILE).stat().st_mtime).strftime('%d.%m.%Y %H:%M')}")


# ------------------------------------------------------------------------------
# 9 - OWNER SUMMARY (NC tracker only)
# ------------------------------------------------------------------------------
st.markdown('<div id="owner-summary"></div>', unsafe_allow_html=True)
st.subheader("Owner summary")
st.caption("**NC tracker rows only** — the cutover population, not the full SAP history. "
           "SAP-only NCs carry no owner (the export has no owner field), so including them "
           "would put thousands of NCs against '(no owner)' and drown the real workload. "
           "Follows the date filter above (NCs created inside the window).")

_OWN_FLT = (" AND " + " AND ".join(_NB_CL)) if _NB_CL else ""

df_owner_sum = _q("""
    SELECT COALESCE(owner,'(no owner)') AS "Owner",
           SUM(CASE WHEN system='SAP'      AND is_open=1 THEN 1 ELSE 0 END) AS "SAP open",
           SUM(CASE WHEN system='EZ1'      AND is_open=1 THEN 1 ELSE 0 END) AS "EZYone open",
           SUM(CASE WHEN system='Blackout' AND is_open=1 THEN 1 ELSE 0 END) AS "Blackout open",
           SUM(CASE WHEN is_open=1  THEN 1 ELSE 0 END) AS "Open",
           SUM(CASE WHEN is_open=0  THEN 1 ELSE 0 END) AS "Closed",
           SUM(CASE WHEN is_open IS NULL THEN 1 ELSE 0 END) AS "No status",
           COUNT(*) AS "Total"
    FROM nc
    WHERE source IN ('tracker','both')""" + _OWN_FLT + """
    GROUP BY 1 ORDER BY "Total" DESC, "Owner\"""", _NB_PR)

if df_owner_sum.empty:
    st.info("No tracker NCs in this selection.")
else:
    _tot = df_owner_sum.drop(columns=["Owner"]).sum()
    _totals = pd.DataFrame([{"Owner": "TOTAL", **_tot.to_dict()}])
    df_owner_show = pd.concat([df_owner_sum, _totals], ignore_index=True)

    st.dataframe(df_owner_show, width='stretch', hide_index=True,
                 height=min(560, 38 * (len(df_owner_show) + 1)),
                 column_config={
                     "Owner": st.column_config.TextColumn(width="medium"),
                     "SAP open": st.column_config.NumberColumn(
                         help="Open NCs whose System column says SAP."),
                     "EZYone open": st.column_config.NumberColumn(
                         help="Open NCs created in EZYone since go-live."),
                     "Blackout open": st.column_config.NumberColumn(
                         help="Open NCs raised during the blackout, not yet migrated."),
                     "No status": st.column_config.NumberColumn(
                         help="Status cell blank in the tracker — neither open nor closed."),
                 })

    _n_closed = int(df_owner_sum["Closed"].sum())
    _n_nostat = int(df_owner_sum["No status"].sum())
    st.caption(
        f"{int(df_owner_sum['Total'].sum())} tracker NCs across {len(df_owner_sum)} owners. "
        f"**Closed reads {_n_closed}** — almost nobody marks an NC Closed in the tracker, so "
        f"this column is close to empty by data, not by filter. **{_n_nostat}** rows have a "
        f"blank status: neither open nor closed, and easy to miss.")

    _open_by_sys = df_owner_sum[["Owner", "SAP open", "EZYone open", "Blackout open"]]
    _melt = _open_by_sys.melt(id_vars="Owner", var_name="System", value_name="Open")
    _melt = _melt[_melt["Open"] > 0]
    if not _melt.empty:
        figo = px.bar(_melt, x="Open", y="Owner", color="System", orientation="h",
                      color_discrete_map={"SAP open": "#1E2761",
                                          "EZYone open": "#F26E21",
                                          "Blackout open": "#4CAF50"})
        figo.update_layout(height=max(300, 30 * df_owner_sum["Owner"].nunique()),
                           margin=dict(l=0, r=0, t=MODEBAR_T, b=0), barmode="stack",
                           yaxis=dict(categoryorder="total ascending"),
                           legend=dict(orientation="h", y=-0.12),
                           xaxis_title="Open NCs", yaxis_title="")
        st.plotly_chart(figo, width='stretch', config=CHART_CONFIG, key="owner_open_sys")
        st.caption("Open NCs per owner, stacked by system. A tall orange or green block means "
                   "that person's workload has already moved to the new system.")

    # RAW view + download: the actual tracker NC rows behind the owner summary
    _own_raw = _q("SELECT * FROM nc WHERE source IN ('tracker','both')" + _OWN_FLT
                  + " ORDER BY owner, created_on DESC", _NB_PR)
    raw_block(_own_raw, key="owner_sum_raw", fname="owner_summary_raw.xlsx", sheet="Owner_raw",
              dl_label="Download tracker NCs")

st.markdown("")


# ------------------------------------------------------------------------------
# 10 - PER-OWNER DETAIL
# ------------------------------------------------------------------------------
st.markdown('<div id="owner-detail"></div>', unsafe_allow_html=True)
st.subheader("Per-owner detail")
st.caption("Pick a name to see only their NCs, what is missing on each, and a workbook you "
           "can send them as-is. Nobody else's rows appear.")

df_mine = None
_own_list = _q("SELECT DISTINCT owner FROM nc WHERE owner IS NOT NULL "
               "AND source IN ('tracker','both') ORDER BY owner")["owner"].tolist()

if not _own_list:
    st.info("No owners in the tracker data.")
else:
    _pick_owner = st.selectbox("Owner", ["— select —"] + _own_list, key="my_owner",
                               help="Owner is recorded in the NC tracker only. "
                                    "SAP-only NCs have no owner and appear for nobody.")
    if _pick_owner != "— select —":
        df_mine = _q("""
            SELECT n.nc_id AS "NC number",
                   COALESCE(n.system,'—')          AS "System",
                   COALESCE(n.tc_id,'—')           AS "TC ID",
                   COALESCE(n.project,'—')         AS "Project",
                   COALESCE(n.flight_unit,'—')     AS "Flight Unit",
                   COALESCE(n.status_state, CASE WHEN n.is_open=1 THEN 'Open'
                        WHEN n.is_open=0 THEN 'Closed' ELSE '(no status)' END) AS "Status",
                   COALESCE(n.classification,'—')  AS "Classification",
                   COALESCE(n.detection_area,'—')  AS "Detection",
                   COALESCE(n.nrb_disposition,'—') AS "NRB disposition",
                   n.created_on AS "Created", n.days_open AS "Days open"
            FROM nc n WHERE n.owner = ? AND n.created_on >= ? AND n.created_on <= ?
            ORDER BY n.created_on DESC""",
                     [_pick_owner, str(date_from), str(date_to) + " 23:59:59"])

        _n_all = len(df_mine)
        _n_op = int((df_mine["Status"] == "Open").sum())
        _n_cl = int((df_mine["Status"] == "Closed").sum())
        _n_ns = int((df_mine["Status"] == "(no status)").sum())
        _by_sys = df_mine[df_mine["Status"] == "Open"]["System"].value_counts().to_dict()
        _sys_txt = " · ".join(f"{_SYS_LABEL.get(k, k)} {v}" for k, v in _by_sys.items()) or "none open"

        for _c, (_l, _v, _d) in zip(st.columns(4), [
            ("Total NCs", _n_all, f"NCs assigned to {_pick_owner}, created inside the date window above."),
            ("Open", _n_op, _sys_txt),
            ("Closed", _n_cl, "Status says Closed."),
            ("No status", _n_ns, "Status cell blank — neither open nor closed."),
        ]):
            _c.markdown(_kpi_card(_l, _v, _d), unsafe_allow_html=True)

        _mp = df_mine["Project"].value_counts().reset_index()
        _mp.columns = ["project", "n"]
        if not _mp.empty:
            st.markdown("**By project**")
            _mp_colors = ["#E53E3E" if p == "—" else "#1E2761" for p in _mp["project"]]
            figm = px.bar(_mp, x="n", y="project", orientation="h", text="n",
                          color="project", color_discrete_sequence=_mp_colors)
            figm.update_layout(height=max(200, 34 * len(_mp)),
                               margin=dict(l=0, r=0, t=MODEBAR_T, b=0), showlegend=False,
                               yaxis=dict(categoryorder="total ascending"),
                               xaxis_title="NCs", yaxis_title="")
            figm.update_traces(textposition="outside",
                               hovertemplate="<b>%{y}</b><br>NCs: %{x}<extra></extra>")
            _headroom(figm, _mp["n"], horizontal=True)
            st.plotly_chart(figm, width='stretch', config=CHART_CONFIG, key="owner_proj")
            if "—" in _mp["project"].values:
                st.caption("Red = no project recorded on those NCs.")

        st.markdown(f"**{_pick_owner}'s NCs**")
        st.dataframe(df_mine, width='stretch', hide_index=True, height=380)

        _gap_cols = ["Project", "Flight Unit", "Classification", "Detection", "NRB disposition"]
        _gaps = pd.DataFrame({
            "Field": _gap_cols,
            "Missing": [int((df_mine[c] == "—").sum()) for c in _gap_cols],
            "Of": _n_all,
        })
        _gaps["%"] = (100 * _gaps["Missing"] / _n_all).round(0).astype(int) if _n_all else 0
        _gaps = _gaps[_gaps["Missing"] > 0]
        if not _gaps.empty:
            st.markdown("**What is missing on these NCs**")
            for _r in _gaps.itertuples(index=False):
                st.markdown(
                    f"- **{_r.Field}** — missing on **{int(_r.Missing)}** of "
                    f"{int(_r.Of)} NCs (**{int(_r._3)}%**)")
        else:
            st.success(f"Nothing missing — all {_n_all} NCs have every field filled in.")

        # RAW download: this owner's full NC rows (all columns), not the trimmed view
        _mine_raw = _q("SELECT * FROM nc WHERE owner = ? ORDER BY created_on DESC",
                       [_pick_owner])
        _pkg = {"My_NCs_raw": _mine_raw, "My_NCs_view": df_mine}
        if not _gaps.empty:
            _pkg["Missing_Fields"] = _gaps
        st.download_button(
            f"📥 Download {_pick_owner}'s report — raw NC rows (.xlsx)",
            build_full_report_bytes(_pkg, include_since=False),
            f"NC_report_{_pick_owner.replace(' ', '_').replace('/', '-')}.xlsx",
            XLSX_MIME, key="dl_mine", type="primary")
        st.caption("One workbook, one person: their raw NC rows + the tidy view + their gaps. Safe to send.")

st.markdown("")