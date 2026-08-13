"""
customer_complaints.py — the "Customer Complaints" tab.

Customer complaints are the highest-severity NCs (the customer already felt the
problem), so they get their own page instead of being one bar on a chart.

Wire it in app.py exactly like the CAPA tab, e.g.:

    import customer_complaints
    tab_dash, tab_cc, tab_capa = st.tabs(
        ["Quality Dashboard", "Customer Complaints", "CAPA"])
    with tab_dash:
        exec(open("dashboard_body.py", encoding="utf-8").read())
    with tab_cc:
        customer_complaints.render()
    with tab_capa:
        capa_view.render()
"""
from datetime import datetime, date
from io import BytesIO
from pathlib import Path
import sqlite3

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

DB_FILE = "quality.db"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# A customer complaint = detection area is the CC code, in either wording
# (SAP free-text "Customer complaint" or the tracker code "CC: Customer Complaint").
CC_WHERE = ("(detection_area LIKE 'CC%' "
            "OR LOWER(detection_area) LIKE '%customer%complaint%')")

NAVY = "#1E2761"
ORANGE = "#F26E21"
RED = "#E53E3E"
GREEN = "#4CAF50"
CLASS_COLORS = {"LLV": NAVY, "MLV": ORANGE, "SLV": GREEN, "SAS": "#1C7293",
                "(no class)": "#B0B7C3"}
CHART_CONFIG = {"displayModeBar": "hover", "displaylogo": False}
MODEBAR_T = 40


def _conn():
    return sqlite3.connect(DB_FILE)


def _q(sql, params=None):
    with _conn() as c:
        return pd.read_sql_query(sql, c, params=params or [])


def _to_excel(df, sheet="Customer_complaints"):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=sheet[:31], index=False)
        ws = w.sheets[sheet[:31]]
        for col in ("A", "B", "C", "D", "E", "F", "G", "H"):
            ws.column_dimensions[col].width = 22
    return buf.getvalue()


def _raw_block(df, key, fname, note=None, preview_rows=200):
    """👁 View raw data toggle + 📥 Download — same rows, current filter."""
    n = len(df)
    c1, c2, _sp = st.columns([1.3, 1.4, 3.3])
    show = c1.toggle("👁 View raw data", key=f"cc_{key}__view",
                     help=f"Show the {n} raw customer-complaint rows on screen.")
    c2.download_button(f"📥 Download ({n})", _to_excel(df), fname, XLSX_MIME,
                       key=f"cc_{key}__dl")
    if show:
        if note:
            st.caption(note)
        if n:
            st.dataframe(df.head(preview_rows), width="stretch", hide_index=True,
                         height=min(400, 80 + 28 * min(n, 11)))
            if n > preview_rows:
                st.caption(f"Showing the first {preview_rows} of {n} rows — the download has all {n}.")
        else:
            st.info("No customer complaints for this selection.")


def render():
    if not Path(DB_FILE).exists():
        st.error(f"`{DB_FILE}` not found. Run `python ingest.py` first.")
        return

    # ---------- compact filter ----------
    _today = date.today()
    _default_from = date(_today.year, 1, 1)
    _min = _q(f"SELECT MIN(created_on) m FROM nc WHERE {CC_WHERE} AND created_on IS NOT NULL")
    _pick_min = date(2015, 1, 1)
    if not _min.empty and _min.iloc[0]["m"]:
        try:
            _pick_min = pd.to_datetime(_min.iloc[0]["m"]).date()
        except Exception:
            pass
    _projects = _q(f"SELECT DISTINCT COALESCE(project,'(no project)') p FROM nc WHERE {CC_WHERE} ORDER BY p")["p"].tolist()
    _owners = _q(f"SELECT DISTINCT COALESCE(owner,'(no owner)') o FROM nc WHERE {CC_WHERE} ORDER BY o")["o"].tolist()
    _sources = _q(f"SELECT DISTINCT source FROM nc WHERE {CC_WHERE} AND source IS NOT NULL ORDER BY source")["source"].tolist()

    _sf = st.session_state.get("cc_from", _default_from)
    _st_ = st.session_state.get("cc_to", _today)
    # Keep the picker defaults inside [_pick_min, max]. With tracker-only data the
    # earliest complaint is Feb 2026, so a default of 01 Jan 2026 falls below
    # min_value and Streamlit raises. Clamp both ends.
    _max_pick = date(_today.year + 1, 12, 31)
    _sf = min(max(_sf, _pick_min), _max_pick)
    _st_ = min(max(_st_, _pick_min), _max_pick)
    _proj_sel = st.session_state.get("cc_proj", [])
    _stat_sel = st.session_state.get("cc_status", "All")
    _type_sel = st.session_state.get("cc_type", "All")
    _owner_sel = st.session_state.get("cc_owner", [])
    _src_sel = st.session_state.get("cc_source", "All")
    _lbl_bits = [f"{_sf.strftime('%d %b %Y')} → {_st_.strftime('%d %b %Y')}"]
    if _stat_sel != "All":
        _lbl_bits.append(_stat_sel)
    if _type_sel != "All":
        _lbl_bits.append(_type_sel)
    if _proj_sel:
        _lbl_bits.append(f"{len(_proj_sel)} project(s)")
    if _owner_sel:
        _lbl_bits.append(f"{len(_owner_sel)} owner(s)")
    if _src_sel != "All":
        _lbl_bits.append(_src_sel)
    with st.expander("🔍  Filter  ·  " + "   ·   ".join(_lbl_bits), expanded=False):
        r1c1, r1c2, r1c3, r1c4 = st.columns(4)
        f_from = r1c1.date_input("From", value=_sf, min_value=_pick_min,
                                 max_value=date(_today.year + 1, 12, 31),
                                 key="cc_from_w", format="DD/MM/YYYY")
        f_to = r1c2.date_input("To", value=_st_, min_value=_pick_min,
                               max_value=date(_today.year + 1, 12, 31),
                               key="cc_to_w", format="DD/MM/YYYY")
        f_status = r1c3.selectbox("Status", ["All", "Open", "Closed"],
                                  index=["All", "Open", "Closed"].index(_stat_sel),
                                  key="cc_status_w")
        f_type = r1c4.selectbox("NC type", ["All", "Production", "Supplier"],
                                index=["All", "Production", "Supplier"].index(_type_sel),
                                key="cc_type_w")
        r2c1, r2c2, r2c3 = st.columns(3)
        f_proj = r2c1.multiselect("Project", _projects,
                                  default=[p for p in _proj_sel if p in _projects],
                                  key="cc_proj_w")
        f_owner = r2c2.multiselect("Owner", _owners,
                                   default=[o for o in _owner_sel if o in _owners],
                                   key="cc_owner_w")
        if _sources:
            f_source = r2c3.selectbox("Data source", ["All"] + _sources,
                                      index=(["All"] + _sources).index(_src_sel)
                                      if _src_sel in (["All"] + _sources) else 0,
                                      key="cc_source_w")
        else:
            f_source = "All"
        b1, b2, _ = st.columns([1, 1, 4])
        if b1.button("✓ Apply", type="primary", key="cc_apply"):
            st.session_state["cc_from"] = f_from
            st.session_state["cc_to"] = f_to
            st.session_state["cc_status"] = f_status
            st.session_state["cc_type"] = f_type
            st.session_state["cc_proj"] = f_proj
            st.session_state["cc_owner"] = f_owner
            st.session_state["cc_source"] = f_source
            st.rerun()
        if b2.button("This year", key="cc_reset"):
            for k in ("cc_from", "cc_to", "cc_status", "cc_type", "cc_proj",
                      "cc_owner", "cc_source", "cc_from_w", "cc_to_w",
                      "cc_status_w", "cc_type_w", "cc_proj_w", "cc_owner_w",
                      "cc_source_w"):
                st.session_state.pop(k, None)
            st.rerun()

    # committed filter → WHERE clause (governs the ENTIRE page)
    d_from = st.session_state.get("cc_from", _default_from)
    d_to = st.session_state.get("cc_to", _today)
    if d_from > d_to:
        d_from, d_to = d_to, d_from
    status = st.session_state.get("cc_status", "All")
    nctype = st.session_state.get("cc_type", "All")
    proj = st.session_state.get("cc_proj", [])
    owner = st.session_state.get("cc_owner", [])
    source = st.session_state.get("cc_source", "All")

    where = [CC_WHERE, "created_on >= ?", "created_on <= ?"]
    params = [str(d_from), str(d_to) + " 23:59:59"]
    if status == "Open":
        where.append("is_open = 1")
    elif status == "Closed":
        where.append("is_open = 0")
    if nctype == "Production":
        where.append("is_supplier_nc = 0")
    elif nctype == "Supplier":
        where.append("is_supplier_nc = 1")
    if proj:
        where.append(f"COALESCE(project,'(no project)') IN ({','.join(['?']*len(proj))})")
        params.extend(proj)
    if owner:
        where.append(f"COALESCE(owner,'(no owner)') IN ({','.join(['?']*len(owner))})")
        params.extend(owner)
    if source != "All":
        where.append("source = ?")
        params.append(source)
    W = "WHERE " + " AND ".join(where)

    st.caption(f"Showing customer complaints created between **{d_from.strftime('%d %b %Y')}** "
               f"and **{d_to.strftime('%d %b %Y')}**"
               + (f" · **{status}**" if status != "All" else "")
               + (f" · **{nctype}**" if nctype != "All" else "")
               + (f" · **{len(proj)} project(s)**" if proj else "")
               + (f" · **{len(owner)} owner(s)**" if owner else "")
               + (f" · **{source}**" if source != "All" else "") + ".")

    # ---------- KPIs ----------
    total = _q(f"SELECT COUNT(*) n FROM nc {W}", params).iloc[0]["n"]
    if total == 0:
        st.info("No customer complaints match this filter.")
        return
    open_n = _q(f"SELECT COUNT(*) n FROM nc {W} AND is_open=1", params).iloc[0]["n"]
    closed_n = _q(f"SELECT COUNT(*) n FROM nc {W} AND is_open=0", params).iloc[0]["n"]
    _yr = str(_today.year)
    ytd = _q(f"SELECT COUNT(*) n FROM nc {W} AND substr(created_on,1,4)=?",
             params + [_yr]).iloc[0]["n"]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Customer complaints", int(total), help="All CC-type NCs in the filter.")
    k2.metric("Still open", int(open_n),
              help="Customer complaints not yet closed.")
    k3.metric("Closed", int(closed_n))
    k4.metric(f"Opened in {_yr}", int(ytd))

    _raw_all = _q(f"SELECT * FROM nc {W} ORDER BY created_on DESC", params)
    _raw_block(_raw_all, key="all", fname="customer_complaints.xlsx",
               note="Every customer complaint in the current filter.")

    st.markdown("---")

    # ---------- Open complaints table (the dangerous ones) ----------
    st.subheader("Open customer complaints")
    df_open = _q(f"""
        SELECT nc_id AS "NC number",
               COALESCE(system,'—')        AS "System",
               COALESCE(project,'—')       AS "Project",
               COALESCE(flight_unit,'—')   AS "Flight unit",
               COALESCE(owner,'—')         AS "Owner",
               COALESCE(classification,'—') AS "Class",
               created_on                  AS "Created",
               days_open                   AS "Days open",
               COALESCE(description,'')    AS "Description"
        FROM nc {W} AND is_open=1
        ORDER BY days_open DESC, created_on ASC""", params)
    if df_open.empty:
        st.info("No open customer complaints in this selection.")
    else:
        st.caption(f"{len(df_open)} open — sorted by age (oldest first). Age in days is in "
                   "the **Days open** column.")
        st.dataframe(df_open, width="stretch", hide_index=True,
                     height=min(460, 80 + 34 * min(len(df_open), 11)))
        st.download_button(f"📥 Download open complaints ({len(df_open)})",
                           _to_excel(df_open, "Open_CC"), "open_customer_complaints.xlsx",
                           XLSX_MIME, key="cc_open_dl", type="primary")

    st.markdown("---")

    # ---------- Trend over time ----------
    st.subheader("Customer complaints over time")
    df_m = _q(f"""SELECT substr(created_on,1,7) AS month, COUNT(*) AS n
                  FROM nc {W} AND created_on IS NOT NULL GROUP BY month ORDER BY month""",
              params)
    if not df_m.empty:
        figm = px.bar(df_m, x="month", y="n", text="n",
                      color_discrete_sequence=[RED])
        figm.update_layout(height=300, margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                           xaxis_title="", yaxis_title="Complaints", showlegend=False)
        figm.update_traces(textposition="outside",
                           hovertemplate="<b>%{x}</b><br>Complaints: %{y}<extra></extra>")
        st.plotly_chart(figm, width="stretch", config=CHART_CONFIG, key="cc_trend")
        st.caption("Customer complaints by the month they were opened.")
    _raw_block(_raw_all, key="trend", fname="customer_complaints_trend.xlsx")

    st.markdown("---")

    # ---------- By project & by class ----------
    st.subheader("Where the complaints come from")
    cpj, ccl = st.columns(2)
    with cpj:
        st.markdown("**By project**")
        df_p = _q(f"""SELECT COALESCE(project,'(no project)') AS project,
                             COUNT(*) AS total,
                             SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) AS open
                      FROM nc {W} GROUP BY project ORDER BY total DESC""", params)
        if not df_p.empty:
            figp = go.Figure()
            figp.add_trace(go.Bar(y=df_p["project"], x=df_p["total"], orientation="h",
                                  name="Total", marker_color=NAVY, text=df_p["total"],
                                  hovertemplate="<b>%{y}</b><br>Total: %{x}<extra></extra>"))
            figp.add_trace(go.Bar(y=df_p["project"], x=df_p["open"], orientation="h",
                                  name="Open", marker_color=RED, text=df_p["open"],
                                  hovertemplate="<b>%{y}</b><br>Open: %{x}<extra></extra>"))
            figp.update_layout(height=max(300, 34 * len(df_p)), barmode="overlay",
                               margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                               legend=dict(orientation="h", y=-0.15),
                               yaxis=dict(categoryorder="total ascending"),
                               xaxis_title="Complaints", yaxis_title="")
            figp.update_traces(textposition="outside")
            st.plotly_chart(figp, width="stretch", config=CHART_CONFIG, key="cc_by_project")
    with ccl:
        st.markdown("**By launcher class**")
        df_c = _q(f"""SELECT COALESCE(launcher_class,'(no class)') AS cls, COUNT(*) AS n
                      FROM nc {W} GROUP BY cls ORDER BY n DESC""", params)
        if not df_c.empty:
            figc = px.pie(df_c, names="cls", values="n", hole=0.45,
                          color="cls", color_discrete_map=CLASS_COLORS)
            figc.update_layout(height=max(300, 34 * len(df_p) if not df_p.empty else 320),
                               margin=dict(l=0, r=0, t=MODEBAR_T, b=0),
                               legend=dict(orientation="h", y=-0.1))
            figc.update_traces(textinfo="label+value",
                               hovertemplate="<b>%{label}</b><br>Complaints: %{value}<extra></extra>")
            st.plotly_chart(figc, width="stretch", config=CHART_CONFIG, key="cc_by_class")
    _raw_block(_raw_all, key="where", fname="customer_complaints_by_project.xlsx",
               note="Every customer complaint in the current filter (each row carries its project & class).")

    st.caption(f"Data source: `{DB_FILE}` · last modified "
               f"{datetime.fromtimestamp(Path(DB_FILE).stat().st_mtime).strftime('%d.%m.%Y %H:%M')}")