"""
capa_view.py — the CAPA tab, rendered inside app.py via `capa_view.render()`.

A zoomable "pizza" (Plotly sunburst/icicle) that goes from the big picture down
to the detail: Launcher class -> Project -> Classification -> CAPA status. A
burnout split by launcher class sits on top, and an RCA-by-department pizza on
the side.

CAPA STATUS — CORRECTED (Adriele): the source of truth is the **Capa Board**
sheet (its `capa_status`), NOT "does the NC have a record". Each CAPA on the
board is Closed / Overdue / Open. Per NC:
    CAPA done    = its CAPA(s) are all Closed
    CAPA overdue = it has a CAPA marked Overdue
    CAPA open    = it has a CAPA still open (or blank) and not overdue
An NC with **no CAPA on the board is not shown here** — it is not "open". This
kills the old phantom (2,755 "open" that were really just NCs without a record).

Launcher class is read straight from the DB column `launcher_class`, which folds
Vulcan into LLV and keeps SAS standalone (LLV / MLV / SLV / SAS).

This module exposes render(); app.py calls it from inside the "CAPA" tab.
Page config and the password gate live in app.py, not here.
"""
import io
import json
import sqlite3
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

DB_FILE = "quality.db"
FEEDBACK_DB = "dashboard_feedback.db"   # persistent — ingest.py never touches it

# CAPA states come from the Capa Board (capa_status), not from record-existence.
DONE = "CAPA done"        # board status = Closed
OPEN = "CAPA open"        # board status = Open / blank (outstanding, not overdue)
OVERDUE = "CAPA overdue"  # board status = Overdue
STATUS_COLORS = {DONE: "#4CAF50", OPEN: "#F2A900", OVERDUE: "#E53E3E"}
# Launcher class read straight from the DB (Vulcan folded into LLV, SAS alone).
CLASS_ORDER = ["LLV", "MLV", "SLV", "SAS", "(no class)"]
CLASS_COLORS = {"LLV": "#1E2761", "MLV": "#F26E21", "SLV": "#5A63A0",
                "SAS": "#1C7293", "(no class)": "#B0B0B0"}


def _norm_status(s):
    s = str(s).strip().lower()
    if s == "closed":
        return "Closed"
    if s == "overdue":
        return "Overdue"
    return "Open"   # 'open' or blank -> still outstanding


def _board_state(status_list):
    """Collapse an NC's CAPA-record statuses (from the Capa Board) into one
    state. Overdue wins, then any not-closed = open, else done."""
    if any(s == "Overdue" for s in status_list):
        return OVERDUE
    if any(s != "Closed" for s in status_list):
        return OPEN
    return DONE


# =========================================================================
# feedback button (top of every tool — Adriele's standing rule)
# =========================================================================
def feedback_widget():
    with st.expander("💡  Improve this tab — send the team your idea",
                     expanded=False):
        with st.form("capa_feedback", clear_on_submit=True):
            c1, c2 = st.columns([1, 1])
            name = c1.text_input("Your name (optional)")
            cat = c2.selectbox("Type", ["Idea / improvement", "Something's wrong",
                                        "Question", "Other"])
            msg = st.text_area("What would make this better?")
            if st.form_submit_button("Send") and msg.strip():
                try:
                    con = sqlite3.connect(FEEDBACK_DB)
                    con.execute("CREATE TABLE IF NOT EXISTS feedback("
                                "ts TEXT, page TEXT, name TEXT, category TEXT, "
                                "message TEXT)")
                    con.execute("INSERT INTO feedback VALUES (?,?,?,?,?)",
                                (datetime.now().isoformat(timespec="seconds"),
                                 "CAPA", name, cat, msg.strip()))
                    con.commit()
                    con.close()
                    st.success("Thank you! Sent to the team.")
                except Exception as e:
                    st.error(f"Could not save: {e}")


# =========================================================================
# data
# =========================================================================
def _classify(c):
    """Normalise a raw classification into Major / Minor / (no class) / other."""
    s = str(c)
    if s.startswith("Major"):
        return "Major"
    if "Minor" in s:
        return "Minor"
    if c is None or s in ("0", "None", "(no class)", "nan"):
        return "(no class)"
    return s


@st.cache_data
def _plotly_js():
    """The plotly.min.js that ships with the installed plotly package, inlined
    so the icicle component needs NO internet — works on Streamlit Cloud and on
    an air-gapped internal server alike."""
    import plotly, os
    p = os.path.join(os.path.dirname(plotly.__file__),
                     "package_data", "plotly.min.js")
    with open(p, encoding="utf-8") as f:
        return f.read()


_STATUS_BUCKET = {"Closed": DONE, "Overdue": OVERDUE, "Open": OPEN}


@st.cache_data(ttl=300)
def load_capa_rows():
    """ONE ROW PER CAPA on the Capa Board — the source of truth. Each CAPA is
    counted by its own board Status: Closed = done, Overdue = overdue, anything
    else (Open / blank) = open. Every row also carries its NC's context (project,
    class, owner, detection area) via a left join, so the drill and the detail
    table work. Launcher class comes from the board's own 'Affected Project'
    prefix (LLV folds in Vulcan, SAS stands alone)."""
    con = sqlite3.connect(DB_FILE)
    df = pd.read_sql(
        "SELECT c.nc_id, c.capa_type, c.capa_status, "
        "c.launcher_class AS capa_launcher, "
        "n.project, n.launcher_class AS nc_launcher, n.classification, "
        "n.owner, n.detection_area, n.defect_code_text, n.created_on, "
        "n.status_state "
        "FROM capa c LEFT JOIN nc n ON c.nc_id = n.nc_id", con)
    con.close()
    df["CAPA"] = df["capa_status"].map(_norm_status).map(_STATUS_BUCKET).fillna(OPEN)
    df["Launcher"] = df["capa_launcher"].fillna(df["nc_launcher"]).fillna("(no class)")
    df["Project"] = df["project"].fillna("(no project)")
    df["Class"] = df["classification"].apply(_classify)
    return df


@st.cache_data(ttl=300)
def load_capa_ncs():
    """ONE ROW PER NC that has a CAPA on the Capa Board — this is what the whole
    tab counts. A single NC usually carries several CAPA lines on the board (RCA +
    correction + corrective + preventive), so counting board *lines* inflated the
    numbers into the thousands (LLV read ~1149). We instead count **NCs**: each
    NC's CAPA state is the worst of its board lines — any Overdue -> overdue, else
    any not-Closed -> open, else done. Now LLV reads its real NC count (~451)."""
    rows = load_capa_rows()   # per CAPA line

    def _collapse(status_series):
        return _board_state([_norm_status(s) for s in status_series])

    agg = (rows.groupby("nc_id", as_index=False)
           .agg(Launcher=("Launcher", "first"),
                Project=("Project", "first"),
                Class=("Class", "first"),
                owner=("owner", "first"),
                detection_area=("detection_area", "first"),
                defect_code_text=("defect_code_text", "first"),
                created_on=("created_on", "first"),
                status_state=("status_state", "first"),
                CAPA=("capa_status", _collapse)))
    return agg


@st.cache_data(ttl=300)
def load_rca_departments():
    """RCA rows joined to their NC so they carry Launcher / Project / Class.
    This lets the RCA donuts follow the SAME drill as the icicle."""
    con = sqlite3.connect(DB_FILE)
    df = pd.read_sql(
        "SELECT c.nc_id, "
        "COALESCE(c.origin_area_l1,'(not recorded)') AS dept, "
        "COALESCE(c.rc_category_l1,'(not recorded)') AS cause, "
        "n.project AS project, n.launcher_class AS launcher_class, "
        "n.classification AS classification "
        "FROM capa c LEFT JOIN nc n ON c.nc_id = n.nc_id "
        "WHERE c.capa_type='RCA'", con)
    con.close()
    # One RCA per NC for the donuts (ingest no longer dedups the capa table, so
    # do it here): prefer a row that actually records an origin/cause.
    df["_score"] = ((df["dept"] != "(not recorded)").astype(int)
                    + (df["cause"] != "(not recorded)").astype(int))
    df = (df.sort_values("_score", kind="stable")
            .drop_duplicates(subset=["nc_id"], keep="last")
            .drop(columns=["_score"]).reset_index(drop=True))
    df["Launcher"] = df["launcher_class"].fillna("(no class)")
    df["Project"] = df["project"].fillna("(no project)")
    df["Class"] = df["classification"].apply(_classify)
    return df


# =========================================================================
# zoomable vertical icicle (replaces the sunburst) — mouse-only pan/zoom
# =========================================================================
def _build_icicle_data(df):
    """Launcher -> Class -> CAPA status, as Plotly icicle arrays.
    Values are counts; colours: navy root, launcher-class colour, grey class,
    green/amber/red CAPA leaves."""
    ids, labels, parents, values, colors = [], [], [], [], []
    tot = len(df)
    if tot == 0:
        return {"ids": [], "labels": [], "parents": [],
                "values": [], "colors": []}
    ids.append("All"); labels.append("All NCs"); parents.append("")
    values.append(tot); colors.append("#2B3A67")
    for lau, gl in df.groupby("Launcher"):
        lid = f"All/{lau}"
        ids.append(lid); labels.append(str(lau)); parents.append("All")
        values.append(len(gl)); colors.append(CLASS_COLORS.get(lau, "#888888"))
        for cls, gc in gl.groupby("Class"):
            cid = f"{lid}/{cls}"
            ids.append(cid); labels.append(str(cls)); parents.append(lid)
            values.append(len(gc)); colors.append("#9AA6C4")
            for stt, gs in gc.groupby("CAPA"):
                sid = f"{cid}/{stt}"
                ids.append(sid); labels.append(str(stt)); parents.append(cid)
                values.append(len(gs))
                colors.append(STATUS_COLORS.get(stt, "#888888"))
    return {"ids": ids, "labels": labels, "parents": parents,
            "values": values, "colors": colors}


_ICICLE_TEMPLATE = """<!DOCTYPE html><html><head><meta charset="utf-8"/>
__PLOTLY_TAG__
<style>
 *{box-sizing:border-box;}
 body{margin:0;font-family:system-ui,-apple-system,"Segoe UI",sans-serif;}
 .hint{display:inline-flex;gap:14px;font-size:.78rem;color:#5B6B78;
   background:#F1F3F8;border:1px solid #E1E5EE;border-radius:8px;
   padding:5px 11px;margin:0 0 8px;}
 .hint b{color:#1E2761;}
 .stage{position:relative;height:600px;border:1px solid #E1E5EE;
   border-radius:12px;overflow:hidden;background:#FBFCFE;cursor:grab;
   touch-action:none;}
 .stage.dragging{cursor:grabbing;}
 .canvas{position:absolute;inset:0;transform-origin:0 0;will-change:transform;}
 #chart{width:100%;height:100%;}
 .zbadge{position:absolute;right:12px;top:10px;z-index:5;background:#1E2761;
   color:#fff;font-size:.72rem;font-weight:700;padding:4px 9px;
   border-radius:20px;opacity:.9;}
 .reset{position:absolute;right:12px;bottom:12px;z-index:5;background:#fff;
   color:#1E2761;border:1px solid #1E2761;border-radius:8px;font-size:.74rem;
   font-weight:700;padding:5px 10px;cursor:pointer;}
</style></head><body>
 <div class="hint"><span>&#128433; <b>Scroll</b> = zoom in / out</span>
   <span><b>Drag</b> = pan</span><span>Double-click = reset</span></div>
 <div class="stage" id="stage">
   <div class="zbadge" id="zbadge">100%</div>
   <button class="reset" id="reset">Reset view</button>
   <div class="canvas" id="canvas"><div id="chart"></div></div>
 </div>
<script id="data" type="application/json">__DATA__</script>
<script>
var D=JSON.parse(document.getElementById('data').textContent);
var fig=[{type:'icicle',ids:D.ids,labels:D.labels,parents:D.parents,
  values:D.values,branchvalues:'total',tiling:{orientation:'v',pad:1},
  marker:{colors:D.colors,line:{width:1,color:'#ffffff'}},
  textfont:{size:12,family:'system-ui'},
  hovertemplate:'<b>%{label}</b><br>%{value} NCs<extra></extra>',
  root:{color:'#EEF1F7'},pathbar:{visible:true,thickness:22}}];
var layout={margin:{t:26,l:6,r:6,b:6},paper_bgcolor:'rgba(0,0,0,0)',
  font:{family:'system-ui'}};
Plotly.newPlot('chart',fig,layout,{displayModeBar:false,responsive:true});
var stage=document.getElementById('stage'),canvas=document.getElementById('canvas'),
    zbadge=document.getElementById('zbadge');
var scale=1,tx=0,ty=0,MIN=1,MAX=8;
function apply(){canvas.style.transform='translate('+tx+'px,'+ty+'px) scale('+scale+')';
  zbadge.textContent=Math.round(scale*100)+'%';}
stage.addEventListener('wheel',function(e){e.preventDefault();
  var r=stage.getBoundingClientRect(),cx=e.clientX-r.left,cy=e.clientY-r.top;
  var f=e.deltaY<0?1.15:1/1.15,ns=Math.min(MAX,Math.max(MIN,scale*f));
  tx=cx-(cx-tx)*(ns/scale);ty=cy-(cy-ty)*(ns/scale);scale=ns;
  if(scale===1){tx=0;ty=0;}apply();},{passive:false});
var drag=false,sx=0,sy=0;
stage.addEventListener('mousedown',function(e){drag=true;sx=e.clientX-tx;
  sy=e.clientY-ty;stage.classList.add('dragging');});
window.addEventListener('mousemove',function(e){if(!drag)return;
  tx=e.clientX-sx;ty=e.clientY-sy;apply();});
window.addEventListener('mouseup',function(){drag=false;
  stage.classList.remove('dragging');});
stage.addEventListener('dblclick',function(){scale=1;tx=0;ty=0;apply();});
document.getElementById('reset').onclick=function(){scale=1;tx=0;ty=0;apply();};
window.__setZoom=function(s,cx,cy){var r=stage.getBoundingClientRect();
  if(cx==null)cx=r.width/2;if(cy==null)cy=r.height/2;
  tx=cx-(cx-tx)*(s/scale);ty=cy-(cy-ty)*(s/scale);scale=s;apply();};
</script></body></html>"""


def icicle_html(df):
    """Standalone HTML for the zoomable vertical icicle of the current view.
    Plotly is inlined (no CDN, no internet needed)."""
    payload = json.dumps(_build_icicle_data(df))
    tag = "<script>" + _plotly_js() + "</script>"
    return (_ICICLE_TEMPLATE
            .replace("__PLOTLY_TAG__", tag)
            .replace("__DATA__", payload))


# =========================================================================
# full cross-filter CAPA page — one self-contained HTML/JS page, rebuilt from
# the live DB on every run. Same page runs standalone on the internal server.
# =========================================================================
import os as _os
_HERE = _os.path.dirname(_os.path.abspath(__file__))


@st.cache_data
def _sheetjs():
    """SheetJS (client-side Excel writer), shipped in the repo so the page
    works with no internet — on Streamlit Cloud and on an air-gapped server."""
    with open(_os.path.join(_HERE, "xlsx.full.min.js"), encoding="utf-8") as f:
        return f.read()


@st.cache_data
def _page_template():
    with open(_os.path.join(_HERE, "capa_page_template.html"),
              encoding="utf-8") as f:
        return f.read()


def _page_data(nc_df, rca_df):
    nc_cols = ["nc_id", "Launcher", "Project", "Class", "CAPA",
               "status_state", "owner", "detection_area", "defect_code_text",
               "created_on"]
    rca_cols = ["nc_id", "Launcher", "Project", "Class", "dept", "cause"]
    nc2 = nc_df[nc_cols].fillna("").astype(str).copy()
    nc2["owner"] = nc2["owner"].replace("", "(unassigned)")
    nc2["detection_area"] = nc2["detection_area"].replace("", "(not recorded)")
    nc2["defect_code_text"] = nc2["defect_code_text"].replace("", "(not coded)")
    return {"nc": nc2.to_dict("records"),
            "rca": rca_df[rca_cols].fillna("").astype(str).to_dict("records")}


def capa_page_html(nc_df, rca_df):
    """Build the full cross-filter CAPA page for the CURRENT data.
    Plotly + SheetJS + the data are all inlined -> no internet needed."""
    payload = json.dumps(_page_data(nc_df, rca_df))
    return (_page_template()
            .replace("__PLOTLY__", _plotly_js(), 1)
            .replace("__SHEETJS__", _sheetjs(), 1)
            .replace("__DATA__", payload, 1))


# =========================================================================
# render — called by app.py from inside the CAPA tab
# =========================================================================
def render():
    feedback_widget()
    st.title("CAPA — Root Cause & Actions")
    st.caption("Source: the **Capa Board** sheet only. Numbers count **NCs** — "
               "each NC's CAPA state is the worst of its board lines: **overdue** "
               "if any line is overdue, else **open** if any line is still open, "
               "else **done**. LLV includes Vulcan; SAS stands alone.")

    nc = load_capa_ncs()
    if nc.empty:
        st.info("No CAPA records found on the board in `quality.db`. "
                "Run the ingest with the CAPA tracker present.")
        return

    # ---- burnout split by launcher class -------------------------------
    st.subheader("CAPA burnout by launcher class")
    st.caption("Each launcher class, counting **NCs**: **done** (green, all board "
               "lines Closed) vs **open** (amber) vs **overdue** (red). "
               "LLV includes Vulcan; SAS stands alone.")
    burn = (nc.groupby(["Launcher", "CAPA"]).size().reset_index(name="n"))
    burn["Launcher"] = pd.Categorical(burn["Launcher"], categories=CLASS_ORDER,
                                      ordered=True)
    burn = burn.sort_values("Launcher")
    figb = px.bar(burn, x="Launcher", y="n", color="CAPA",
                  color_discrete_map=STATUS_COLORS, text="n",
                  category_orders={"Launcher": CLASS_ORDER,
                                   "CAPA": [DONE, OPEN, OVERDUE]})
    figb.update_layout(barmode="stack", height=340,
                       margin=dict(t=10, l=0, r=0, b=0),
                       xaxis_title="", yaxis_title="NCs",
                       legend_title_text="")
    figb.update_traces(textposition="inside")
    st.plotly_chart(figb, width='stretch')

    # coverage table under the burnout
    piv = (nc.pivot_table(index="Launcher", columns="CAPA", values="nc_id",
                          aggfunc="count", fill_value=0)
             .reindex(CLASS_ORDER).dropna(how="all"))
    for _col in (DONE, OPEN, OVERDUE):
        if _col not in piv.columns:
            piv[_col] = 0
    piv["Total"] = piv[DONE] + piv[OPEN] + piv[OVERDUE]
    piv["Coverage"] = (100 * piv[DONE] / piv["Total"].replace(0, 1)).round(0)
    piv = piv.rename(columns={DONE: "Done", OPEN: "Open", OVERDUE: "Overdue"})
    st.dataframe(piv[["Done", "Open", "Overdue", "Total", "Coverage"]]
                 .style.format({"Coverage": "{:.0f}%"}),
                 width='stretch')
    st.caption("Counting NCs. ‘Done’ = all board lines Closed · ‘Overdue’ = has an "
               "overdue line · ‘Open’ = still open · ‘Coverage’ = done ÷ total.")

    st.divider()

    # ---- the full cross-filter CAPA view (one self-contained HTML/JS page) --
    # Only the CAPAs that still need action: open + overdue. The done ones are
    # already counted in the burnout / coverage above; here we focus the drill on
    # what is still outstanding, so the pizza never reads in the thousands.
    st.subheader("Still open — cross-filter the outstanding CAPAs")
    st.caption("This panel shows **only open and overdue** CAPAs (done ones are "
               "excluded). Counting NCs. Click the pizza to drill; the donuts and "
               "table follow your selection.")
    nc_action = nc[nc["CAPA"].isin([OPEN, OVERDUE])].reset_index(drop=True)
    if nc_action.empty:
        st.info("No open or overdue CAPAs in the board — all are closed.")
        return
    rca_df = load_rca_departments()
    rca_df = rca_df[rca_df["nc_id"].isin(set(nc_action["nc_id"]))].reset_index(drop=True)
    components.html(capa_page_html(nc_action, rca_df), height=2150, scrolling=False)