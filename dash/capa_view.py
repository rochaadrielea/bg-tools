"""
capa_view.py — the CAPA tab, rendered inside app.py via `capa_view.render()`.

A zoomable "pizza" (Plotly sunburst) that goes from the big picture down to the
detail: Launcher class -> Project -> Classification -> CAPA open/done. Click a
wedge to zoom in, click the middle to zoom out, and the report table below
always shows exactly what you're looking at. A burnout split by launcher class
(LLV / MLV / SLV) sits on top, and an RCA-by-department pizza on the side.

RULE (Adriele, current): EVERY NC owes a CAPA. So an NC with no CAPA record on
file = CAPA OPEN (outstanding), whether the NC is closed or not. An NC that has
a CAPA record (RCA / CA / PA / Ext-8D) = CAPA DONE.

  * "Done" now includes Ext-8D (external supplier 8D). It is a corrective-action
    record like the others, so an NC whose only action is an Ext-8D counts as
    covered. (Before the ingest fix, Ext-8D rows were dropped and 131 covered
    NCs wrongly showed as open.)

Launcher classes — confirmed against the CAPA tracker's own "Affected Project"
column, which literally prefixes LLV_ / MLV_ / SLV_:
    LLV = Ariane (A6) + Relativity (RS) + MHI_H3 (H3)   [+ Atlas, Vulcan share the LLV prefix]
    MLV = Vega
    SLV = Flexline + SAS
    (Vulcan is kept as its own bucket; no-project rows show as '(no project)')

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

LAUNCHER = {"Ariane": "LLV", "Relativity": "LLV", "MHI_H3": "LLV",
            "Vega": "MLV", "Flexline": "SLV", "SAS": "SLV", "Vulcan": "Vulcan"}
OPEN, DONE = "CAPA open", "CAPA done"
STATUS_COLORS = {OPEN: "#E53E3E", DONE: "#4CAF50"}
# fixed colours per launcher class so the burnout bars read consistently
CLASS_COLORS = {"LLV": "#1E2761", "MLV": "#F26E21", "SLV": "#5A63A0",
                "Vulcan": "#4CAF50", "(no project)": "#B0B0B0"}


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


@st.cache_data(ttl=300)
def load_nc():
    con = sqlite3.connect(DB_FILE)
    nc = pd.read_sql(
        "SELECT nc_id, project, classification, owner, detection_area, "
        "defect_code_text, created_on, closure_date, is_open, status_state "
        "FROM nc", con)
    capa_ncs = set(pd.read_sql("SELECT DISTINCT nc_id FROM capa", con)["nc_id"])
    con.close()
    nc["Launcher"] = nc["project"].map(LAUNCHER).fillna("(no project)")
    nc["Project"] = nc["project"].fillna("(no project)")

    nc["Class"] = nc["classification"].apply(_classify)
    nc["CAPA"] = nc["nc_id"].apply(lambda x: DONE if x in capa_ncs else OPEN)
    return nc


@st.cache_data(ttl=300)
def load_rca_departments():
    """RCA rows joined to their NC so they carry Launcher / Project / Class.
    This lets the RCA donuts follow the SAME drill as the icicle: pick a
    launcher and the department + root-cause donuts re-scope to it."""
    con = sqlite3.connect(DB_FILE)
    df = pd.read_sql(
        "SELECT c.nc_id, "
        "COALESCE(c.origin_area_l1,'(not recorded)') AS dept, "
        "COALESCE(c.rc_category_l1,'(not recorded)') AS cause, "
        "n.project AS project, n.classification AS classification "
        "FROM capa c LEFT JOIN nc n ON c.nc_id = n.nc_id "
        "WHERE c.capa_type='RCA'", con)
    con.close()
    df["Launcher"] = df["project"].map(LAUNCHER).fillna("(no project)")
    df["Project"] = df["project"].fillna("(no project)")
    df["Class"] = df["classification"].apply(_classify)
    return df


# =========================================================================
# zoomable vertical icicle (replaces the sunburst) — mouse-only pan/zoom
# =========================================================================
def _build_icicle_data(df):
    """Launcher -> Class -> CAPA open/done, as Plotly icicle arrays.
    Values are counts; colours: navy root, launcher-class colour, grey class,
    red/green CAPA leaves."""
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
    st.caption("Rule: every NC owes a CAPA. No CAPA on record = **CAPA open** "
               "(even if the NC is closed). 'Done' counts RCA, CA, PA and "
               "Ext-8D (external supplier 8D).")

    nc = load_nc()

    # ---- burnout split by launcher class (LLV / MLV / SLV) --------------
    st.subheader("CAPA burnout by launcher class")
    st.caption("Each launcher class: how many of its NCs still owe a CAPA "
               "(red) vs how many are covered (green). LLV = Ariane + RS + H3, "
               "MLV = Vega, SLV = Flexline + SAS.")
    burn = (nc.groupby(["Launcher", "CAPA"]).size()
              .reset_index(name="n"))
    # keep a stable class order
    _order = ["LLV", "MLV", "SLV", "Vulcan", "(no project)"]
    burn["Launcher"] = pd.Categorical(burn["Launcher"], categories=_order,
                                      ordered=True)
    burn = burn.sort_values("Launcher")
    figb = px.bar(burn, x="Launcher", y="n", color="CAPA",
                  color_discrete_map=STATUS_COLORS, text="n",
                  category_orders={"Launcher": _order})
    figb.update_layout(barmode="stack", height=320,
                       margin=dict(t=10, l=0, r=0, b=0),
                       xaxis_title="", yaxis_title="NCs",
                       legend_title_text="")
    figb.update_traces(textposition="inside")
    st.plotly_chart(figb, width='stretch')

    # a small coverage table under the burnout
    piv = (nc.pivot_table(index="Launcher", columns="CAPA", values="nc_id",
                          aggfunc="count", fill_value=0)
             .reindex(_order).dropna(how="all"))
    for _col in (OPEN, DONE):
        if _col not in piv.columns:
            piv[_col] = 0
    piv["Total"] = piv[OPEN] + piv[DONE]
    piv["Coverage"] = (100 * piv[DONE] / piv["Total"].replace(0, 1)).round(0)
    piv = piv.rename(columns={OPEN: "Open", DONE: "Done"})
    st.dataframe(piv[["Open", "Done", "Total", "Coverage"]]
                 .style.format({"Coverage": "{:.0f}%"}),
                 width='stretch')

    st.divider()

    # ---- the full cross-filter CAPA view (one self-contained HTML/JS page) --
    # Linked brushing: click the icicle OR either donut and the whole view +
    # the Excel download follow the selection; clear a chip (or Clear all) and
    # it returns to the full view. Rebuilt from the LIVE DB on every run, so it
    # always reflects the latest ingest. This is the same page that will run
    # standalone on the internal server; here it is embedded via components.html.
    rca_df = load_rca_departments()
    components.html(capa_page_html(nc, rca_df), height=2150, scrolling=False)