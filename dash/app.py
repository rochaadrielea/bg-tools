"""
app.py — Quality BRM Dashboard (entry point).

Three top tabs:

    [ Quality Dashboard ]  [ Customer Complaints ]  [ CAPA ]

  * Quality Dashboard — rendered from `dashboard_body.py`.
  * Customer Complaints — rendered by `customer_complaints.render()`.
  * CAPA — rendered by `capa_view.render()`.

IMPORTANT — why the body is exec'd rather than pasted here.
This file used to hold a full inline COPY of the dashboard body indented under
`with tab_dash:`. That meant `dashboard_body.py` sat on disk doing nothing:
editing it changed the file but not the page, and the only way to change the
dashboard was to edit this file instead. The two copies drifted apart.

There is now exactly one copy of the dashboard, in `dashboard_body.py`, and this
file reads it. `exec` is used rather than `import` because the body is a script,
not a module: it runs top to bottom on every rerun, which is what Streamlit
needs. `with` is not a Python scope, so every `st.*` call inside the body still
renders into the Dashboard tab, and its `with st.sidebar:` block still reaches
the sidebar (the sidebar is tab-independent).

To change the dashboard, edit `dashboard_body.py`. Never paste it back in here.

Run:
    streamlit run app.py
"""
from pathlib import Path

import streamlit as st

import capa_view
import customer_complaints

# The password gate is disabled. To switch it back on, uncomment both lines.
# from auth import check_password

st.set_page_config(
    page_title="BU Launchers - Quality BRM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# if not check_password():
#     st.stop()

tab_dash, tab_cc, tab_capa = st.tabs(
    ["Quality Dashboard", "Customer Complaints", "CAPA"])

# Make the top tab labels large, bold and high-contrast so they are always visible.
# Streamlit 1.60 renders each tab as div[data-testid="stTab"][role="tab"]
# (react-aria) with the label in an inner <p> -- NOT <button>/data-baseweb.
# Target that structure directly.
st.markdown("""
    <style>
    /* Streamlit's top header bar is fixed, opaque white and ~60px tall
       (z-index 999990). The dashboard's reduced .block-container top padding
       pulled the tab row UP underneath it, hiding the labels behind the header.
       Restore enough top padding so the whole tab row clears the header. */
    .block-container,
    [data-testid="stMainBlockContainer"] {
        padding-top: 4.5rem !important;
    }
    .stTabs [role="tablist"] {
        gap: 6px;
        border-bottom: 3px solid #1E2761;
        margin-bottom: 0.4rem;
    }
    .stTabs [data-testid="stTab"] {
        padding: 10px 24px;
        background: #EEF1F7;
        border-radius: 8px 8px 0 0;
    }
    .stTabs [data-testid="stTab"] p {
        font-size: 1.15rem !important;
        font-weight: 800 !important;
        color: #1E2761 !important;
    }
    .stTabs [data-testid="stTab"][aria-selected="true"] {
        background: #1E2761;
    }
    .stTabs [data-testid="stTab"][aria-selected="true"] p {
        color: #FFFFFF !important;
    }
    /* The Customer Complaints tab is the high-severity one — flag it in red
       when it is the active tab so it always reads as the "danger" page. */
    .stTabs [data-testid="stTab"]:nth-of-type(2)[aria-selected="true"] {
        background: #E53E3E;
    }
    .stTabs .react-aria-SelectionIndicator {
        background: #F26E21 !important;
    }
    </style>
""", unsafe_allow_html=True)

_BODY = Path(__file__).parent / "dashboard_body.py"

with tab_dash:
    if not _BODY.exists():
        st.error(f"`{_BODY.name}` not found next to app.py — the dashboard cannot render.")
    else:
        exec(compile(_BODY.read_text(encoding="utf-8"), str(_BODY), "exec"), globals())

with tab_cc:
    customer_complaints.render()

with tab_capa:
    capa_view.render()