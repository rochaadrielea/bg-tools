"""
app.py - web front end for the OpCenter <-> Teamcenter join.

The command-line tool reads two files from a folder. This page takes the same
two files as browser uploads instead, so nothing has to be copied onto the
server first. The joining logic is NOT duplicated here: it is imported from
nc_opcenter_tc.py, which stays the single place the rules live.

Run:
    cd ~/bgtools/nc_opcenter_tc
    ~/bgtools/dash/quality/bin/streamlit run app.py \
        --server.address 0.0.0.0 --server.port 8503 \
        --server.baseUrlPath opcenter --server.maxUploadSize 200
"""
import io
import getpass
import json
import socket
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from nc_opcenter_tc import (OPC_COLS, TC_COLS, build, read_opcenter,
                            read_teamcenter)

# Feedback goes to the same table as the dashboard and ADAB, so every tool's
# feedback lands in one place rather than a file only this page reads.
FEEDBACK_DB = Path.home() / "bgtools" / "dash" / "quality.db"
APP_NAME = "NC OpCenter/Teamcenter join"
APP_VERSION = "web-2026-08"

st.set_page_config(page_title="OpCenter + Teamcenter", layout="wide")


def feedback_widget():
    with st.expander("Improve this tool - send the team your idea", expanded=False):
        with st.form("opc_feedback", clear_on_submit=True):
            c1, c2 = st.columns(2)
            name = c1.text_input("Your name (optional)")
            cat = c2.selectbox("Type", ["Idea / improvement", "Something's wrong",
                                        "Question", "Other"])
            msg = st.text_area("What would make this better?")
            if st.form_submit_button("Send") and msg.strip():
                try:
                    con = sqlite3.connect(FEEDBACK_DB)
                    con.execute(
                        "INSERT INTO feedback (ts, app, version, category, "
                        "message, user, machine, context) VALUES (?,?,?,?,?,?,?,?)",
                        (datetime.now().isoformat(timespec="seconds"),
                         APP_NAME, APP_VERSION, cat, msg.strip(),
                         getpass.getuser(), socket.gethostname(),
                         json.dumps({"submitted_by": name.strip() or "(anonymous)"})))
                    con.commit()
                    con.close()
                    st.success("Thank you! Sent to the team.")
                except Exception as e:
                    st.error(f"Could not save: {e}")


def to_excel(hit, miss, tc_only, sources):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        hit.to_excel(w, sheet_name="Matched", index=False)
        miss.to_excel(w, sheet_name="OpCenter only", index=False)
        tc_only.to_excel(w, sheet_name="Teamcenter only", index=False)
        sources.to_excel(w, sheet_name="Sources", index=False)
        for nm, df in [("Matched", hit), ("OpCenter only", miss),
                       ("Teamcenter only", tc_only), ("Sources", sources)]:
            ws = w.sheets[nm]
            ws.freeze_panes = "A2"
            for i, col in enumerate(df.columns, start=1):
                wid = max([len(str(col))] +
                          [len(str(v)) for v in df[col].head(200)]) if len(df) else len(str(col))
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(46, wid + 2)
    return buf.getvalue()


feedback_widget()
st.title("OpCenter + Teamcenter")
st.caption(
    "Joins the OpCenter nonconformance export to the Teamcenter issue export. "
    "The key is exact: OpCenter's **Identifier** is the same `NC_` id that sits "
    "inside Teamcenter's **Object** cell, so no name or fuzzy matching is used. "
    "The join is what turns an OpCenter account code into a person's name and "
    "attaches the IR number.")

c1, c2 = st.columns(2)
opc_file = c1.file_uploader(
    "OpCenter export (.csv, .xlsx, .xlsm)", type=["csv", "xlsx", "xlsm"],
    help="NonConformance export from OpCenter. The raw CSV, or the same file "
         "re-saved in Excel - both work. It just needs an 'Identifier' column.")
tc_file = c2.file_uploader(
    "Teamcenter export (.xlsm / .xlsx)", type=["xlsm", "xlsx"],
    help="The timestamped export with an 'Object' column of IR- values. "
         "Use the newest one - older snapshots have fewer filled fields.")

if not (opc_file and tc_file):
    st.info("Upload both files to build the report.")
    st.stop()

try:
    opc = read_opcenter(opc_file)
except SystemExit as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"Could not read the OpCenter export: {e}\n\n"
             "It needs an **Identifier** column. CSV (any separator) and Excel "
             "both work - check you uploaded the OpCenter file and not the "
             "Teamcenter one.")
    st.stop()

try:
    tc = read_teamcenter(tc_file)
except Exception as e:
    st.error(f"Could not read the Teamcenter export: {e}")
    st.stop()

if tc.empty:
    st.error("No IR- rows found in the Teamcenter file. Check it is the right export.")
    st.stop()

hit, miss, tc_only = build(opc, tc)

sources = pd.DataFrame([
    {"Role": "OpCenter", "File": opc_file.name, "Rows": len(opc)},
    {"Role": "Teamcenter", "File": tc_file.name, "Rows": len(tc)},
    {"Role": "Report generated", "File": "",
     "Rows": datetime.now().strftime("%Y-%m-%d %H:%M")},
])

m1, m2, m3, m4 = st.columns(4)
m1.metric("Matched", len(hit), help="OpCenter NCs that have a Teamcenter issue.")
m2.metric("OpCenter only", len(miss),
          help="Nonconformances raised in production with NO Teamcenter issue. "
               "Usually the ones worth looking at.")
m3.metric("Teamcenter only", len(tc_only),
          help="Issues with no matching OpCenter row - free-text issues, or an "
               "NC id absent from this export.")
m4.metric("OpCenter rows", len(opc))

st.download_button(
    "Download Excel (4 tabs)", to_excel(hit, miss, tc_only, sources),
    file_name=f"nc_opcenter_tc_{datetime.now():%Y-%m-%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

t1, t2, t3, t4, t5 = st.tabs(
    [f"Matched ({len(hit)})", f"OpCenter only ({len(miss)})",
     f"Teamcenter only ({len(tc_only)})", f"All Teamcenter ({len(tc)})",
     "Field coverage"])

with t1:
    st.dataframe(hit, width="stretch", hide_index=True)

with t2:
    st.caption("A nonconformance recorded in production that never became a "
               "Teamcenter issue. Check whether it should have.")
    st.dataframe(miss, width="stretch", hide_index=True)

with t3:
    st.caption("Split by reason - a free-text issue carries no NC id at all, so "
               "it cannot match by construction.")
    st.dataframe(tc_only, width="stretch", hide_index=True)

with t4:
    st.caption("Every row of the Teamcenter export, with the `Object` cell "
               "split into its parts. The letter code in front of a "
               "description is the same Detection taxonomy the NC tracker "
               "uses, so a row with no NC id is still classifiable by where "
               "it was found.")
    st.dataframe(tc, width="stretch", hide_index=True)
    if tc["Detection"].notna().any():
        det = (tc["Detection"].value_counts().rename_axis("Detection")
               .reset_index(name="Issues"))
        st.caption("Detection areas found in the export:")
        st.dataframe(det, width="stretch", hide_index=True)

with t5:
    st.caption("How often each Teamcenter field is actually filled on the "
               "matched rows. A low number means the field is empty in "
               "Teamcenter, not that the join failed.")
    if len(hit):
        cov = pd.DataFrame([
            {"Field": c,
             "Filled": int(hit[c].notna().sum()),
             "Of": len(hit),
             "%": round(100 * hit[c].notna().sum() / len(hit))}
            for c in TC_COLS])
        st.dataframe(cov, width="stretch", hide_index=True)
    st.dataframe(sources, width="stretch", hide_index=True)