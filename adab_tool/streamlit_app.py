#!/usr/bin/env python3
"""ADAB Compare — web app (Streamlit).

Same engine as the desktop tool (adab_batch_compare.py), but reachable by anyone
in the organisation from a browser: no install, no .exe. Upload the As-Design and
the As-Built file(s), click Run, download the report.

Run locally / on a server:
    pip install -r requirements.txt
    streamlit run streamlit_app.py

Files that must sit next to this one:
    adab_batch_compare.py   (the engine)
    matchcore/              (optional; the engine has a built-in fallback)
"""
import os
import glob
import tempfile
import traceback

import streamlit as st
import adab_batch_compare as core

# ------------------------------------------------ Beyond Gravity dark theme ----
st.set_page_config(page_title="ADAB Compare", page_icon="🛰️", layout="centered")
st.markdown("""
<style>
:root { --bg:#081521; --card:#0F2536; --bd:#21455E; --ink:#EAF2F8;
        --sub:#8FA9BC; --acc:#1E9BE0; --accb:#4FC3F7; }
.stApp { background:#081521; color:#EAF2F8; }
h1,h2,h3,h4,label,p,span,div { color:#EAF2F8; }
.bg-head { background:linear-gradient(180deg,#050C14,#0C3A5C);
           border-bottom:2px solid #1E9BE0; border-radius:10px;
           padding:18px 22px; margin-bottom:8px; }
.bg-head h1 { margin:0; font-size:26px; }
.bg-head .dot { color:#4FC3F7; }
.bg-head p { margin:2px 0 0; color:#BBD4E6; font-size:13px; }
.stButton>button, .stDownloadButton>button {
    background:#1E9BE0; color:white; border:0; border-radius:8px;
    font-weight:700; padding:10px 18px; }
.stButton>button:hover, .stDownloadButton>button:hover { background:#1685C4; color:white; }
[data-testid="stFileUploaderDropzone"] { background:#0A1B29; border:1px solid #21455E; }
.stProgress > div > div > div > div { background-color:#1E9BE0; }
code, .stCode { background:#06111C !important; }
</style>
<div class="bg-head">
  <h1>ADAB Compare <span class="dot">●</span></h1>
  <p>As-Design vs As-Built traceability · beyond gravity</p>
</div>
""", unsafe_allow_html=True)


def _milestone(line):
    low = line.strip().lower()
    if low.startswith("as-design"):          return 0.10
    if low.startswith("combine:"):           return 0.15
    if "roles ->" in low:                    return 0.30
    if "match engines" in low:               return 0.45
    if "name engine" in low or "description engine" in low: return 0.70
    if "line conservation" in low:           return 0.82
    if "distinct parts ->" in low:           return 0.96
    if low.startswith(("done", "finished")): return 1.0
    return None


st.write("")
design_up = st.file_uploader("**1 · As-Design (the F- baseline)**",
                             type=["xlsx", "xlsm", "xls"], key="design")
built_up = st.file_uploader("**2 · As-Built source** — one file, or several to combine",
                            type=["xlsx", "xlsm", "xls"], accept_multiple_files=True,
                            key="built")
combine = st.checkbox("Combine all As-Built files into one list (one report)", value=False)
st.caption("Report tabs are always “In As-Built, not in Design” / “In Design, not in As-Built”.")

run = st.button("Run · As-Built vs As-Design", use_container_width=True, type="primary")

if run:
    if design_up is None or not built_up:
        st.error("Please upload the As-Design file and at least one As-Built file.")
        st.stop()

    work = tempfile.mkdtemp(prefix="adab_")
    out_dir = os.path.join(work, "out")
    os.makedirs(out_dir, exist_ok=True)

    # save As-Design
    design_path = os.path.join(work, design_up.name)
    with open(design_path, "wb") as f:
        f.write(design_up.getbuffer())

    # save As-Built (one file -> pass file; many -> pass a folder)
    if len(built_up) == 1 and not combine:
        b0 = os.path.join(work, built_up[0].name)
        with open(b0, "wb") as f:
            f.write(built_up[0].getbuffer())
        built_arg = b0
    else:
        bdir = os.path.join(work, "built")
        os.makedirs(bdir, exist_ok=True)
        for uf in built_up:
            with open(os.path.join(bdir, uf.name), "wb") as f:
                f.write(uf.getbuffer())
        built_arg = bdir

    bar = st.progress(0.0, text="Starting…")
    logbox = st.empty()
    lines = []

    def progress(msg):
        lines.append(str(msg))
        logbox.code("\n".join(lines[-200:]), language=None)
        frac = _milestone(str(msg))
        if frac is not None:
            bar.progress(frac, text=str(msg).strip()[:80] or "Working…")

    try:
        with st.spinner("Comparing…"):
            core.run_compare(design_path, built_arg, out_dir,
                             combine=combine, progress=progress,
                             built_label="As Built")
        bar.progress(1.0, text="Finished ✓")
    except Exception as e:
        st.error(f"Error: {e}")
        st.code(traceback.format_exc())
        st.stop()

    reports = sorted(glob.glob(os.path.join(out_dir, "*.xlsx")))
    if not reports:
        st.warning("No report was produced — check the log above.")
    else:
        st.success(f"Done — {len(reports)} report(s) ready.")
        for rp in reports:
            with open(rp, "rb") as f:
                st.download_button(
                    f"⬇ Download  {os.path.basename(rp)}", f.read(),
                    file_name=os.path.basename(rp),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)