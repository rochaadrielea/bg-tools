#!/usr/bin/env python3
"""ADAB Compare — web app (Streamlit), Beyond Gravity dark theme.

Two pages:
  1 · Build ABCL — compare As-Design vs As-Built (same engine as the desktop
      tool, adab_batch_compare.py). Every run is now auto-recorded to a history
      database so it can be looked up later.
  2 · Research  — search that history: type/scan a material (or batch, document,
      serial) and see every past comparison it appeared in, with the exact
      As-Design / As-Built / report files from each run to download.

Recording is ADDITIVE (adab_history.py): it only reads the report the engine
already writes and stores the rows. If recording ever fails, the comparison
itself is untouched.

Run on the server:
    ~/bgtools/dash/quality/bin/streamlit run streamlit_app.py \
        --server.address 0.0.0.0 --server.port 8502 --server.baseUrlPath adab
"""
import os
import re
import io
import glob
import hashlib
import zipfile
import tempfile
import datetime
import traceback

import streamlit as st
import streamlit.components.v1 as components
import adab_batch_compare as core

# comparison memory (Research). Additive: only reads the report + stores rows.
try:
    import adab_history as history
    _HIST_OK = True
except Exception as _h_err:
    _HIST_OK = False
    _HIST_IMPORT_ERR = str(_h_err)

# Where the history lives — next to this app, so it persists on the server
# (override with ADAB_HISTORY_DIR). runs + files land in <here>/adab_data/.
HISTORY_BASE = os.environ.get(
    "ADAB_HISTORY_DIR",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "adab_data"))

# shared feedback store (same DB the desktop tools write to — Adriele's standing
# rule: every UI tool has a feedback button that saves to the quality database).
APP_VERSION = "web-2026-08"
try:
    from feedback import submit_feedback, CATEGORIES, default_db_path
    _FB_OK = True
except Exception as _fb_err:          # feedback.py missing on the server, etc.
    _FB_OK = False
    _FB_IMPORT_ERR = str(_fb_err)
    CATEGORIES = ["Bug / something broke", "Idea / improvement",
                  "Wrong result", "Question", "Other"]
    def default_db_path():
        return os.path.join(os.path.expanduser("~"), ".beyondgravity", "feedback.db")


def _dedupe(files):
    """Count a file only once. Detects the SAME file even if it was renamed, by
    hashing its content — so an identical upload isn't merged twice. Returns
    (unique_files, skipped_names)."""
    seen, uniq, dups = set(), [], []
    for f in files or []:
        try:
            h = hashlib.md5(f.getvalue()).hexdigest()
        except Exception:
            h = f.name           # fall back to name if content can't be read
        if h in seen:
            dups.append(f.name)
            continue
        seen.add(h)
        uniq.append(f)
    return uniq, dups


def _is_adab_output(f):
    """True if an uploaded file is a report ADAB itself produced (so it must NOT
    be used as the As-Design). Detected by our own naming convention."""
    n = (getattr(f, "name", "") or "").lower()
    return (n.startswith("adab_") or n.startswith("report complete")
            or "missing items" in n)


def _save_attachments(files, who):
    """Save feedback attachments (screenshots / docs) next to the feedback DB,
    return the saved file names (recorded in the entry's context)."""
    if not files:
        return []
    import datetime
    saved = []
    try:
        base = os.path.dirname(os.path.abspath(default_db_path()))
        adir = os.path.join(base, "feedback_attachments")
        os.makedirs(adir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_who = re.sub(r"[^A-Za-z0-9_-]+", "_", (who or "anon"))[:20] or "anon"
        for i, f in enumerate(files):
            safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", f.name)[-60:]
            fn = f"{stamp}_{safe_who}_{i+1}_{safe_name}"
            with open(os.path.join(adir, fn), "wb") as out:
                out.write(f.getbuffer())
            saved.append(fn)
    except Exception as e:
        saved.append(f"(attachment save failed: {type(e).__name__}: {e})")
    return saved

st.set_page_config(page_title="ADAB Compare", page_icon="🛰️", layout="centered")

st.markdown("""
<style>
:root { --acc:#1E9BE0; --accb:#4FC3F7; }
.stApp { background:#081521; color:#FFFFFF; }
/* hide Streamlit's white top toolbar / header / rainbow bar */
header[data-testid="stHeader"], [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"] { display:none !important; }
#MainMenu, footer { visibility:hidden; }

/* EVERY letter white */
.stApp, .stApp p, .stApp label, .stApp span, .stApp div, .stApp li,
h1,h2,h3,h4,h5,h6, [data-testid="stMarkdownContainer"] * { color:#FFFFFF !important; }
.block-container { max-width:1000px; padding-top:1.0rem; padding-bottom:2rem; }

.bg-head { background:linear-gradient(180deg,#050C14 0%,#0C3A5C 100%);
           border-bottom:3px solid #1E9BE0; border-radius:14px;
           padding:24px 28px; margin-bottom:18px; }
.bg-head h1 { margin:0; font-size:38px; font-weight:800; }
.bg-head .dot { color:#4FC3F7 !important; }
.bg-head p { margin:6px 0 0; font-size:15px; }

.sec { font-size:19px; font-weight:700; color:#4FC3F7 !important; margin:0 0 2px; }
.hint { font-size:13.5px; margin-bottom:4px; }

div[data-testid="stVerticalBlockBorderWrapper"]{
    background:#0F2536; border:1px solid #21455E !important;
    border-radius:14px; padding:12px 18px 16px; margin-bottom:14px; }

[data-testid="stFileUploaderDropzone"]{ background:#0A1B29; border:1px dashed #2C5875; padding:20px; }
[data-testid="stFileUploaderDropzone"] *{ color:#FFFFFF !important; }
[data-testid="stFileUploaderDropzone"] button{
    background:#1E9BE0 !important; color:#fff !important; border:0 !important;
    font-weight:700 !important; padding:8px 22px !important; }
/* Style Streamlit's per-file chips dark (instead of the ugly white pill) and
   KEEP them visible — so a file that fails to upload shows its red error icon
   here and the user can see exactly which one to re-add (Adriele). */
[data-testid="stFileUploaderFile"]{ background:#0A1B29 !important;
    border:1px solid #2C5875 !important; border-radius:8px !important;
    margin-bottom:4px !important; color:#FFFFFF !important; }
.okmsg{ color:#5AD08A !important; font-weight:700; font-size:14px; margin-top:8px; }

.stButton>button, .stDownloadButton>button{
    background:#1E9BE0; color:#fff !important; border:0; border-radius:10px;
    font-weight:800; font-size:18px; padding:15px 22px; width:100%; }
.stButton>button:hover, .stDownloadButton>button:hover{ background:#1685C4; }
.stDownloadButton>button{ font-size:16px; padding:12px 18px; }

/* button state while the engine works / when it finishes */
.run-state{ border-radius:10px; font-weight:800; font-size:18px; text-align:center;
            padding:15px 22px; width:100%; box-sizing:border-box; color:#fff !important;
            letter-spacing:.4px; }
/* WORKING: bright blue moving gradient with a glow (aerospace-serious) */
.run-state.busy{
    background:linear-gradient(90deg,#0C3A5C,#1E9BE0,#7FD4FF,#1E9BE0,#0C3A5C);
    background-size:220% 100%; border:1px solid #7FD4FF;
    box-shadow:0 0 20px rgba(127,212,255,.55);
    animation:bgshine 7s ease-in-out infinite; }
@keyframes bgshine { 0%{background-position:220% 0;} 100%{background-position:-220% 0;} }
/* DONE / calm confirmation */
.run-state.done{ background:linear-gradient(90deg,#0E7A38,#17A34A); }
/* little white spinner shown inside the working button */
.spin{ display:inline-block; width:15px; height:15px; margin-right:11px;
       border:3px solid rgba(255,255,255,.35); border-top-color:#fff;
       border-radius:50%; vertical-align:-2px; animation:spin 1.6s linear infinite; }
@keyframes spin{ to{ transform:rotate(360deg);} }

/* 3 · Report — make the download area stand out (Adriele: people should look here) */
.report-drop{ border:2px dashed #1E9BE0; border-radius:12px; background:#0A1B29;
    padding:30px 20px; text-align:center; font-size:16px; font-weight:600;
    color:#BFE3F7 !important; }
.report-drop .big{ display:block; font-size:30px; margin-bottom:6px; color:#4FC3F7 !important; }
.report-drop b{ color:#FFFFFF !important; }
.report-ready{ border:2px solid #4FC3F7; border-radius:12px; background:#0C2A3E;
    padding:14px 16px; box-shadow:0 0 16px rgba(79,195,247,.35); margin-bottom:8px; }
/* BIG red banner when a file fails to upload / process (Adriele) */
.bigwarn{ background:#3A0E12; border:2px solid #E15554; border-radius:12px;
    padding:18px 20px; margin-bottom:14px; box-shadow:0 0 16px rgba(225,85,84,.35); }
.bigwarn .h{ color:#FF6B6B !important; font-size:20px; font-weight:800; margin-bottom:6px; }
.bigwarn .l{ color:#FFD9D9 !important; font-size:15px; font-weight:600; }
/* amber caution shown by the uploaders */
.caution{ background:#2A2410; border:1px solid #C99A2E; border-radius:10px;
    padding:11px 15px; margin-top:8px; }
.caution .h{ color:#FFE9A8 !important; font-size:14px; font-weight:800; margin-bottom:3px; }
.caution .l{ color:#F0DCA0 !important; font-size:13px; }
.caution b{ color:#FFFFFF !important; }

/* feedback box fields — dark, white text */
.stTextArea textarea, .stTextInput input{
    background:#0A1B29 !important; color:#FFFFFF !important; border:1px solid #2C5875 !important; }
/* the "start writing" prompt text — readable, not near-invisible grey */
.stTextArea textarea::placeholder, .stTextInput input::placeholder{
    color:#8FB6D0 !important; opacity:1 !important; font-style:italic; }
[data-baseweb="select"] > div{ background:#0A1B29 !important; border:1px solid #2C5875 !important; }
[data-baseweb="select"] *{ color:#FFFFFF !important; }

.stProgress > div > div > div > div { background-color:#1E9BE0; }

/* the friendly log box — dark blue, white letters */
pre, code { background:#06111C !important; color:#FFFFFF !important; font-size:14px;
            border:1px solid #21455E; border-radius:8px; line-height:1.5; }
[data-testid="stCheckbox"] label p, [data-testid="stMultiSelect"] * { color:#FFFFFF !important; }

/* ---- top page nav (Build ABCL | Research) styled like tabs ---- */
div[role="radiogroup"].adabnav { flex-direction:row; gap:8px; margin-bottom:16px; }
div[role="radiogroup"] > label{ background:#0F2536; border:1px solid #21455E;
    border-bottom:none; border-radius:10px 10px 0 0; padding:9px 20px; margin:0;
    color:#8FB6D0 !important; font-weight:700; }

/* ---- Research page ---- */
.rz-mat{ font-size:22px; font-weight:800; color:#7FD4FF !important;
    font-family:Consolas,monospace; }
.rz-desc{ color:#9Fc0d6 !important; font-size:14px; }
.chip{ display:inline-block; background:#0F2536; border:1px solid #21455E;
    border-radius:10px; padding:9px 14px; margin:0 8px 8px 0; min-width:118px; }
.chip .k{ color:#8FB6D0 !important; font-size:11px; text-transform:uppercase; letter-spacing:.4px; }
.chip .v{ font-size:19px; font-weight:800; margin-top:2px; }
.chip .v.green{ color:#3FBF7A !important; } .chip .v.blue{ color:#7FD4FF !important; }
.rz-hdr{ display:flex; gap:10px; color:#8FB6D0 !important; font-size:11px;
    text-transform:uppercase; letter-spacing:.4px; border-bottom:1px solid #21455E;
    padding:6px 4px; font-weight:700; }
.rz-cell{ font-size:13px; }
.rz-mono{ font-family:Consolas,monospace; }
.stpill{ display:inline-block; padding:2px 9px; border-radius:999px; font-size:12px; font-weight:700; }
.stpill.m{ color:#3FBF7A !important; border:1px solid rgba(63,191,122,.5); background:rgba(63,191,122,.10); }
.stpill.d{ color:#D9A441 !important; border:1px solid rgba(217,164,65,.5); background:rgba(217,164,65,.10); }
.stpill.b{ color:#5FA8DE !important; border:1px solid rgba(95,168,222,.5); background:rgba(95,168,222,.10); }
.rz-note{ color:#9FC0D6 !important; font-size:12.5px; margin-top:10px; line-height:1.6; }
.rz-run{ color:#CFE6F5 !important; font-size:13px; }
</style>
<div class="bg-head">
  <h1>ADAB Compare <span class="dot">●</span></h1>
  <p>As-Design vs As-Built traceability&nbsp;·&nbsp;beyond gravity</p>
</div>
""", unsafe_allow_html=True)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

_CNT = re.compile(r"\[(\d+)/(\d+)\]")
_READ = re.compile(r"read (\d+)/(\d+)")


def _fraction(line):
    """Progress 0..1 from an engine log line."""
    low = line.strip().lower()
    m = _CNT.search(low)
    if m:
        return int(m.group(1)) / max(1, int(m.group(2)))
    m = _READ.search(low)
    if m:
        return 0.15 + 0.45 * (int(m.group(1)) / max(1, int(m.group(2))))
    if low.startswith("as-design"):          return 0.10
    if low.startswith("combine:"):           return 0.15
    if "roles ->" in low:                    return 0.30
    if "match engines" in low:               return 0.45
    if "name engine" in low or "description engine" in low: return 0.70
    if "line conservation" in low:           return 0.82
    if "distinct parts ->" in low:           return 0.96
    if low.startswith(("done", "finished")): return 1.0
    return None


def _friendly(line):
    """Turn a technical engine log line into a plain message for the user.
    Returns None for lines that should stay hidden."""
    m = line.strip()
    low = m.lower()
    r = _READ.search(low)
    if r:
        return f"Reading your files… {r.group(1)} of {r.group(2)}"
    c = _CNT.search(low)
    if c:
        return f"✓ Report {c.group(1)} of {c.group(2)} finished"
    if low.startswith("as-designed"):
        return "Reading your As-Design file…"
    if low.startswith("combine:"):
        return "Combining your As-Built files into one…"
    if low.startswith("comparing") and "parallel" in low:
        return "Comparing your files (all at once)…"
    if "roles ->" in low:
        return "Checking which file is the design and which is the as-built…"
    if "match engines" in low:
        return "Matching part numbers and names…"
    if "name engine" in low:
        return "Looking for parts with the same name but a different number…"
    if "line conservation" in low:
        return "Keeping every line — nothing is dropped ✓"
    if "distinct parts ->" in low:
        mm = re.search(r"matched (\d+).*?missing (\d+).*?extra (\d+)", low)
        if mm:
            return (f"Matched {mm.group(1)} parts  ·  {mm.group(2)} only in your "
                    f"design  ·  {mm.group(3)} only in the as-built")
    if low.startswith(("done", "finished")):
        return "Finished! Your report is ready below. ✓"
    if "!!!" in m or "warning" in low or low.startswith("skipped") or "error" in low:
        return "⚠ " + m
    return None


# --------------------------------------------------------------------------- #
#  history recording (called after every successful run — never breaks a run)
# --------------------------------------------------------------------------- #
def _record_history(design_arg, built_saved_paths, out_dir, run_label, combine):
    """Store every report the engine just wrote into the comparison history.
    Fully guarded: any failure here is swallowed so the comparison is untouched.
    Returns (n_recorded, note_or_None)."""
    if not _HIST_OK:
        return 0, "history module not available"
    try:
        reports = [p for p in sorted(glob.glob(os.path.join(out_dir, "*.xlsx")))
                   if not os.path.basename(p).upper().startswith("MISSING ITEMS")]
        if not reports:
            return 0, None
        design_list = ([design_arg] if os.path.isfile(design_arg)
                       else sorted(glob.glob(os.path.join(design_arg, "*"))))
        n = 0
        if combine or len(reports) == 1:
            bpath = built_saved_paths[0] if built_saved_paths else None
            for rp in reports:
                if history.record_run(design_list, bpath, rp, built_label="As Built",
                                      run_label=run_label, base_dir=HISTORY_BASE):
                    n += 1
        else:
            # map each "report complete - <base>.xlsx" back to its As-Built file
            bmap = {core._clean_base(os.path.splitext(os.path.basename(b))[0]): b
                    for b in built_saved_paths}
            for rp in reports:
                base = os.path.splitext(os.path.basename(rp))[0]
                bb = base.split(" - ", 1)[1] if " - " in base else base
                bpath = bmap.get(core._clean_base(bb)) or (
                    built_saved_paths[0] if built_saved_paths else None)
                if history.record_run(design_list, bpath, rp, built_label="As Built",
                                      run_label=run_label, base_dir=HISTORY_BASE):
                    n += 1
        return n, None
    except Exception as e:
        return 0, f"{type(e).__name__}: {e}"


# =========================================================================== #
#  PAGE 1 — BUILD ABCL
# =========================================================================== #
def render_build():
    # --- inputs -------------------------------------------------------------
    with st.container(border=True):
        st.markdown('<div class="sec">1 · As-Design (the F- baseline)</div>', unsafe_allow_html=True)
        st.markdown('<div class="hint">Click <b>Browse files</b> — one MBOM, or select several at '
                    'once (Ctrl+A). Several files are merged into one baseline. '
                    '<b>Put here what you treat as the design / baseline</b> — usually your MBOM, '
                    'but any file (even a previous report) is accepted; you choose the role.</div>',
                    unsafe_allow_html=True)
        design_up_raw = st.file_uploader("As-Design", type=["xlsx", "xlsm", "xls", "csv"],
                                         accept_multiple_files=True, key="design",
                                         label_visibility="collapsed")
        design_up, design_dups = _dedupe(design_up_raw)
        if design_dups:
            st.markdown(f'<div class="hint" style="color:#E0A423 !important;">↺ Skipped '
                        f'{len(design_dups)} duplicate file(s) — same content as one already '
                        f'added, so counted once: {", ".join(design_dups)}</div>',
                        unsafe_allow_html=True)
        if design_up:
            if len(design_up) == 1:
                st.markdown(f'<div class="okmsg">✓ Uploaded: {design_up[0].name}</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="okmsg">✓ Uploaded {len(design_up)} files (merged): '
                            f'{", ".join(f.name for f in design_up)}</div>',
                            unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown('<div class="sec">2 · As-Built source</div>', unsafe_allow_html=True)
        st.markdown('<div class="hint">Click <b>Browse files</b> — pick one file, or select several at '
                    'once (open the folder and choose all the files). Tick Combine to treat them as one list. '
                    '<b>Put here what was actually built</b> — a list, scan, SAP / mb51 export, or even a '
                    'previous report; any file is accepted.</div>',
                    unsafe_allow_html=True)
        built_up_raw = st.file_uploader("As-Built", type=["xlsx", "xlsm", "xls", "csv"],
                                        accept_multiple_files=True, key="built",
                                        label_visibility="collapsed")
        built_up, built_dups = _dedupe(built_up_raw)
        if built_dups:
            st.markdown(f'<div class="hint" style="color:#E0A423 !important;">↺ Skipped '
                        f'{len(built_dups)} duplicate file(s) — same content as one already '
                        f'added, so counted once: {", ".join(built_dups)}</div>',
                        unsafe_allow_html=True)
        if built_up:
            if len(built_up) == 1:
                st.markdown(f'<div class="okmsg">✓ Uploaded: {built_up[0].name}</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="okmsg">✓ Uploaded {len(built_up)} files: '
                            f'{", ".join(f.name for f in built_up)}</div>',
                            unsafe_allow_html=True)
        combine = st.checkbox("Combine all As-Built files into one list (one report)", value=False)

    # Detector: watch the uploaders and show a red bar ONLY when a file fails to
    # upload, naming it (Adriele).
    components.html("""
    <script>
    (function(){
      try{
        var doc = window.parent.document;
        var BID = "adab-upload-fail-banner";
        function looksRed(el){
          try{
            var s = window.parent.getComputedStyle(el);
            function red(v){ var m=v&&v.match(/rgba?\\((\\d+),\\s*(\\d+),\\s*(\\d+)/);
              if(!m) return false; return (+m[1])>170 && (+m[2])<110 && (+m[3])<110; }
            return red(s.color) || red(s.fill);
          }catch(e){ return false; }
        }
        function scan(){
          try{
            var chips = doc.querySelectorAll('[data-testid="stFileUploaderFile"]');
            var failed = [];
            chips.forEach(function(ch){
              var err = false;
              if (ch.querySelector('[data-testid$="ErrorMessage"], [role="alert"]')) err = true;
              if (!err){ var els=ch.querySelectorAll('*');
                for (var i=0;i<els.length;i++){ if(looksRed(els[i])){ err=true; break; } } }
              if (err){
                var nm = ch.querySelector('[data-testid="stFileUploaderFileName"]');
                var t = nm ? nm.textContent : (ch.textContent||"a file");
                failed.push(t.replace(/\\s*\\d+(\\.\\d+)?\\s*(K|M|G)?B\\s*$/i,"").trim().slice(0,70));
              }
            });
            var b = doc.getElementById(BID);
            if (failed.length){
              if(!b){ b=doc.createElement("div"); b.id=BID;
                b.style.cssText="position:fixed;top:0;left:0;right:0;z-index:99999;"
                  +"background:#B4232A;color:#fff;padding:14px 22px;font-family:Arial,sans-serif;"
                  +"font-size:16px;font-weight:700;box-shadow:0 2px 12px rgba(0,0,0,.45);";
                doc.body.appendChild(b); }
              b.innerHTML = "⚠ These file(s) did NOT upload: "
                + failed.map(function(n){return "<u>"+n+"</u>";}).join(",  ")
                + " &nbsp;—&nbsp; remove them (✗) and add again (close them in Excel first).";
            } else if (b){ b.remove(); }
          }catch(e){}
        }
        setInterval(scan, 900); scan();
      }catch(e){}
    })();
    </script>
    """, height=0)

    # 2b · a short nickname for this comparison (used as the run label in Research)
    with st.container(border=True):
        st.markdown('<div class="sec">3 · Short name for this comparison</div>',
                    unsafe_allow_html=True)
        st.markdown('<div class="hint">A quick nickname so you can find this run later in '
                    '<b>Research</b> — e.g. <i>“A6 C6001L”</i>, <i>“Vulcan dome F02”</i>. '
                    'Optional; if you leave it blank we use the As-Built file name.</div>',
                    unsafe_allow_html=True)
        run_label = st.text_input("Short name", key="run_label",
                                  placeholder="e.g. A6 C6001L  ·  Vulcan dome F02",
                                  label_visibility="collapsed")

    st.markdown('<div class="hint" style="margin:2px 2px 8px;">Report tabs are always '
                '“In As-Built, not in Design” / “In Design, not in As-Built”. '
                'Every run is saved to <b>Research</b> automatically.</div>',
                unsafe_allow_html=True)

    btn_ph = st.empty()
    run = btn_ph.button("Run · As-Built vs As-Design", type="primary", key="runbtn")

    # progress + log always visible
    st.write("")
    c1, c2 = st.columns([4, 1])
    status_ph = c1.empty()
    pct_ph = c2.empty()
    bar = st.progress(0.0)
    status_ph.markdown('<div class="sec" style="font-size:15px;color:#FFFFFF !important;">Ready.</div>',
                       unsafe_allow_html=True)
    pct_ph.markdown('<div style="text-align:right;font-weight:800;color:#4FC3F7;font-size:17px;">0%</div>',
                    unsafe_allow_html=True)

    st.markdown('<div class="hint" style="margin-top:10px;">What’s happening</div>', unsafe_allow_html=True)
    log_ph = st.empty()
    log_ph.code("Upload your files and press Run — I’ll tell you each step here.", language=None)

    # save context for the feedback box (works even after reruns)
    st.session_state["_ctx"] = {
        "as_design_files": len(design_up) if design_up else 0,
        "as_built_files": len(built_up) if built_up else 0,
        "combine": bool(combine),
    }

    # --- run ----------------------------------------------------------------
    if run:
        if not design_up or not built_up:
            st.error("Please upload the As-Design file(s) and at least one As-Built file.")
            st.stop()

        btn_ph.markdown('<div class="run-state busy"><span class="spin"></span>'
                        'Running the engine…</div>', unsafe_allow_html=True)

        work = tempfile.mkdtemp(prefix="adab_")
        out_dir = os.path.join(work, "out")
        os.makedirs(out_dir, exist_ok=True)

        # As-Design: one file, or several merged into one baseline (pass a folder)
        if len(design_up) == 1:
            design_arg = os.path.join(work, design_up[0].name)
            with open(design_arg, "wb") as f:
                f.write(design_up[0].getbuffer())
        else:
            ddir = os.path.join(work, "design")
            os.makedirs(ddir, exist_ok=True)
            for uf in design_up:
                with open(os.path.join(ddir, uf.name), "wb") as f:
                    f.write(uf.getbuffer())
            design_arg = ddir

        built_saved = []
        if len(built_up) == 1 and not combine:
            b0 = os.path.join(work, built_up[0].name)
            with open(b0, "wb") as f:
                f.write(built_up[0].getbuffer())
            built_arg = b0
            built_saved = [b0]
        else:
            bdir = os.path.join(work, "built")
            os.makedirs(bdir, exist_ok=True)
            for uf in built_up:
                p = os.path.join(bdir, uf.name)
                with open(p, "wb") as f:
                    f.write(uf.getbuffer())
                built_saved.append(p)
            built_arg = bdir

        shown = []
        maxf = [0.0]

        def progress(msg):
            nice = _friendly(str(msg))
            if nice:
                shown.append(nice)
                log_ph.code("\n".join(shown[-200:]), language=None)
            frac = _fraction(str(msg))
            if frac is not None:
                maxf[0] = max(maxf[0], frac)
                bar.progress(maxf[0])
                if nice:
                    status_ph.markdown(
                        f'<div class="sec" style="font-size:15px;color:#FFFFFF !important;">{nice}</div>',
                        unsafe_allow_html=True)
                pct_ph.markdown(
                    f'<div style="text-align:right;font-weight:800;color:#4FC3F7;font-size:17px;">{int(maxf[0]*100)}%</div>',
                    unsafe_allow_html=True)

        try:
            results = core.run_compare(design_arg, built_arg, out_dir, combine=combine,
                                       progress=progress, built_label="As Built",
                                       workers=os.cpu_count())
            bar.progress(1.0)
            pct_ph.markdown('<div style="text-align:right;font-weight:800;color:#4FC3F7;font-size:17px;">100%</div>',
                            unsafe_allow_html=True)
            btn_ph.markdown('<div class="run-state done">✓ Finished — report is ready below</div>',
                            unsafe_allow_html=True)
        except Exception as e:
            btn_ph.markdown('<div class="run-state done" style="background:#B4232A;">✕ Stopped — see the message below</div>',
                            unsafe_allow_html=True)
            st.error(f"Something went wrong: {e}")
            st.code(traceback.format_exc())
            st.stop()

        # --- record every run to the comparison history (Research) ----------
        # Done BEFORE we rename the files for download, so it copies the real
        # report the engine wrote. Never allowed to break the run.
        n_rec, hist_note = _record_history(design_arg, built_saved, out_dir,
                                           run_label.strip(), combine)
        st.session_state["hist_saved"] = n_rec
        st.session_state["hist_note"] = hist_note

        # read the reports into memory so they survive the download clicks (reruns).
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %Hh%Mm%Ss")
        reps = []
        for rp in sorted(glob.glob(os.path.join(out_dir, "*.xlsx"))):
            base, ext = os.path.splitext(os.path.basename(rp))
            with open(rp, "rb") as f:
                reps.append({"name": f"{base} {stamp}{ext}", "data": f.read()})
        st.session_state["reports"] = reps
        st.session_state["run_id"] = st.session_state.get("run_id", 0) + 1

        # collect any file that did NOT produce a report, so we can shout about it
        problems = []
        for r in (results or []):
            if not isinstance(r, dict):
                continue
            if r.get("error"):
                problems.append(f"{r.get('unit', 'a file')} — {r['error']}")
            elif r.get("empty"):
                problems.append(f"{r.get('unit', 'a file')} — no As-Built data found (empty file)")
        full_reports = [r for r in reps
                        if not r["name"].upper().startswith("MISSING ITEMS")]
        if not combine and len(built_up) > len(full_reports) and not problems:
            problems.append(f"Only {len(full_reports)} of {len(built_up)} files produced "
                            "a report — one or more could not be read.")
        st.session_state["problems"] = problems

    # --- report / download (persists across download clicks) ----------------
    with st.container(border=True):
        st.markdown('<div class="sec">4 · Report</div>', unsafe_allow_html=True)
        _probs = st.session_state.get("problems")
        if _probs:
            items = "".join(f'<div class="l">• {msg}</div>' for msg in _probs)
            st.markdown('<div class="bigwarn"><div class="h">⚠ Some files did NOT make it '
                        'into a report</div>' + items +
                        '<div class="l" style="margin-top:8px;">Re-add the file(s) above '
                        'and run again. If you saw a red “Network Error” when uploading, '
                        'that file didn’t reach the server — refresh (Ctrl+Shift+R) and '
                        'add it again.</div></div>', unsafe_allow_html=True)
        reps = st.session_state.get("reports")
        rid = st.session_state.get("run_id", 0)   # fresh widget keys per run
        # small confirmation that the run was saved to Research
        if reps and st.session_state.get("hist_saved"):
            st.markdown(f'<div class="okmsg">🔎 Saved to Research — '
                        f'{st.session_state["hist_saved"]} run(s) recorded. Look it up '
                        f'any time on the Research page.</div>', unsafe_allow_html=True)
        elif reps and st.session_state.get("hist_note"):
            st.markdown(f'<div class="hint" style="color:#E0A423 !important;">Note: this run '
                        f'was not saved to Research ({st.session_state["hist_note"]}). The '
                        f'report above is unaffected.</div>', unsafe_allow_html=True)
        if not reps:
            st.markdown('<div class="report-drop"><span class="big">⬇</span>'
                        'Your Excel report will appear <b>here</b> to download '
                        'after you press <b>Run</b>.</div>', unsafe_allow_html=True)
        elif len(reps) == 1:
            st.success("Done — your report is ready.")
            st.download_button(f"⬇  Download  {reps[0]['name']}", reps[0]["data"],
                               file_name=reps[0]["name"], mime=XLSX_MIME,
                               key=f"dl_single_{rid}")
        else:
            st.success(f"Done — {len(reps)} reports ready. Tick the ones you want, then download.")
            names = [r["name"] for r in reps]
            chosen = st.multiselect("Choose reports", names, default=names,
                                    label_visibility="collapsed", key=f"choose_{rid}")
            if chosen:
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
                    for r in reps:
                        if r["name"] in chosen:
                            z.writestr(r["name"], r["data"])
                st.download_button(f"⬇  Download selected ({len(chosen)}) as ZIP",
                                   buf.getvalue(), file_name="adab_reports.zip",
                                   mime="application/zip", key=f"dl_zip_{rid}")
            with st.expander("…or download one at a time"):
                for r in reps:
                    st.download_button(f"⬇  {r['name']}", r["data"],
                                       file_name=r["name"], mime=XLSX_MIME,
                                       key=f"one_{rid}_{r['name']}")


# =========================================================================== #
#  PAGE 2 — RESEARCH
# =========================================================================== #
_FIELD_CHOICES = {
    "Material number": "material",
    "Batch / Charge": "batch",
    "Document / project": "document",
    "Serial / Equipment": "serial",
}
_STATUS_CLASS = {"Matched": "m", "In Design, not built": "d", "In Built, not design": "b"}


def _status_pill(status):
    s = status or ""
    cls = "m"
    if s.startswith("In Design"):
        cls = "d"
    elif s.startswith("In Built"):
        cls = "b"
    elif s.startswith("Matched"):
        cls = "m"
    return f'<span class="stpill {cls}">{s}</span>'


def _fmt_date(iso):
    if not iso:
        return "—"
    try:
        return datetime.datetime.fromisoformat(iso).strftime("%d/%m/%Y")
    except Exception:
        return str(iso)[:10]


def _dl_from_disk(path, label, key):
    """A download button that reads a stored history file from disk on demand."""
    if path and os.path.exists(path):
        try:
            with open(path, "rb") as f:
                data = f.read()
            st.download_button(label, data, file_name=os.path.basename(path),
                               key=key, mime=XLSX_MIME)
            return
        except Exception:
            pass
    st.markdown(f'<div class="hint" style="opacity:.5;">{label} —</div>',
                unsafe_allow_html=True)


def render_research():
    if not _HIST_OK:
        st.error("The Research history module isn’t available on the server "
                 f"(adab_history.py). Details: {_HIST_IMPORT_ERR}")
        return

    db = history.db_path_for(HISTORY_BASE)

    with st.container(border=True):
        st.markdown('<div class="sec">Research — search the comparison history</div>',
                    unsafe_allow_html=True)
        st.markdown('<div class="hint">Type or scan a <b>material number</b> (or switch to '
                    'batch, document, serial). ADAB looks it up across <b>every comparison '
                    'ever run</b> and shows when it was checked, against which documents, the '
                    'result, and the files from that run.</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([3, 1.2])
        text = c1.text_input("Search", key="rz_text",
                             placeholder="e.g. N38432174  ·  CH0405  ·  A6 C6001L",
                             label_visibility="collapsed")
        field_label = c2.selectbox("Search by", list(_FIELD_CHOICES.keys()),
                                   key="rz_field", label_visibility="collapsed")
    field = _FIELD_CHOICES[field_label]

    if not text or not text.strip():
        # show what's in the memory so the page isn't empty
        try:
            import sqlite3
            if os.path.exists(db):
                con = sqlite3.connect(db)
                nr = con.execute("SELECT COUNT(*) FROM runs").fetchone()[0]
                nm = con.execute("SELECT COUNT(*) FROM run_materials").fetchone()[0]
                last = con.execute("SELECT run_label, ts FROM runs ORDER BY ts DESC "
                                   "LIMIT 5").fetchall()
                con.close()
                st.markdown(f'<div class="rz-note">The memory holds <b>{nr}</b> comparison '
                            f'run(s) and <b>{nm}</b> recorded material rows. '
                            f'Search above to look one up.</div>', unsafe_allow_html=True)
                if last:
                    rows = "".join(f'<div class="rz-run">• {(rl or "—")} '
                                   f'<span style="color:#8FB6D0;">({_fmt_date(ts)})</span></div>'
                                   for rl, ts in last)
                    st.markdown('<div class="hint" style="margin-top:8px;">Most recent runs:</div>'
                                + rows, unsafe_allow_html=True)
            else:
                st.markdown('<div class="rz-note">No comparisons recorded yet — run one on the '
                            '<b>Build ABCL</b> page and it will appear here.</div>',
                            unsafe_allow_html=True)
        except Exception:
            pass
        return

    q = text.strip()

    # material search: headline chips first
    if field == "material":
        summ = history.material_summary(db, q)
        if not summ.get("times"):
            st.markdown(f'<div class="rz-note">No record of <b>{q}</b> in any comparison yet. '
                        f'Try a partial number, or switch the search type.</div>',
                        unsafe_allow_html=True)
            return
        st.markdown(f'<div style="margin:4px 0 10px;"><span class="rz-mat">{summ["material"]}</span>'
                    f'&nbsp;&nbsp;<span class="rz-desc">{summ.get("description") or ""}</span></div>',
                    unsafe_allow_html=True)
        latest = summ.get("latest_status", "")
        chips = (
            f'<div class="chip"><div class="k">Times seen</div><div class="v blue">{summ["times"]}</div></div>'
            f'<div class="chip"><div class="k">Distinct documents</div><div class="v">{summ["documents"]}</div></div>'
            f'<div class="chip"><div class="k">First recorded</div><div class="v" style="font-size:15px">{_fmt_date(summ["first"])}</div></div>'
            f'<div class="chip"><div class="k">Last recorded</div><div class="v" style="font-size:15px">{_fmt_date(summ["last"])}</div></div>'
            f'<div class="chip"><div class="k">Latest status</div><div class="v {"green" if latest.startswith("Matched") else ""}" style="font-size:15px">{latest}</div></div>'
        )
        st.markdown(f'<div style="margin-bottom:6px;">{chips}</div>', unsafe_allow_html=True)
        rows = summ["rows"]
    else:
        rows = history.search(db, q, field)
        if not rows:
            st.markdown(f'<div class="rz-note">No comparison history matched <b>{q}</b> '
                        f'by {field_label.lower()}. Try a partial value or another search type.</div>',
                        unsafe_allow_html=True)
            return
        st.markdown(f'<div class="rz-note">Found <b>{len(rows)}</b> record(s) matching '
                    f'<b>{q}</b> by {field_label.lower()}, newest first.</div>',
                    unsafe_allow_html=True)

    # results — one line per record, newest first, with the run's 3 files
    with st.container(border=True):
        head = st.columns([1.3, 3.2, 2.2, 0.8, 1.6, 0.7, 3.2])
        for col, name in zip(head, ["Date", "Comparison (As-Design ↔ As-Built)",
                                    "Status", "Rev", "Batch / Charge", "Qty", "Files"]):
            col.markdown(f'<div class="rz-hdr" style="border:none;padding:2px;">{name}</div>',
                         unsafe_allow_html=True)
        for i, r in enumerate(rows[:300]):
            cols = st.columns([1.3, 3.2, 2.2, 0.8, 1.6, 0.7, 3.2])
            comp = f'{r.get("run_label") or r.get("built_name") or "—"}'
            design_name = r.get("design_name") or "As-Design"
            cols[0].markdown(f'<div class="rz-cell rz-mono">{_fmt_date(r.get("ts"))}</div>',
                             unsafe_allow_html=True)
            cols[1].markdown(f'<div class="rz-cell">{design_name} ↔ {comp}<br>'
                             f'<span style="color:#8FB6D0;font-size:11px;" class="rz-mono">'
                             f'{r.get("material_raw","")}</span></div>', unsafe_allow_html=True)
            cols[2].markdown(f'<div class="rz-cell">{_status_pill(r.get("status"))}</div>',
                             unsafe_allow_html=True)
            cols[3].markdown(f'<div class="rz-cell rz-mono">{r.get("revision") or "—"}</div>',
                             unsafe_allow_html=True)
            cols[4].markdown(f'<div class="rz-cell rz-mono">{r.get("batch") or "—"}</div>',
                             unsafe_allow_html=True)
            bq = r.get("built_qty")
            bq = "" if bq is None else (int(bq) if float(bq).is_integer() else bq)
            cols[5].markdown(f'<div class="rz-cell rz-mono">{bq if bq != "" else "—"}</div>',
                             unsafe_allow_html=True)
            with cols[6]:
                fc = st.columns(3)
                with fc[0]:
                    _dl_from_disk(r.get("design_file"), "Design", f"d_{i}_{r.get('run_id')}")
                with fc[1]:
                    _dl_from_disk(r.get("built_file"), "Built", f"b_{i}_{r.get('run_id')}")
                with fc[2]:
                    _dl_from_disk(r.get("report_file"), "Match", f"m_{i}_{r.get('run_id')}")

    if len(rows) > 300:
        st.markdown(f'<div class="hint">Showing the newest 300 of {len(rows)} records — '
                    f'narrow the search to see more.</div>', unsafe_allow_html=True)


# =========================================================================== #
#  FEEDBACK (shown on both pages — standing rule: every tool has it)
# =========================================================================== #
def render_feedback():
    st.write("")
    if "show_fb" not in st.session_state:
        st.session_state["show_fb"] = False
    fbcol, _fbsp = st.columns([2, 3])
    with fbcol:
        if st.button("💬  FEEDBACK?  CLICK HERE", key="fb_toggle"):
            st.session_state["show_fb"] = not st.session_state["show_fb"]

    if st.session_state["show_fb"]:
        with st.container(border=True):
            st.markdown('<div class="hint" style="margin-bottom:8px;">Spot a bug, '
                        'have an idea, or something worked really well? Tell us — '
                        'good and bad, it all goes straight to the quality team. '
                        'Add your name if you’d like — it’s optional.</div>',
                        unsafe_allow_html=True)
            if st.session_state.pop("fb_flash", None):
                st.success(st.session_state.pop("fb_flash_msg", "Thank you — your "
                                                "feedback was saved. ✓"))
            nz = st.session_state.get("fb_nonce", 0)
            fb_name = st.text_input("Your name (optional)", key=f"fb_name_{nz}",
                                    placeholder="Your name — optional",
                                    label_visibility="collapsed")
            FB_CATEGORIES = list(CATEGORIES)
            if not any(("praise" in c.lower() or "worked well" in c.lower())
                       for c in FB_CATEGORIES):
                FB_CATEGORIES.insert(1, "Praise / what worked well 👍")
            fb_cat = st.selectbox("Type", FB_CATEGORIES, key=f"fb_cat_{nz}")
            fb_msg = st.text_area("Your feedback", key=f"fb_msg_{nz}", height=95,
                                  placeholder="Tell us what happened, what worked well, "
                                              "a wrong result, or your idea…",
                                  label_visibility="collapsed")
            st.markdown('<div class="hint" style="margin:2px 0 4px;">Attach a screenshot '
                        'or a file (optional) — helps us see exactly what you mean.</div>',
                        unsafe_allow_html=True)
            fb_files = st.file_uploader("Attach files", accept_multiple_files=True,
                                        key=f"fb_files_{nz}", label_visibility="collapsed")
            bcol, _sp = st.columns([1, 3])
            with bcol:
                send_fb = st.button("Send", key=f"fb_send_{nz}")
            if send_fb:
                if not fb_msg.strip():
                    st.warning("Please type your feedback before sending.")
                elif not _FB_OK:
                    st.error("Feedback module isn’t available on the server "
                             f"(feedback.py). Details: {_FB_IMPORT_ERR}")
                else:
                    try:
                        who = fb_name.strip()
                        attached = _save_attachments(fb_files, who)
                        _c = st.session_state.get("_ctx", {})
                        ctx = {
                            "submitted_by": who or "(anonymous)",
                            "attachments": attached,
                            "as_design_files": _c.get("as_design_files", 0),
                            "as_built_files": _c.get("as_built_files", 0),
                            "combine": _c.get("combine", False),
                            "reports_ready": len(st.session_state.get("reports") or []),
                            "page": st.session_state.get("adab_page", "Build ABCL"),
                        }
                        submit_feedback("ADAB Compare (web)", APP_VERSION,
                                        fb_cat, fb_msg.strip(), context=ctx)
                        extra = (f" ({len(attached)} file(s) attached)"
                                 if attached else "")
                        st.session_state["fb_flash"] = True
                        st.session_state["fb_flash_msg"] = (
                            f"Thank you{', ' + who if who else ''} — your feedback "
                            f"was saved{extra}. ✓")
                        st.session_state["fb_nonce"] = nz + 1
                        st.rerun()
                    except Exception as e:
                        st.error(f"Could not save your feedback: {e}")


# =========================================================================== #
#  NAV + DISPATCH
# =========================================================================== #
page = st.radio("Page", ["1 · Build ABCL", "2 · Research"],
                horizontal=True, label_visibility="collapsed", key="adab_page_radio")
st.session_state["adab_page"] = page

if page.startswith("2"):
    render_research()
else:
    render_build()

render_feedback()