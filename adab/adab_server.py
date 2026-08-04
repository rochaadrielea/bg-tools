#!/usr/bin/env python3
"""ADAB Compare - web app (FastAPI back end + built-in HTML/JS front end).

ONE file. Reuses the SAME engine as the desktop app (adab_batch_compare +
matchcore); nothing about the comparison changes. Drop THIS file into
adab_tool/ next to the engine.

Run on your server:
    pip install fastapi "uvicorn[standard]" python-multipart
    uvicorn adab_server:app --host 0.0.0.0 --port 8000
Then open  http://<server>:8000
"""
import os, uuid, tempfile, traceback
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse, FileResponse, HTMLResponse

import adab_batch_compare as core
import batch_finder

JOBS = os.path.join(tempfile.gettempdir(), "adab_web_jobs")
os.makedirs(JOBS, exist_ok=True)

SOURCE_TYPES = {
    "Label scanner": "Scanner",
    "Manual List": "Manual",
    "Reserved Logistic": "Reserved",
    "mb51": "mb51",
    "Team center": "Teamcenter",
    "Find Batch": "FindBatch",
}

app = FastAPI(title="ADAB Compare")

INDEX_HTML = r"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>ADAB Compare</title>
<style>
:root{--bg:#fff;--panel:#f7f7f9;--line:#e6e5ec;--ink:#1c1a24;--mut:#6d6a7a;--fade:#a6a3b3;
--vio:#5433cf;--viobg:rgba(84,51,207,.09);--ok:#217a54;--okbg:#e7f3ec;--dng:#bd5a54;--dngbg:#fbeceb;--warn:#c1892f;--warnbg:#fdf3e3;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;font-size:14px}
.wrap{max-width:820px;margin:0 auto;padding:22px 22px 60px}
.brand{display:flex;align-items:center;gap:7px;font-weight:800;font-size:15px;color:#14121a}
.brand .dot{width:8px;height:8px;border-radius:50%;background:var(--vio);margin-top:2px}
.brand span{font-weight:400;color:var(--mut);font-size:12px;margin-left:6px}
h1{font-size:22px;margin:10px 0 2px}
.sub{color:var(--mut);font-size:13px;margin-bottom:18px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px;margin-bottom:16px}
label.f{display:block;font-weight:700;font-size:12.5px;margin:12px 0 5px}
label.f:first-child{margin-top:0}
input[type=file],select{width:100%;padding:9px 10px;border:1px solid var(--line);border-radius:8px;background:#fff;font:inherit;font-size:13px}
.hint{font-size:11px;color:var(--fade);margin-top:4px}
.rowc{display:flex;align-items:center;gap:8px;margin-top:12px}
.note{background:var(--viobg);border:1px solid var(--vio);border-radius:8px;padding:9px 11px;font-size:12px;color:var(--vio);margin-top:12px;display:none}
button.run{margin-top:16px;width:100%;background:var(--vio);color:#fff;border:0;border-radius:10px;padding:13px;font-size:15px;font-weight:700;cursor:pointer}
button.run:disabled{opacity:.55;cursor:default}
.out{margin-top:6px}
.sumrow{display:flex;gap:10px;flex-wrap:wrap;margin:6px 0 12px}
.stat{background:#fff;border:1px solid var(--line);border-radius:10px;padding:9px 13px;min-width:96px}
.stat .n{font-size:22px;font-weight:800}
.stat .l{font-size:10.5px;color:var(--mut);margin-top:2px}
.stat.ok .n{color:var(--ok)}.stat.dng .n{color:var(--dng)}.stat.warn .n{color:var(--warn)}.stat.vio .n{color:var(--vio)}
.dl{display:block;background:var(--okbg);border:1px solid var(--ok);color:var(--ok);text-decoration:none;font-weight:700;border-radius:9px;padding:11px 13px;margin-top:8px;font-size:13px}
.dl:hover{filter:brightness(.97)}
pre.log{background:#12121a;color:#d6d6e0;border-radius:10px;padding:12px;font-size:11.5px;white-space:pre-wrap;max-height:260px;overflow:auto;margin-top:12px}
.err{color:var(--dng);font-weight:700}
</style></head><body><div class="wrap">
<div class="brand">beyond gravity<span class="dot"></span><span>ADAB Compare</span></div>
<h1>ADAB Compare</h1>
<div class="sub">As-Design vs As-Built traceability. Upload the baseline and the source, pick the type, Run. Accepts Excel only: .xlsx .xlsm .xls</div>

<div class="card">
  <label class="f">1) As-Design (the F- baseline)</label>
  <input id="design" type="file" accept=".xlsx,.xlsm,.xls">
  <label class="f">2) As-Built source (one or more Excel files)</label>
  <input id="built" type="file" accept=".xlsx,.xlsm,.xls" multiple>
  <div class="hint" id="builthint"></div>
  <label class="f">3) What KIND of As-Built source is this?</label>
  <select id="source"></select>
  <div class="rowc" id="combinerow" style="display:none">
    <input type="checkbox" id="combine"><label for="combine" style="font-size:12.5px">These files are ONE list &mdash; combine into one report</label>
  </div>
  <div class="note" id="fbnote">Find Batch: As-Design = the list MISSING batches (e.g. RED). As-Built = the label/scan list that HAS batches. Fills Charge + Rev and flags FOUND / MULTIPLE / NONE.</div>
  <button class="run" id="runbtn">Run</button>
</div>

<div class="out" id="out"></div>
</div>
<script>
const $=s=>document.querySelector(s);
fetch('/api/source-types').then(r=>r.json()).then(list=>{
  $('#source').innerHTML=list.map(s=>`<option>${s}</option>`).join('');
});
function refresh(){
  const nb=$('#built').files.length;
  $('#builthint').textContent = nb? `${nb} file(s) selected` : '';
  $('#combinerow').style.display = nb>1 ? 'flex':'none';
  $('#fbnote').style.display = $('#source').value==='Find Batch' ? 'block':'none';
}
$('#built').addEventListener('change',refresh);
$('#source').addEventListener('change',refresh);

$('#runbtn').addEventListener('click',async()=>{
  const d=$('#design').files[0], b=$('#built').files;
  const out=$('#out'); out.innerHTML='';
  if(!d||!b.length){ out.innerHTML='<p class="err">Please choose the As-Design and at least one As-Built file.</p>'; return; }
  const fd=new FormData();
  fd.append('source_type',$('#source').value);
  fd.append('combine', $('#combine').checked?'true':'false');
  fd.append('design',d);
  for(const f of b) fd.append('built',f);
  $('#runbtn').disabled=true; $('#runbtn').textContent='Running...';
  try{
    const r=await fetch('/api/run',{method:'POST',body:fd});
    const j=await r.json();
    render(j);
  }catch(e){ out.innerHTML='<p class="err">Network error: '+e+'</p>'; }
  $('#runbtn').disabled=false; $('#runbtn').textContent='Run';
});

function stat(n,l,cls){return `<div class="stat ${cls}"><div class="n">${n}</div><div class="l">${l}</div></div>`;}
function render(j){
  const out=$('#out');
  if(!j.ok){ out.innerHTML='<p class="err">Run failed.</p><pre class="log">'+(j.log||'')+'</pre>'; return; }
  let html='<div class="card">';
  const s=j.summary;
  if(s.mode==='find_batch'){
    html+='<h3 style="margin:0 0 8px">Find Batch result</h3><div class="sumrow">'
      +stat(s.need,'need rows','vio')+stat(s.found,'FOUND','ok')
      +stat(s.multiple,'MULTIPLE','warn')+stat(s.none,'NONE','dng')+'</div>';
  }else{
    html+='<h3 style="margin:0 0 8px">Compare result ('+s.label+')</h3>';
    for(const u of s.units){
      if(u.error){ html+='<p class="err">'+ (u.unit||'') +': '+u.error+'</p>'; continue; }
      html+='<div style="font-weight:700;margin:6px 0 4px">'+(u.unit||'')+'</div><div class="sumrow">'
        +stat(u.parts_matched??'-','matched','ok')
        +stat(u.parts_missing??'-','in design, not in source','dng')
        +stat(u.parts_extra??'-','in source, not in design','warn')
        +stat(u.deviations??'-','deviations','vio')+'</div>';
    }
  }
  for(const f of (j.files||[])) html+='<a class="dl" href="'+f.url+'">Download '+f.name+'</a>';
  html+='<pre class="log">'+(j.log||'')+'</pre></div>';
  out.innerHTML=html;
}
</script></body></html>
"""


@app.get("/", response_class=HTMLResponse)
def index():
    return INDEX_HTML


@app.get("/api/source-types")
def source_types():
    return list(SOURCE_TYPES.keys())


def _save(upload: UploadFile, folder: str) -> str:
    os.makedirs(folder, exist_ok=True)
    p = os.path.join(folder, os.path.basename(upload.filename))
    with open(p, "wb") as f:
        f.write(upload.file.read())
    return p


@app.post("/api/run")
async def run(source_type: str = Form(...),
              combine: bool = Form(False),
              design: UploadFile = File(...),
              built: list[UploadFile] = File(...)):
    job = uuid.uuid4().hex[:12]
    work = os.path.join(JOBS, job)
    out = os.path.join(work, "out"); os.makedirs(out, exist_ok=True)
    log = []
    def prog(m): log.append(str(m))
    try:
        dpath = _save(design, os.path.join(work, "design"))
        bpaths = [_save(u, os.path.join(work, "built")) for u in built]
        built_arg = bpaths[0] if len(bpaths) == 1 else os.path.dirname(bpaths[0])
        if source_type == "Find Batch":
            _, (n, f, m, z) = batch_finder.run(dpath, built_arg, out, progress=prog)
            summary = {"mode": "find_batch", "need": n, "found": f,
                       "multiple": m, "none": z}
        else:
            label = SOURCE_TYPES.get(source_type, "As-Built")
            results = core.run_compare(dpath, built_arg, out, combine=combine,
                                       progress=prog, built_label=label)
            summary = {"mode": "compare", "label": label, "units": list(results)}
    except Exception as e:
        prog("ERROR: " + str(e)); prog(traceback.format_exc())
        return JSONResponse({"ok": False, "log": "\n".join(log)}, status_code=500)
    files = [{"name": fn, "url": "/api/download/" + job + "/" + fn}
             for fn in sorted(os.listdir(out)) if fn.lower().endswith(".xlsx")]
    return {"ok": True, "log": "\n".join(log), "summary": summary, "files": files}


@app.get("/api/download/{job}/{name}")
def download(job: str, name: str):
    path = os.path.join(JOBS, os.path.basename(job), "out", os.path.basename(name))
    if not os.path.exists(path):
        return JSONResponse({"error": "not found"}, status_code=404)
    return FileResponse(path, filename=name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
