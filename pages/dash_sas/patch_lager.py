#!/usr/bin/env python3
"""
dash_sas/app/page.html — three additions, no layout changes:
  1. "Lager: hidden" button (default OFF) in the Batches (idx 3) section header.
  2. When OFF, strip the Lager entry from batch_class / batch_copq / batch_defect_heat.
  3. Make those 3 charts taller (CSS min-height on their cards, via data-chart).

Every anchor verified against the real file. Backup + verify-all + idempotent.
    python3 patch_lager_final.py --dry-run
    python3 patch_lager_final.py
"""
import sys, shutil, time
from pathlib import Path
P = Path.home()/"bgtools"/"pages"/"dash_sas"/"app"/"page.html"

def main():
    dry = "--dry-run" in sys.argv
    a=[x for x in sys.argv[1:] if not x.startswith("--")]
    p=Path(a[0]) if a else P
    if not p.is_file(): print(f"NOT FOUND: {p}"); return 2
    t=p.read_text(encoding="utf-8")
    if "includeLager" in t: print("Already patched — nothing to do."); return 0

    E=[]

    # 1 — state + stripper helpers, before PALETTE (PALETTE definitely exists)
    E.append(("helpers", "const PALETTE=",
"""let includeLager=false;
function toggleLager(){
  includeLager=!includeLager;
  const b=document.getElementById('lagerBtn');
  if(b)b.textContent=includeLager?'Lager: shown':'Lager: hidden';
  if(window.LAST_DATA)render(window.LAST_DATA);
}
function _dropStack(d){if(includeLager||!d||!d.rows)return d;
  const k=d.rows.map(r=>r!=='Lager');const s={};
  Object.keys(d.series||{}).forEach(x=>{s[x]=d.series[x].filter((v,i)=>k[i]);});
  return {...d,rows:d.rows.filter((r,i)=>k[i]),series:s};}
function _dropBar(d){if(includeLager||!d||!d.labels)return d;
  const k=d.labels.map(l=>l!=='Lager');
  return {...d,labels:d.labels.filter((l,i)=>k[i]),values:d.values.filter((v,i)=>k[i])};}
function _dropHeat(d){if(includeLager||!d||!d.rows)return d;
  const k=d.rows.map(r=>r!=='Lager');
  return {...d,rows:d.rows.filter((r,i)=>k[i]),z:d.z.filter((r,i)=>k[i])};}
const PALETTE="""))

    # 2 — store data so the toggle can re-render
    E.append(("laststore","function render(d){",
              "function render(d){window.LAST_DATA=d;"))

    # 3 — strip Lager from the three payloads (dashboard + section views share these)
    E.append(("bc","mkStacked(d.batch_class,","mkStacked(_dropStack(d.batch_class),"))
    E.append(("copq","mkBar(d.batch_copq,","mkBar(_dropBar(d.batch_copq),"))
    E.append(("heat","mkHeat(d.batch_defect_heat)","mkHeat(_dropHeat(d.batch_defect_heat))"))

    # 4 — the button, injected into the Batches (idx '3') section header only.
    #     Anchor is the exact sec() innerHTML template line.
    E.append(("button",
        "s.innerHTML=(title?'<div class=\"sec-head\"><span class=\"idx\">'+idx+'</span><h2>'+title+'</h2>'+",
        "s.innerHTML=(title?'<div class=\"sec-head\"><span class=\"idx\">'+idx+'</span><h2>'+title+'</h2>'+"
        "(idx==='3'?'<button id=\"lagerBtn\" class=\"btn small\" style=\"margin-left:12px\" onclick=\"toggleLager()\">Lager: hidden</button>':'')+"))

    # 5 — taller charts via CSS on their cards (card() sets c.dataset.chart=id)
    E.append(("tallcss","</style>",
        "\n[data-chart=\"batch_class\"] .plot,"
        "[data-chart=\"batch_copq\"] .plot,"
        "[data-chart=\"batch_defect_heat\"] .plot{min-height:460px}\n</style>"))

    missing=[n for n,o,_ in E if t.count(o)<1]
    if missing:
        print("ABORTED — anchors not found, nothing written:")
        for n,o,_ in E: print(f"  [{'ok' if t.count(o)>=1 else 'MISSING'}] {n}")
        return 1
    out=t
    for _,o,nw in E: out=out.replace(o,nw,1)
    if dry:
        print(f"DRY RUN ok — all {len(E)} anchors matched. {len(t)}->{len(out)} bytes.")
        return 0
    bak=p.with_suffix(p.suffix+f".bak.{int(time.time())}")
    shutil.copy2(p,bak); p.write_text(out,encoding="utf-8")
    print(f"Patched {p}\nBackup  {bak}\nRestart dash_sas and hard-refresh.")
    return 0
sys.exit(main())