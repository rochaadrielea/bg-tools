#!/usr/bin/env python3
"""
Change the Minor/Major per batch chips from two-toggle to THREE exclusive:
  All (default) · Supplier · Production
- Exclusive: only one active at a time.
- Default: All (transparent chips, All highlighted).
- All = full batch_class incl. Unclassified; Supplier/Production = that lane only.
Reverses the previous multi-select and installs the exclusive version.
Backup + verify + idempotent.
"""
import sys, shutil, time
from pathlib import Path
P = Path.home()/"bgtools"/"pages"/"dash_sas"/"app"/"page.html"

def main():
    dry="--dry-run" in sys.argv
    a=[x for x in sys.argv[1:] if not x.startswith("--")]
    p=Path(a[0]) if a else P
    if not p.is_file(): print(f"NOT FOUND: {p}"); return 2
    t=p.read_text(encoding="utf-8")
    if "bcOrigin=" in t: print("Already the exclusive version."); return 0
    if "bcLanes" not in t:
        print("Multi-select version not found — nothing to convert."); return 1

    E=[]

    # 1. Replace the state+builder block. Anchor the whole multi-select block start.
    E.append(("state_block",
        "let bcLanes={supplier:true,production:true};",
        "let bcOrigin='all';  // all | supplier | production (exclusive)"))

    # 2. Replace toggleBcLane with setBcOrigin (exclusive).
    E.append(("toggle_fn",
        """function toggleBcLane(l){
  bcLanes[l]=!bcLanes[l];
  if(!bcLanes.supplier&&!bcLanes.production){bcLanes[l]=true;return;} // keep >=1 on
  document.querySelectorAll('#bcLaneChips .bcchip').forEach(c=>{
    c.classList.toggle('on',bcLanes[c.dataset.l]);});
  if(window.LAST_DATA)render(window.LAST_DATA);
}""",
        """function setBcOrigin(v){
  bcOrigin=v;
  document.querySelectorAll('#bcLaneChips .bcchip').forEach(c=>{
    c.classList.toggle('on',c.dataset.v===v);});
  if(window.LAST_DATA)render(window.LAST_DATA);
}"""))

    # 3. Replace bcBuild — pick the payload by exclusive choice.
    E.append(("build_fn",
        """function bcBuild(d){
  const src=d.batch_class_lanes;
  if(!src)return d.batch_class;   // fallback
  const rows=src.rows;
  const cls=['Minor','Major','Unclassified'];
  const series={};
  cls.forEach(c=>{
    series[c]=rows.map((_,i)=>{
      let n=0;
      if(bcLanes.supplier&&src.supplier[c])n+=src.supplier[c][i]||0;
      if(bcLanes.production&&src.production[c])n+=src.production[c][i]||0;
      return n;});
  });
  return {rows,series,click:'batch'};
}""",
        """function bcBuild(d){
  if(bcOrigin==='all')return d.batch_class;   // all lanes, incl. Unclassified
  const src=d.batch_class_lanes;
  if(!src)return d.batch_class;
  const rows=src.rows;
  const lane=bcOrigin==='supplier'?src.supplier:src.production;
  const series={};
  ['Minor','Major','Unclassified'].forEach(c=>{
    series[c]=rows.map((_,i)=>(lane[c]?lane[c][i]||0:0));});
  return {rows,series,click:'batch'};
}"""))

    # 4. Replace the chip row markup — three chips, All default-on.
    E.append(("chips_markup",
        "'<span class=\"bcchip on\" data-l=\"supplier\" onclick=\"toggleBcLane(\\'supplier\\')\">Supplier</span>'"
        "+'<span class=\"bcchip on\" data-l=\"production\" onclick=\"toggleBcLane(\\'production\\')\">Production</span>'",
        "'<span class=\"bcchip on\" data-v=\"all\" onclick=\"setBcOrigin(\\'all\\')\">All</span>'"
        "+'<span class=\"bcchip\" data-v=\"supplier\" onclick=\"setBcOrigin(\\'supplier\\')\">Supplier</span>'"
        "+'<span class=\"bcchip\" data-v=\"production\" onclick=\"setBcOrigin(\\'production\\')\">Production</span>'"))

    missing=[n for n,o,_ in E if t.count(o)<1]
    if missing:
        print("ABORTED — anchors not found:")
        for n,o,_ in E: print(f"  [{'ok' if t.count(o)>=1 else 'MISSING'}] {n}")
        return 1
    out=t
    for _,o,nw in E: out=out.replace(o,nw,1)
    if dry: print(f"DRY RUN ok — all {len(E)} anchors matched."); return 0
    bak=p.with_suffix(p.suffix+f".bak.{int(time.time())}")
    shutil.copy2(p,bak); p.write_text(out,encoding="utf-8")
    print(f"Patched {p}\nBackup {bak}\nHard-refresh (no restart needed).")
    return 0
sys.exit(main())