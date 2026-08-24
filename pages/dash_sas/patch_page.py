#!/usr/bin/env python3
"""
Add the document SOURCE to dash_sas/page.html.

Four edits, all additive:
  1. RAWDEF.cols     + 'source_file'      (the static "10 oldest open" table)
  2. RAWDEF.labels   + source_file:'Source'
  3. dqPanel(...)    receives d.dq_by_file
  4. dqPanel renders a "By document" table next to "By source system"

Nothing is written unless ALL FOUR match exactly. A timestamped backup is made
first. Run it as many times as you like — it detects an already-patched file and
stops.

    python3 patch_page.py                 # patch
    python3 patch_page.py --dry-run       # show what would change, write nothing
"""
from __future__ import annotations
import argparse
import shutil
import sys
import time
from pathlib import Path

DEFAULT = Path.home() / "bgtools" / "pages" / "dash_sas" / "page.html"

EDITS = [
    (
        "RAWDEF.cols",
        "  'vendor_clean','copq'],",
        "  'vendor_clean','copq','source_file'],",
    ),
    (
        "RAWDEF.labels",
        "  disposition:'Disposition',vendor_clean:'Vendor',copq:'CoPQ'}};",
        "  disposition:'Disposition',vendor_clean:'Vendor',copq:'CoPQ',\n"
        "  source_file:'Source'}};",
    ),
    (
        "dqPanel call",
        "  b.appendChild(dqPanel(d.dq,d.dq_by_source));",
        "  b.appendChild(dqPanel(d.dq,d.dq_by_source,d.dq_by_file));",
    ),
    (
        "dqPanel signature",
        "function dqPanel(dq,bySource){",
        "function dqPanel(dq,bySource,byFile){",
    ),
    (
        "By document table",
        "        '</tr>').join('')+'</table></div>';\n"
        "  }\n"
        "  s.innerHTML=html;",
        "        '</tr>').join('')+'</table></div>';\n"
        "  }\n"
        "  // per DOCUMENT: which uploaded file each row came from. Separate from\n"
        "  // 'By source system' above — that one is SAP vs Teamcenter, this one is\n"
        "  // the file. Rows loaded before source_file was recorded show as blank.\n"
        "  if(byFile&&Object.keys(byFile).length){\n"
        "    const files=Object.keys(byFile).sort();\n"
        "    html+='<div class=\"card\" style=\"margin-top:12px\"><h3>By document</h3>'+\n"
        "      '<p class=\"hint\">The file each NC was loaded from. When the same NC is '+\n"
        "      'loaded again, the newer file replaces the name.</p>'+\n"
        "      '<table><tr><th>Gap</th>'+\n"
        "      files.map(x=>'<th>'+x+'</th>').join('')+'</tr>'+\n"
        "      DQ_FIELDS.map(([k,l])=>'<tr><td>'+l+'</td>'+\n"
        "        files.map(x=>{const g=byFile[x];const v=g[k]||0;\n"
        "          const pct=g.rows?Math.round(100*v/g.rows):0;\n"
        "          const col=pct>=80?'var(--red)':(pct>=30?'var(--amber)':'var(--muted)');\n"
        "          return '<td style=\"color:'+col+'\">'+v+' of '+g.rows+\n"
        "                 ' <span class=\"small\">('+pct+'%)</span></td>';}).join('')+\n"
        "        '</tr>').join('')+'</table></div>';\n"
        "  }\n"
        "  s.innerHTML=html;",
    ),
]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("path", nargs="?", default=str(DEFAULT))
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    p = Path(args.path).expanduser()
    if not p.is_file():
        print(f"NOT FOUND: {p}")
        return 2

    text = p.read_text(encoding="utf-8")

    if "source_file:'Source'" in text:
        print("Already patched — source_file is present. Nothing to do.")
        return 0

    missing = [name for name, old, _ in EDITS if text.count(old) != 1]
    if missing:
        print("ABORTED — nothing was written.")
        for name, old, _ in EDITS:
            n = text.count(old)
            mark = "ok" if n == 1 else f"FOUND {n} TIMES"
            print(f"  [{mark}] {name}")
        print("\nThe file differs from what this patch expects. Send the file and")
        print("it will be patched by hand instead of guessing.")
        return 1

    out = text
    for _, old, new in EDITS:
        out = out.replace(old, new, 1)

    if args.dry_run:
        print("DRY RUN — all 5 edits matched. Nothing written.")
        print(f"  size {len(text)} -> {len(out)} bytes (+{len(out) - len(text)})")
        return 0

    bak = p.with_suffix(p.suffix + f".bak.{int(time.time())}")
    shutil.copy2(p, bak)
    p.write_text(out, encoding="utf-8")
    print(f"Patched {p}")
    print(f"Backup  {bak}")
    print(f"  size {len(text)} -> {len(out)} bytes")
    print("\nRestart the dash_sas service, then hard-refresh (Ctrl+Shift+R).")
    print("To undo:  cp '" + str(bak) + "' '" + str(p) + "'")
    return 0


if __name__ == "__main__":
    sys.exit(main())