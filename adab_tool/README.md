# ADAB Compare — As-Design vs As-Built traceability

Compares an **As-Design** parts list against an **As-Built source** and writes an
Excel report you can hand to a reviewer. The As-Built source can be any of:
**Scan** (label scanner), **Manual list**, **Reserved** (SAP), **mb51** (SAP
movements), or **Teamcenter As-Built**.

## Run it
```
run_adab.bat
```
or, with the (quality) env active:
```
python adab_gui.py
```

## In the app
1. **As-Design** — pick the F- baseline (Browse File *or* Browse Folder).
2. **As-Built source** — pick a file or a folder.
3. **Source type** — Scan / Manual / Reserved / mb51 / Teamcenter. This names the
   report tabs so the direction is obvious.
4. **Combine** — tick when several files together form ONE list.
5. **Run** — reports land in the Output folder.

## The report (5 tabs)
| Tab | What it shows |
|---|---|
| **Summary** | Distinct-part counts that reconcile by hand (design / built / matched / missing / extra). |
| **Parts Reconciliation** | One row per distinct part: design qty, built qty, positions, status (MATCHED / SHORT / OVER / MISSING / EXTRA). |
| **Matched** | The matched pairs with the design + built detail and deviations. |
| **In Design, not in \<source\>** | Design parts absent from that source (e.g. "In Design, not in Reserved"). |
| **In \<source\>, not in Design** | Source parts absent from the design (e.g. "In Reserved, not in Design"). |

The last two tab names change with the source type you picked — that's the
contextual naming.

## Folder layout (see ARCHITECTURE.md for the why)
```
adab_tool/
├─ adab_gui.py             UI only (entry point)
├─ adab_batch_compare.py   the comparison engine
├─ matchcore/              reusable matching engine (normalize + hash + vector)
├─ feedback.py             the shared "improve this tool" button
├─ run_adab.bat            launcher
├─ ADAB_Compare.spec       PyInstaller build spec
├─ README.md               this file
└─ ARCHITECTURE.md         how the pieces fit + where to add things
```
