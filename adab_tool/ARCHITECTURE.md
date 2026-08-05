# ADAB Compare — architecture

Built to **grow**. Three layers, each with one job, so a change stays in one place.

```
        ┌──────────────────────────────────────────┐
        │  adab_gui.py         (LAYER 1 — UI)        │
        │  pick inputs, source type, combine, Run    │
        │  knows nothing about how comparison works  │
        └───────────────────┬────────────────────────┘
                            │ calls run_compare(...)
        ┌───────────────────▼────────────────────────┐
        │  adab_batch_compare.py  (LAYER 2 — ENGINE)  │
        │  read files -> match -> reconcile -> write  │
        │  the ADAB rules (positions, quantity, tabs) │
        └───────────────────┬────────────────────────┘
                            │ uses (for "is this the same part?")
        ┌───────────────────▼────────────────────────┐
        │  matchcore/          (LAYER 3 — MATCHING)   │
        │  normalize · hash · vectorize · verdict     │
        │  reusable, tested, no knowledge of ADAB     │
        └──────────────────────────────────────────────┘
```

## Layer 1 — `adab_gui.py` (the window)
Only builds the window and collects inputs. On Run it calls one function:
`core.run_compare(design, built, out, combine=..., built_label=...)`. If you add
a new source type, add it to `SOURCE_TYPES` here — nothing else changes.

## Layer 2 — `adab_batch_compare.py` (the engine)
The comparison rules live here. Key functions, in the order the data flows:
- `resolve_design` / file-or-folder handling → `load_bom` reads the parts sheet
  (`_best_sheet` picks the real list; `_adapt_source` maps Scan/Manual columns
  onto the names the engine expects).
- `split_three` → Matched / Unmatched-built / Unmatched-design at the position level.
- `part_reconciliation` → the auditable DISTINCT-PART headline (Summary + counts).
- `write_summary_tab`, `write_reconciliation_tab`, `write_matched_tab`,
  `write_single_side_tab` → the 5 report tabs. The two unmatched tabs are named
  from `built_label` ("In Design, not in <label>").

**Where it will grow (planned seams):** this one file can later split into
`engine/io.py` (reading + adapters), `engine/match.py` (split_three, reconciliation)
and `engine/report.py` (the xlsx writers) without touching Layer 1 or Layer 3.

## Layer 3 — `matchcore/` (the matching brain)
Decides "are these two rows the same part?" safely. Never depends on ADAB.
- `normalize.py` — the rule: TYPE / CASE / WHITESPACE / ACCENT can never change a
  result (`7004369` == `7004369.0`; "Seal Ring" == "SEAL  RING").
- `hashing.py` — exact fingerprint (SHA-256 of normalized content).
- `vectorize.py` — TF-IDF cosine (near match on descriptions), deterministic.
- `match.py` — `agree(...)` → STRONG / CONFLICT / DESC_ONLY / MATERIAL_ONLY / WEAK.
- `test_invariance.py` — proves the guarantees. Run: `python matchcore/test_invariance.py`.

**Not yet wired into the engine** — matchcore is ready; the next step is to have
`adab_batch_compare` call `matchcore.agree` so the Matched tab also shows a
description verdict and a "Review (Description)" tab lists CONFLICT + DESC_ONLY.

## The rule for growing this
- New **source type** → Layer 1 (`SOURCE_TYPES`) + maybe a small mapping in
  `_adapt_source` (Layer 2). Nothing in Layer 3.
- New **comparison logic** → Layer 2 only.
- New **matching subtlety** (e.g. strip OCR dots, add embeddings) → Layer 3 only,
  behind the same `matchcore` API, so every tool that uses it benefits.
