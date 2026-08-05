# matchcore — normalization-safe matching engine

A small, dependency-free engine for deciding whether two part rows describe the
**same product**, using three independent witnesses and a normalization layer
that can never be fooled by type, case, whitespace, accents or unicode form.

## Why it exists
On the ADAB compare, matching on the material number alone can be wrong: a float
`7004369.0` vs text `7004369`, a mislabelled ticket, or an OCR slip. This engine
adds the **description** as a second witness and guarantees formatting can never
change a result.

## The rule (non-negotiable)
Everything goes through `normalize` first:
- `norm_key(v)` — part/material/batch numbers. Type- and case-invariant,
  whitespace-collapsed, **suffix-preserving** (`C3529115-C` ≠ `C3529115`).
- `norm_text(v)` — descriptions. NFKC + accent-strip + casefold + punctuation→space.
So `7004369` / `7004369.0` / `" 7004369 "` are one key; `Seal Ring` / `SEAL  RING`
/ `séal ring` are one description.

## Modules
- `normalize.py` — coerce / norm_key / norm_text / tokens  (the foundation)
- `hashing.py`   — content_hash / desc_hash / record_hash   (exact witness, SHA-256)
- `vectorize.py` — Vectorizer (TF-IDF) + cosine             (near witness, deterministic)
- `match.py`     — agree(...) -> Agreement                  (combine into a verdict)

## Verdicts (two-witness logic)
| verdict | meaning |
|---|---|
| STRONG | material AND description agree → same part, confirmed |
| CONFLICT | material equal BUT descriptions disagree → **investigate** (wrong number / mislabel) |
| DESC_ONLY | material differ BUT descriptions match → possible same part, different/typo number |
| MATERIAL_ONLY | material equal, a description is missing |
| WEAK | neither agrees |

## Usage
```python
from matchcore import norm_key, desc_hash, build_vectorizer, agree

vec = build_vectorizer(design_descriptions, built_descriptions)  # fit once
a = agree(design_id, design_desc, built_part, built_desc, vec)
print(a.verdict, a.desc_similarity)
```

## Guarantees, tested (`test_invariance.py`)
Type / case / whitespace / accent / unicode never change a key, a hash, or a
verdict. Run: `python matchcore/test_invariance.py`.

Deterministic by design: TF-IDF (not neural embeddings) so an audit reproduces
the exact number offline, with no model download or API. An embedding backend
can be slotted behind the same `Vectorizer` interface later if ever needed.
