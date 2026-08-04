"""
match.py — combine the signals into one verdict (the "two-witness" rule).

A material-number match alone can be a mislabel or an OCR slip. This layer adds
the description as an independent witness:

  material_equal   normalised part numbers are equal (type/case/space-proof)
  desc_hash_equal  normalised descriptions are byte-identical (exact witness)
  desc_similarity  TF-IDF cosine of the descriptions, 0..1 (near witness)

Verdict:
  STRONG        material AND description agree            -> same part, confirmed
  CONFLICT      material equal BUT descriptions disagree  -> INVESTIGATE (wrong
                                                             number or mislabel)
  DESC_ONLY     material differ BUT descriptions match    -> possible same part,
                                                             different/typo'd number
  MATERIAL_ONLY material equal, description missing on a side
  WEAK          neither agrees
"""
from dataclasses import dataclass, asdict

from .normalize import norm_key, norm_text
from .hashing import desc_hash
from .vectorize import Vectorizer, cosine


@dataclass
class Agreement:
    material_equal: bool
    desc_hash_equal: bool
    desc_similarity: float
    verdict: str

    def as_dict(self):
        d = asdict(self)
        d["desc_similarity"] = round(self.desc_similarity, 3)
        return d


def build_vectorizer(*desc_iterables):
    corpus = []
    for it in desc_iterables:
        corpus.extend(it)
    return Vectorizer(corpus)


def agree(mat_a, desc_a, mat_b, desc_b, vec, sim_threshold=0.60):
    """Score one candidate pair. `vec` is a fitted Vectorizer."""
    ka, kb = norm_key(mat_a), norm_key(mat_b)
    material_equal = bool(ka) and ka == kb

    ta, tb = norm_text(desc_a), norm_text(desc_b)
    both_desc = bool(ta) and bool(tb)
    hash_equal = both_desc and desc_hash(desc_a) == desc_hash(desc_b)
    sim = cosine(vec.vec(desc_a), vec.vec(desc_b)) if both_desc else 0.0
    desc_agrees = hash_equal or sim >= sim_threshold

    if material_equal and desc_agrees:
        verdict = "STRONG"
    elif material_equal and not both_desc:
        verdict = "MATERIAL_ONLY"
    elif material_equal and not desc_agrees:
        verdict = "CONFLICT"
    elif not material_equal and desc_agrees:
        verdict = "DESC_ONLY"
    else:
        verdict = "WEAK"

    return Agreement(material_equal, hash_equal, sim, verdict)
