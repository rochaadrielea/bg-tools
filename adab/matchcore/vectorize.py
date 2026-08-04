"""
vectorize.py — deterministic TF-IDF description vectoriser + cosine similarity.

Why TF-IDF and not neural embeddings: this must run offline on the work laptop,
give the SAME number every time (an audit can reproduce it), and add no heavy
dependency or model download. TF-IDF over normalised tokens does exactly the job
here — flagging when two descriptions are the same part worded differently — and
is fully deterministic. (An embedding backend can be added behind the same
Vectorizer interface later if ever needed.)
"""
import math
from collections import Counter
from .normalize import tokens


class Vectorizer:
    """Fit IDF on a corpus (design + built descriptions), then vectorise any text."""

    def __init__(self, corpus):
        docs = [tokens(c) for c in corpus]
        df = Counter()
        for d in docs:
            for t in set(d):
                df[t] += 1
        n = max(1, len(docs))
        # smoothed idf; unseen tokens get the max idf (treated as rare/informative)
        self.idf = {t: math.log((1 + n) / (1 + c)) + 1.0 for t, c in df.items()}
        self._default_idf = math.log((1 + n) / 1) + 1.0

    def vec(self, text):
        tf = Counter(tokens(text))
        return {t: c * self.idf.get(t, self._default_idf) for t, c in tf.items()}


def cosine(a, b):
    """Cosine similarity of two sparse dict vectors, 0.0 .. 1.0."""
    if not a or not b:
        return 0.0
    dot = sum(w * b[t] for t, w in a.items() if t in b)
    na = math.sqrt(sum(w * w for w in a.values()))
    nb = math.sqrt(sum(w * w for w in b.values()))
    return (dot / (na * nb)) if (na and nb) else 0.0
