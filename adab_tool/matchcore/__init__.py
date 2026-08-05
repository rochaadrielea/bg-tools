"""matchcore — normalization-safe matching engine (keys + hash + TF-IDF vectors).

Public API:
    from matchcore import norm_key, norm_text, desc_hash, record_hash
    from matchcore import Vectorizer, cosine, build_vectorizer, agree, Agreement
"""
from .normalize import coerce, norm_key, norm_text, tokens
from .hashing import content_hash, desc_hash, record_hash
from .vectorize import Vectorizer, cosine
from .match import build_vectorizer, agree, Agreement

__all__ = ["coerce", "norm_key", "norm_text", "tokens",
           "content_hash", "desc_hash", "record_hash",
           "Vectorizer", "cosine", "build_vectorizer", "agree", "Agreement"]
__version__ = "0.1.0"
