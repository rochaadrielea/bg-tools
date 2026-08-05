"""
hashing.py — a stable content fingerprint. Two records with the same NORMALISED
content produce the same SHA-256, regardless of type/case/spacing. Used as the
exact "second witness": if the material numbers match AND the description hashes
match, it is the same product with very high confidence.
"""
import hashlib
from .normalize import norm_key, norm_text

_SEP = "␟"   # unit separator — cannot occur in normalised text


def content_hash(*values, text=False):
    """SHA-256 of the normalised values joined by a separator.
    text=False -> norm_key (part numbers, suffix-preserving)
    text=True  -> norm_text (descriptions)."""
    norm = norm_text if text else norm_key
    joined = _SEP.join(norm(v) for v in values)
    return hashlib.sha256(joined.encode("utf-8")).hexdigest()


def desc_hash(desc):
    """Fingerprint of a single description (normalised free text)."""
    return content_hash(desc, text=True)


def record_hash(material, description, batch=None):
    """Identity fingerprint of a part row: material (+optional batch) as keys,
    description as text. Same physical part+batch+description -> same hash."""
    # ONE sha256 over the whole combination (material [+batch] as keys, plus the
    # description as text), so every character of the fingerprint depends on the
    # full content — truncating it (e.g. [:12]) still reflects material AND
    # description. Same normalised combination -> same hash; any change -> change.
    keys = [material] if batch is None else [material, batch]
    payload = _SEP.join(norm_key(k) for k in keys) + _SEP + "text:" + norm_text(description)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()
