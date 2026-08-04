"""
normalize.py — the foundation. Every comparison in the engine goes through here
first, so a difference of TYPE, CASE, WHITESPACE, ACCENT or UNICODE FORM can
NEVER change a result. This is the rule Adriele set: 7004369 (text) and
7004369.0 (float) are the same part; "Seal Ring", "SEAL  RING" and "séal ring"
are the same description.
"""
import re
import unicodedata

_WS = re.compile(r"\s+")
_NONWORD = re.compile(r"[^a-z0-9]+")


def coerce(v):
    """Any cell value -> a canonical string, independent of its Python type.
    float 7004369.0 -> '7004369'; NaN/None -> ''; 12 -> '12'."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:                 # NaN
            return ""
        if v.is_integer():         # 7004369.0 -> '7004369'  (Excel float coercion)
            return str(int(v))
        return repr(v)
    if isinstance(v, int):
        return str(v)
    return str(v)


def norm_key(v):
    """Match key for PART / MATERIAL / BATCH numbers.

    Type-invariant, case-invariant, whitespace-collapsed — but SUFFIX-PRESERVING:
    C3529115-C is NOT the same part as C3529115, so we do not strip the suffix.
    Only formatting noise (type, case, spaces) is removed.
    """
    s = coerce(v).strip().upper()
    return _WS.sub(" ", s)


def _strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(c))


def norm_text(v):
    """Normalise FREE TEXT (descriptions) for hashing and vectorising.

    NFKC unicode fold -> strip accents -> casefold -> punctuation to space ->
    collapse whitespace. So 'Ø12 SEAL-RING', 'o12 seal ring' compare equal on
    the parts that matter and never differ by punctuation or case.
    """
    s = unicodedata.normalize("NFKC", coerce(v))
    s = _strip_accents(s).casefold()
    s = _NONWORD.sub(" ", s)
    return _WS.sub(" ", s).strip()


def tokens(v):
    """norm_text split into tokens (for the vectoriser)."""
    t = norm_text(v)
    return t.split() if t else []
