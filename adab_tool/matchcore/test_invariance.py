"""Proves the hard requirement: TYPE / CASE / WHITESPACE / ACCENT / UNICODE
must NEVER change a key, a hash, or a match verdict."""
import os, sys; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from matchcore.normalize import norm_key, norm_text
from matchcore.hashing import desc_hash, record_hash
from matchcore.vectorize import Vectorizer, cosine
from matchcore.match import build_vectorizer, agree

ok = True
def check(name, cond):
    global ok
    ok = ok and cond
    print(("PASS" if cond else "FAIL"), name)

# 1) KEY invariance: float vs text vs spaced vs cased  -> identical
variants = ["7004369", 7004369, 7004369.0, " 7004369 ", "7004369\n", "  7004369"]
keys = {norm_key(v) for v in variants}
check("material 7004369: text/int/float/spaces -> ONE key", keys == {"7004369"})

# but a real suffix must NOT collapse (different parts stay different)
check("suffix preserved: C3529115-C != C3529115",
      norm_key("c3529115-c") != norm_key("C3529115"))
check("suffix key case-insensitive", norm_key("c3529115-c") == norm_key("C3529115-C"))

# 2) DESCRIPTION hash invariance: case / double-space / accent / unicode
d = ["Seal Ring", "SEAL  RING", "seal   ring", "séal ring", "Seal\tRing"]
hs = {desc_hash(x) for x in d}
check("description hash: case/space/accent -> ONE hash", len(hs) == 1)
check("different description -> different hash",
      desc_hash("Seal Ring") != desc_hash("Pyro Holder"))

# 3) record_hash: material float vs text + description case -> same identity
check("record_hash type+case invariant",
      record_hash(7004369.0, "SEAL RING") == record_hash("7004369", "seal ring"))

# 4) VECTOR similarity: identical text = 1.0, reworded high, unrelated low
vec = Vectorizer(["retained gravity screw m6", "seal ring viton",
                  "pyro holder bracket", "housing gravity subassembly"])
check("cosine identical = 1.0",
      abs(cosine(vec.vec("SEAL RING VITON"), vec.vec("seal ring viton")) - 1.0) < 1e-9)
check("cosine unrelated < 0.2",
      cosine(vec.vec("seal ring viton"), vec.vec("pyro holder bracket")) < 0.2)

# 5) VERDICTS: the two-witness logic
v = build_vectorizer(["seal ring", "retained gravity screw", "pyro holder"])
# same number, float vs text, same desc different case -> STRONG
check("STRONG: same part confirmed",
      agree(7004369.0, "Seal Ring", "7004369", "SEAL RING", v).verdict == "STRONG")
# same number, contradictory description -> CONFLICT (the money finding)
check("CONFLICT: same number, different description",
      agree("7004369", "Seal Ring", "7004369", "Pyro Holder", v).verdict == "CONFLICT")
# different number, same description -> DESC_ONLY (possible typo/relabel)
check("DESC_ONLY: same desc, different number",
      agree("7004369", "Seal Ring", "7004370", "seal ring", v).verdict == "DESC_ONLY")

print("\nALL PASSED" if ok else "\nSOME FAILED")
import sys; sys.exit(0 if ok else 1)
