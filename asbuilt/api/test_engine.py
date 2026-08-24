"""Engine tests. No database needed:  python3 test_engine.py"""
import engine as E

def chk(name, got, want):
    ok = got == want
    print(("PASS " if ok else "FAIL ") + name, "" if ok else f"  got={got!r} want={want!r}")
    return ok

fails = 0
# --- match_key: type, case, space blind; suffix preserved
fails += not chk("float key",   E.match_key("7004369.0"), "7004369")
fails += not chk("case/space",  E.match_key(" c3529115-c "), "C3529115-C")
fails += not chk("suffix kept", E.match_key("C3529115-C") != E.match_key("C3529115"), True)
fails += not chk("dotted part", E.match_key("2104.3275"), "2104.3275")

# --- quantity sums across positions, positions counted
g = E.group([{"material":"X","qty":2,"revision":"A"},{"material":"x ","qty":3,"revision":"A"}])
fails += not chk("qty summed",  g["X"]["qty"], 5.0)
fails += not chk("positions",   g["X"]["positions"], 2)

exp = [{"material":"A","qty":10,"revision":"A","traceable":True},
       {"material":"B","qty":5,"revision":"B"},
       {"material":"C","qty":4,"revision":"C","traceable":True},
       {"material":"D","qty":1,"revision":"A"},
       {"material":"E","qty":2,"revision":"A"}]
pre = [{"material":"A","qty":10,"revision":"A","batch":"1234"},
       {"material":"B","qty":3,"revision":"B","batch":"999"},
       {"material":"C","qty":4,"revision":"D","batch":"777"},
       {"material":"D","qty":1,"revision":"A"},
       {"material":"Z","qty":9,"revision":"A"}]
r = E.compare(exp, pre)
v = {x["key"]: x["verdict"] for x in r["rows"]}
fails += not chk("A matched",        v["A"], "OK")
fails += not chk("B short",          v["B"], "SHORT")
fails += not chk("C revision",       v["C"], "REVISION")
fails += not chk("D no batch flag",  v["D"], "OK")        # not traceable -> no flag
fails += not chk("E missing",        v["E"], "MISSING")
fails += not chk("Z extra",          v["Z"], "EXTRA")
fails += not chk("attention count",  r["counts"]["ATTENTION"], 4)

# traceable + no batch -> flagged; same part not traceable -> silent
one = E.compare([{"material":"T","qty":1,"traceable":True}], [{"material":"T","qty":1}])
fails += not chk("traceable no batch", one["rows"][0]["verdict"], "BATCH")
two = E.compare([{"material":"T","qty":1}], [{"material":"T","qty":1}])
fails += not chk("plain no batch",     two["rows"][0]["verdict"], "OK")

# revision only judges when BOTH sides carry one
rv = E.compare([{"material":"R","qty":1,"revision":"A"}], [{"material":"R","qty":1}])
fails += not chk("revision one-sided", rv["rows"][0]["verdict"], "OK")

# --- header mapping on a real-world header row with a banner above it
sheet = [["NOTE: fed back from a previous report","",""],
         ["Material","Rev","Charge","Menge","Bezeichnung"],
         ["3001180","D","293520","2","Vent Port"],
         ["","","","",""],
         ["C3529124-C","A","288018","144","USIT-Ring"]]
rows, fields = E.read_rows(sheet)
fails += not chk("banner skipped", len(rows), 2)
fails += not chk("fields found",   sorted(fields), ["batch","description","material","qty","revision"])
fails += not chk("qty parsed",     rows[1]["qty"], 144.0)

try:
    E.read_rows([["alpha","beta"],["1","2"]])
    print("FAIL junk sheet should raise")
    fails += 1
except ValueError as e:
    print("PASS junk sheet raises:", str(e)[:52])

print("\n" + ("ALL PASS" if not fails else f"{fails} FAILED"))