import json, sys, pathlib, statistics
S = pathlib.Path(sys.argv[1]); tag = sys.argv[2]
legs = {}
for leg in ("A1","B1","B2","A2"):
    p = S / f"abba-{tag}-{leg}.json"
    d = json.loads(p.read_text())
    for r in d["results"]:
        key = (r["case"], r.get("cache_state"), (r.get("corpus") or {}).get("shape"))
        e = r["elapsed_ns"]
        legs.setdefault(key, {})[leg] = e
rows = []
for key in sorted(legs):
    L = legs[key]
    if set(L) != {"A1","B1","B2","A2"}: continue
    rows.append((key, L))
def pct(a, b):  # reduction from a (before) to b (after), positive = faster
    return (a - b) / a * 100.0
print(f"{'case/cache/shape':<58} {'stat':<6} {'A1':>12} {'B1':>12} {'B2':>12} {'A2':>12} {'A1->B1 %':>9} {'A2->B2 %':>9} {'A/A %':>8}")
for key, L in rows:
    label = "/".join(str(x) for x in key)
    for stat in ("p50","mean","p95","p99"):
        a1,b1,b2,a2 = (L[k][stat] for k in ("A1","B1","B2","A2"))
        print(f"{label:<58} {stat:<6} {a1:>12.0f} {b1:>12.0f} {b2:>12.0f} {a2:>12.0f} {pct(a1,b1):>9.2f} {pct(a2,b2):>9.2f} {abs(pct(a1,a2)):>8.2f}")
    print(f"{label:<58} {'n':<6} " + " ".join(f"{len(L[k]['samples']):>12}" for k in ("A1","B1","B2","A2")))
