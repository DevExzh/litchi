"""Record 0361a: per scenario, leaf and unified medians per process, the
unified/leaf ratio, and the after/base ratio of the leaf and unified legs."""
import json, statistics, sys, os
d = sys.argv[1]
labels = [f"{i}-{'base' if i in (1, 4, 5, 8) else 'after'}" for i in range(1, 9)]
def load(fmt, label):
    if fmt == "numbers":
        p = os.path.join(d, f"numbers-{label}.json")
        return {s["name"]: (s["leaf_median_ns"] / 1000, s["unified_median_ns"] / 1000)
                for s in json.load(open(p))["scenarios"]}
    rows = {}
    for line in open(os.path.join(d, f"{fmt}-{label}.jsonl")):
        r = json.loads(line)
        if "leg" in r:
            rows.setdefault(r["scenario"], {}).setdefault(r["leg"], []).append(r["median_ns"])
    return {s: (statistics.median(v["leaf"]) / 1000, statistics.median(v["unified"]) / 1000)
            for s, v in rows.items()}
for fmt in ("keynote", "pages", "numbers"):
    data = {l: load(fmt, l) for l in labels}
    print(f"== {fmt}: scenario | base leaf/unified (median of 4) | after leaf/unified | unified/leaf base..after | after/base leaf, unified")
    for s in data[labels[0]]:
        b = [data[l][s] for l in labels if l.endswith("base") and s in data[l]]
        a = [data[l][s] for l in labels if l.endswith("after") and s in data[l]]
        bl, bu = statistics.median(x[0] for x in b), statistics.median(x[1] for x in b)
        al, au = statistics.median(x[0] for x in a), statistics.median(x[1] for x in a)
        rb = [x[1] / x[0] for x in b]; ra = [x[1] / x[0] for x in a]
        print(f"{s:22} {bl:9.1f}/{bu:9.1f} {al:9.1f}/{au:9.1f}  {min(rb):.3f}-{max(rb):.3f} / {min(ra):.3f}-{max(ra):.3f}  {al/bl:.3f} {au/bu:.3f}")
