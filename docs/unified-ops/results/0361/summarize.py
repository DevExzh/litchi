import json, statistics, sys, os
d = sys.argv[1]
labels = ["1-base", "2-after", "3-after", "4-base"]
out = []
for fmt in ["keynote", "pages"]:
    table = {}
    for l in labels:
        p = os.path.join(d, f"{fmt}-{l}.jsonl")
        if not os.path.exists(p): continue
        rows = {}
        for line in open(p):
            r = json.loads(line)
            if "leg" not in r: continue
            rows.setdefault(r["scenario"], {}).setdefault(r["leg"], []).append(r["median_ns"])
        for s, legs in rows.items():
            lf = statistics.median(legs["leaf"]); un = statistics.median(legs["unified"])
            table.setdefault(s, {})[l] = (lf / 1000, un / 1000, un / lf)
    out.append(f"== {fmt} (median of per-sample medians; leaf/unified us, ratio unified/leaf)")
    out.append(f"{'scenario':24}" + "".join(f"{l:>28}" for l in labels))
    for s, v in table.items():
        cells = "".join(f"{('%.1f/%.1f %.3f' % v[l]) if l in v else '-':>28}" for l in labels)
        out.append(f"{s:24}{cells}")
table = {}
for l in labels:
    p = os.path.join(d, f"numbers-{l}.json")
    if not os.path.exists(p): continue
    for s in json.load(open(p))["scenarios"]:
        table.setdefault(s["name"], {})[l] = (s["leaf_median_ns"] / 1000, s["unified_median_ns"] / 1000, s["ratio"], s["outputs_identical"])
out.append("== numbers (leaf/unified us, ratio unified/leaf; identical outputs)")
out.append(f"{'scenario':24}" + "".join(f"{l:>28}" for l in labels))
for s, v in table.items():
    cells = "".join(f"{('%.1f/%.1f %.3f' % v[l][:3]) if l in v else '-':>28}" for l in labels)
    ident = next(iter(v.values()))[3]
    out.append(f"{s:24}{cells} {ident}")
print("\n".join(out))
# Record 0361: the batch-closure 1k/10k cases per process.
import glob
rows = {}
for p in sorted(glob.glob(os.path.join(d, "abba", "batch-A-*.json"))):
    name = os.path.basename(p)[len("batch-A-"):-len(".json")]
    for c in json.load(open(p))["cases"]:
        rows.setdefault((c["format"], c["scenario"], c["operations"]), {})[name] = (
            c["leaf_median_ns"] / 1000, c["unified_median_ns"] / 1000, c["ratio"])
cols = ["base-1", "after-2", "after-3", "base-4"]
print("== batch-closure (leaf/unified us, ratio unified/leaf)")
print(f"{'case':36}" + "".join(f"{c:>28}" for c in cols))
for k, v in sorted(rows.items()):
    print(f"{'%s %s %d' % k:36}" + "".join(f"{('%.1f/%.1f %.3f' % v[c]) if c in v else '-':>28}" for c in cols))
