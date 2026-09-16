"""Per-selector p50 for the 0653 harness ABBA, with the A/A floor of the same window."""
import glob, json, os, sys
OUT = sys.argv[1]

def cases(path):
    with open(path) as handle:
        doc = json.load(handle)
    out = {}
    for row in doc.get("results", []):
        name = row.get("case") or row.get("name")
        stats = row.get("elapsed") or row.get("timing") or row
        for key in ("p50_ns", "median_ns", "p50"):
            value = stats.get(key) if isinstance(stats, dict) else None
            if name and isinstance(value, (int, float)):
                out[name] = value
                break
    return out

for prefix in ("x", "p"):
    legs = {}
    for path in sorted(glob.glob(os.path.join(OUT, prefix + "-*.json"))):
        leg = os.path.basename(path).split("-", 1)[1].split(".")[0]
        legs.setdefault(leg, []).append(cases(path))
    if not legs:
        continue
    names = sorted({name for blocks in legs.values() for block in blocks for name in block})
    print("%-46s %14s %14s %9s %9s" % ("selector", "before p50", "after p50", "delta", "A/A"))
    for name in names:
        def p50(leg):
            values = sorted(block[name] for block in legs.get(leg, []) if name in block)
            return values[len(values) // 2] if values else None
        before, after = p50("before"), p50("after")
        blocks = [block[name] for leg in ("before", "floorA", "floorB")
                  for block in legs.get(leg, []) if name in block]
        if before is None or after is None or not blocks:
            continue
        floor = 100.0 * (max(blocks) - min(blocks)) / min(blocks)
        print("%-46s %14d %14d %+8.2f%% %8.2f%%"
              % (name, before, after, 100.0 * (after - before) / before, floor))
    print()
