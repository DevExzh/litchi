"""Per-selector p50 for the 0653 marker-selector ABBA, with the A/A floor and
the marker/control ratio on each leg."""
import glob, json, os, sys
OUT = sys.argv[1]

def cases(path):
    doc = json.load(open(path))
    return {row["case"]: row["elapsed_ns"]["p50"] for row in doc["results"]}

legs = {}
for path in sorted(glob.glob(os.path.join(OUT, "*.json"))):
    leg = os.path.basename(path).split("-", 1)[1].split(".")[0]
    legs.setdefault(leg, []).append(cases(path))
names = sorted({n for bs in legs.values() for b in bs for n in b})

def p50(leg, name):
    vs = sorted(b[name] for b in legs.get(leg, []) if name in b)
    return vs[len(vs) // 2] if vs else None

print("%-48s %14s %14s %9s %9s" % ("selector", "before p50", "after p50", "delta", "A/A"))
for name in names:
    before, after = p50("before", name), p50("after", name)
    blocks = [b[name] for leg in ("before", "floorA", "floorB") for b in legs.get(leg, []) if name in b]
    if before is None or after is None or not blocks:
        continue
    floor = 100.0 * (max(blocks) - min(blocks)) / min(blocks)
    print("%-48s %14d %14d %+8.2f%% %8.2f%%" % (name, before, after, 100.0 * (after - before) / before, floor))
print()
print("%-40s %12s %12s" % ("marker / control ratio", "before", "after"))
for marker in names:
    if "_control_" in marker:
        continue
    control = marker.replace("_marker_", "_marker_control_")
    if control not in names:
        continue
    b, a = p50("before", marker), p50("after", marker)
    bc, ac = p50("before", control), p50("after", control)
    if None in (b, a, bc, ac):
        continue
    print("%-40s %11.2fx %11.2fx" % (marker.replace("_marker", ""), b / bc, a / ac))
