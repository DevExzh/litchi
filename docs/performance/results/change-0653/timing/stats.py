"""Percentile summary and paired deltas for the 0588 timing legs."""
import glob, os, statistics, sys

OUT = sys.argv[1]

def load(pattern):
    values = []
    for path in sorted(glob.glob(pattern)):
        with open(path) as handle:
            values.extend(int(line) for line in handle if line.strip())
    return sorted(values)

def pct(values, q):
    if not values:
        return float("nan")
    index = min(len(values) - 1, max(0, int(round(q * (len(values) - 1)))))
    return values[index]

def show(name, values):
    print("%-34s n=%-4d p50=%10.0f mean=%10.0f p95=%10.0f p99=%10.0f"
          % (name, len(values), pct(values, .50), statistics.fmean(values),
             pct(values, .95), pct(values, .99)))
    return pct(values, .50)

tags = sorted({os.path.basename(p).split(".")[0] for p in glob.glob(OUT + "/*.txt")})
for tag in tags:
    before = load(f"{OUT}/{tag}.before.*.txt")
    after = load(f"{OUT}/{tag}.after.*.txt")
    floor_a = load(f"{OUT}/{tag}.floorA.*.txt")
    floor_b = load(f"{OUT}/{tag}.floorB.*.txt")
    print("== " + tag)
    b50 = show("  before", before)
    a50 = show("  after", after)
    print("  after vs before p50: %+.2f%%   before vs after p50: %+.2f%%"
          % (100.0 * (a50 - b50) / b50, 100.0 * (b50 - a50) / a50))
    if floor_a and floor_b:
        f1 = show("  floor A (before)", floor_a)
        f2 = show("  floor B (before)", floor_b)
        print("  A/A floor p50: %+.2f%%   p99: %+.2f%%"
              % (100.0 * (f2 - f1) / f1,
                 100.0 * (pct(floor_b, .99) - pct(floor_a, .99)) / pct(floor_a, .99)))
    print()
