"""Percentile summary and paired deltas for the 0653 XLSX timing legs.

The A/A floor is the widest p50 spread among *all* before-leg blocks measured
in the same window -- the two `before` blocks of the ABBA and the two dedicated
floor blocks -- so that a delta is compared against the spread of the identical
binary rather than against a single favourable pair.
"""
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
    print("%-30s n=%-4d p50=%11.0f mean=%11.0f p95=%11.0f p99=%11.0f"
          % (name, len(values), pct(values, .50), statistics.fmean(values),
             pct(values, .95), pct(values, .99)))
    return pct(values, .50)

tags = sorted({os.path.basename(p).split(".")[0] for p in glob.glob(OUT + "/*.txt")})
for tag in tags:
    print("== " + tag)
    before = load(f"{OUT}/{tag}.before.*.txt")
    after = load(f"{OUT}/{tag}.after.*.txt")
    b50 = show("  before (pooled)", before)
    a50 = show("  after (pooled)", after)
    blocks = []
    for pattern in (f"{OUT}/{tag}.before.1.txt", f"{OUT}/{tag}.before.2.txt",
                    f"{OUT}/{tag}.floorA.1.txt", f"{OUT}/{tag}.floorA.2.txt",
                    f"{OUT}/{tag}.floorB.1.txt", f"{OUT}/{tag}.floorB.2.txt"):
        values = load(pattern)
        if values:
            blocks.append(pct(values, .50))
    floor = 100.0 * (max(blocks) - min(blocks)) / min(blocks)
    print("  after vs before p50: %+.2f%%   before vs after p50: %+.2f%%"
          % (100.0 * (a50 - b50) / b50, 100.0 * (b50 - a50) / a50))
    print("  A/A floor (widest p50 spread over %d before blocks): %.2f%%   blocks: %s"
          % (len(blocks), floor, ", ".join("%d" % value for value in blocks)))
