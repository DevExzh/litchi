#!/usr/bin/env python3
"""Paired-median summary for change 0641's timing legs.

A1/A2/A3 are the before binary, B1/B2 the after binary, run in the order
A1 B1 B2 A2 A3 on one pinned CPU. `dir1` is A1->B1, `dir2` is A2->B2, and the
A/A floor is A1 against A3: the same binary against itself in the same window.
"""
import json, os, statistics, sys

S = "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0641"

def load(path):
    return sorted(int(line) for line in open(path) if line.strip())

def pct(values, q):
    values = sorted(values)
    if q == 50:
        return statistics.median(values)
    index = min(len(values) - 1, int(round(q / 100.0 * (len(values) - 1))))
    return values[index]

rows = []
for line in open(sys.argv[1] if len(sys.argv) > 1 else S + "/time-scenarios.txt"):
    label = line.split("|")[0].strip()
    if not label:
        continue
    legs = {}
    for leg in ("A1", "B1", "B2", "A2", "A3"):
        p = f"{S}/timing/{label}-{leg}.txt"
        if not os.path.exists(p) or os.path.getsize(p) == 0:
            legs = None
            break
        legs[leg] = load(p)
    if legs is None:
        print(f"{label}: missing legs")
        continue
    a1, b1, b2, a2, a3 = (legs[k] for k in ("A1", "B1", "B2", "A2", "A3"))
    row = dict(
        scenario=label,
        n=len(a1),
        before_p50=pct(a1 + a2, 50), after_p50=pct(b1 + b2, 50),
        before_mean=round(statistics.mean(a1 + a2)), after_mean=round(statistics.mean(b1 + b2)),
        before_p95=pct(a1 + a2, 95), after_p95=pct(b1 + b2, 95),
        before_p99=pct(a1 + a2, 99), after_p99=pct(b1 + b2, 99),
        dir1=round(100.0 * (pct(b1, 50) - pct(a1, 50)) / pct(a1, 50), 2),
        dir2=round(100.0 * (pct(b2, 50) - pct(a2, 50)) / pct(a2, 50), 2),
        aa=round(100.0 * (pct(a3, 50) - pct(a1, 50)) / pct(a1, 50), 2),
        aa2=round(100.0 * (pct(a2, 50) - pct(a1, 50)) / pct(a1, 50), 2),
        bb=round(100.0 * (pct(b2, 50) - pct(b1, 50)) / pct(b1, 50), 2),
    )
    rows.append(row)

print("%-24s %5s %11s %11s %8s %8s %8s %8s %8s" %
      ("scenario", "n", "before p50", "after p50", "dir1", "dir2", "A/A", "A/A2", "B/B"))
for r in rows:
    print("%-24s %5d %11d %11d %7.2f%% %7.2f%% %7.2f%% %7.2f%% %7.2f%%" %
          (r["scenario"], r["n"], r["before_p50"], r["after_p50"],
           r["dir1"], r["dir2"], r["aa"], r["aa2"], r["bb"]))
print()
print("%-24s %11s %11s %11s %11s" % ("scenario", "before p95", "after p95", "before p99", "after p99"))
for r in rows:
    print("%-24s %11d %11d %11d %11d" %
          (r["scenario"], r["before_p95"], r["after_p95"], r["before_p99"], r["after_p99"]))
json.dump(rows, open(S + "/timing/summary.json", "w"), indent=1)
