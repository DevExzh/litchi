"""Summarize abba.sh output: per route and fixture, the median of the
per-process p50s for each leg and the after/before ratio."""
import json, statistics, sys, collections
rows = [json.loads(line) for line in open(sys.argv[1])]
groups = collections.defaultdict(lambda: {"before": [], "after": []})
loads = [row["load1"] for row in rows]
for row in rows:
    key = (row["route"], row["path"].split("test-data/", 1)[1])
    groups[key][row["leg"]].append(row["p50_ns"] / 1000.0)
print(f"processes: {len(rows)}; 1-min load during run: min {min(loads)} max {max(loads)}")
print("| Route | Fixture | n (per leg) | before p50 µs (median, range) | after p50 µs (median, range) | after/before |")
print("|---|---|---|---|---|---|")
for (route, fixture), legs in groups.items():
    b, a = legs["before"], legs["after"]
    mb, ma = statistics.median(b), statistics.median(a)
    print(f"| `{route}` | `{fixture}` | {len(b)} / {len(a)} | {mb:,.1f} ({min(b):,.1f}–{max(b):,.1f}) | {ma:,.1f} ({min(a):,.1f}–{max(a):,.1f}) | {ma/mb:.3f} |")
