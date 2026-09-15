"""Paired p50/mean/p95/p99 for change 0621's A1 B1 B2 A2 legs plus the A/A floor."""
import json, statistics, sys, pathlib

OUT = pathlib.Path(sys.argv[1])

def samples(tag):
    data = json.loads((OUT / f"{tag}.json").read_text())
    return data["elapsed_samples_ns"]

def stats(values):
    ordered = sorted(values)
    def q(p):
        index = min(len(ordered) - 1, int(round(p * (len(ordered) - 1))))
        return ordered[index]
    return {
        "p50": statistics.median(ordered),
        "mean": statistics.fmean(ordered),
        "p95": q(0.95),
        "p99": q(0.99),
        "n": len(ordered),
    }

rows = []
for label in ("open-fs", "text-fs", "open-fac", "onecell"):
    legs = {tag: stats(samples(f"{label}-{tag}")) for tag in ("A1", "B1", "B2", "A2", "AA1", "AA2")}
    for stat in ("p50", "mean", "p95", "p99"):
        first = (legs["B1"][stat] - legs["A1"][stat]) / legs["A1"][stat] * 100
        second = (legs["B2"][stat] - legs["A2"][stat]) / legs["A2"][stat] * 100
        floor = (legs["AA2"][stat] - legs["AA1"][stat]) / legs["AA1"][stat] * 100
        rows.append((label, stat, legs["A1"][stat], legs["B1"][stat], first, second, floor))

print(f"{'case':10s} {'stat':5s} {'before_ns':>12s} {'after_ns':>12s} {'dir1_%':>8s} {'dir2_%':>8s} {'AA_%':>8s}")
for label, stat, before, after, first, second, floor in rows:
    print(f"{label:10s} {stat:5s} {before:12,.0f} {after:12,.0f} {first:8.2f} {second:8.2f} {floor:8.2f}")
