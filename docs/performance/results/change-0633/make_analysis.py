#!/usr/bin/env python3
"""Renders the retained JSON summaries into the tables change 0633 cites.

Sections, in the order the record uses them:

1. the framing attribution of one `Snapshot::from_bytes` (the frozen design);
2. deterministic counters, before against after;
3. callgrind per-operation inclusive Ir, with the rows that carry the change;
4. native `perf stat` cycles and instructions;
5. paired wall clock, both rounds, with the same-binary floors;
6. the registered selectors.
"""
import json
import subprocess
import sys
from pathlib import Path

here = Path(__file__).parent


def table(rows, headers, aligns=None):
    widths = [max(len(str(h)), *(len(str(r[i])) for r in rows)) if rows else len(str(h))
              for i, h in enumerate(headers)]
    aligns = aligns or [">"] * len(headers)
    print("  ".join(f"{h:{a}{w}}" for h, a, w in zip(headers, aligns, widths)))
    print("  ".join("-" * w for w in widths))
    for row in rows:
        print("  ".join(f"{str(c):{a}{w}}" for c, a, w in zip(row, aligns, widths)))


print("# Change 0633 analysis tables")
print("#")
print("# One 'operation' is one whole open + stage + commit + publish of the")
print("# named scenario through the retained probe, including its per-iteration")
print("# copy of the fixture bytes. Callgrind runs SHA-256 in software, so every")
print("# fingerprint term is an upper bound there and the perf table is the")
print("# native counterpart.")
print()

print("=" * 78)
print("1. FRAMING ATTRIBUTION (before leg) -- what the two framing passes cost")
print("=" * 78)
print((here / "framing-attribution.txt").read_text())

print("=" * 78)
print("2. DETERMINISTIC COUNTERS")
print("=" * 78)


def counters(leg):
    out = {}
    for line in (here / f"counters-{leg}.jsonl").read_text().splitlines():
        row = json.loads(line)
        if "refused" in row:
            continue
        out[(Path(row["input"]).name, row["operation"])] = row
    return out


cb, ca = counters("before"), counters("after")
rows = []
for key in sorted(cb):
    for metric in ("allocations", "allocated_bytes", "peak_live_bytes", "published_bytes"):
        x, y = cb[key].get(metric), ca[key].get(metric)
        if x is None or y is None:
            continue
        delta = "identical" if x == y else f"{(y - x) / x * 100:+.2f}%" if x else "n/a"
        rows.append([f"{key[0]}/{key[1]}", metric, f"{x:,}", f"{y:,}", delta])
table(rows, ["scenario", "metric", "before", "after", "delta"], ["<", "<", ">", ">", ">"])
print()
same = sum(1 for key in cb for m in ("allocations", "allocated_bytes", "peak_live_bytes",
                                     "published_bytes")
           if cb[key].get(m) == ca[key].get(m))
print(f"({len(cb)} scenarios x 4 counters; {same} identical)")
print()
print("source-backed overlay diagnostics, before against after:")
diff = [k for k in cb if json.dumps(cb[k].get("diagnostics"), sort_keys=True)
        != json.dumps(ca[k].get("diagnostics"), sort_keys=True)]
print(f"  rows whose complete diagnostics object differs: {len(diff)}")
print()

print("=" * 78)
print("3. CALLGRIND, per-operation inclusive Ir")
print("=" * 78)
b = json.load(open(here / "callgrind/callgrind-before.json"))
a = json.load(open(here / "callgrind/callgrind-after.json"))
rows = []
for key in sorted(b):
    tb, ta = b[key]["PROGRAM TOTALS"], a[key]["PROGRAM TOTALS"]
    rows.append([key, f"{tb:,}", f"{ta:,}", f"{(ta - tb) / tb * 100:+.2f}%"])
table(rows, ["scenario", "before Ir/op", "after Ir/op", "delta"], ["<", ">", ">", ">"])
print()
print("the rows that carry the change (54016/number-source-backed):")
B, A = b["54016/number-source-backed"], a["54016/number-source-backed"]
rows = []
for key in set(B) | set(A):
    delta = A.get(key, 0) - B.get(key, 0)
    if abs(delta) > 1_000_000:
        rows.append([delta, key])
rows.sort(key=lambda r: -abs(r[0]))
for delta, key in rows[:14]:
    name = key.split(":", 1)[-1].split(" [")[0]
    print(f"  {delta:+15,d}  {name[:120]}")
print()

print("=" * 78)
print("4. NATIVE perf stat, per operation")
print("=" * 78)
pb = json.load(open(here / "perf/perf-before.json"))
pa = json.load(open(here / "perf/perf-after.json"))
rows = []
for key in sorted(pb):
    x, y = pb[key], pa[key]
    rows.append([key,
                 f"{x['cycles']:,.0f}", f"{y['cycles']:,.0f}",
                 f"{(y['cycles'] - x['cycles']) / x['cycles'] * 100:+.2f}%",
                 f"{x['instructions']:,.0f}", f"{y['instructions']:,.0f}",
                 f"{(y['instructions'] - x['instructions']) / x['instructions'] * 100:+.2f}%"])
table(rows, ["scenario", "cycles before", "cycles after", "d",
             "instr before", "instr after", "d"], ["<", ">", ">", ">", ">", ">", ">"])
print()

print("=" * 78)
print("5. PAIRED WALL CLOCK, A1 B1 B2 A2, both rounds")
print("=" * 78)
for label, directory in (("round 1", "latency"), ("round 2", "latency-round2"),
                         ("round 3 (quoted: smallest floor on both changed scenarios)",
                          "latency-round3")):
    print(f"\n-- {label}: p50 ns per phase")
    d = json.load(open(here / directory / "latency-summary.json"))
    rows = []
    for key in sorted(d):
        for phase in ("open_ns", "commit_ns"):
            v = d[key].get(phase)
            if not v or (phase == "commit_ns" and v["a1"]["p50"] < 5000):
                continue
            rows.append([key, phase.removesuffix("_ns"),
                         f"{v['a1']['p50']:,.0f}", f"{v['b1']['p50']:,.0f}",
                         f"{v['delta_forward_p50_pct']:+.2f}%",
                         f"{v['delta_reverse_p50_pct']:+.2f}%",
                         f"{v['floor_a_p50_pct']:+.2f}%",
                         f"{v['floor_b_p50_pct']:+.2f}%"])
    table(rows, ["scenario", "phase", "a1 p50", "b1 p50", "fwd", "rev", "A/A", "B/B"],
          ["<", "<", ">", ">", ">", ">", ">", ">"])
    print("\n   the changed scenario at every quantile:")
    for key in ("54016/number-source-backed", "cv/number-source-backed"):
        v = d[key]["commit_ns"]
        for q in ("mean", "p50", "p95", "p99"):
            print(f"   {key:28s} {q:5s} a1={v['a1'][q]:>12,.0f} b1={v['b1'][q]:>12,.0f} "
                  f"b2={v['b2'][q]:>12,.0f} a2={v['a2'][q]:>12,.0f}  "
                  f"fwd {v[f'delta_forward_{q}_pct']:+7.2f}%  rev {v[f'delta_reverse_{q}_pct']:+7.2f}%  "
                  f"A/A {v[f'floor_a_{q}_pct']:+6.2f}%  B/B {v[f'floor_b_{q}_pct']:+6.2f}%")
print()

print("=" * 78)
print("6. REGISTERED SELECTORS")
print("=" * 78)
print((here / "selectors/selector-summary.txt").read_text())

print((here / "owned-source-chase.txt").read_text())
print()

print("=" * 78)
print("7. CORRECTNESS ORACLES")
print("=" * 78)
print((here / "corpus-summary.txt").read_text())
print("-- 0541-style first-error matrix over synthetic malformed workbooks")
before = (here / "matrix/matrix-before.jsonl").read_text().splitlines()
after = (here / "matrix/matrix-after.jsonl").read_text().splitlines()
print(f"   cases: {len(before)}; rows differing between the legs: "
      f"{sum(1 for x, y in zip(before, after) if x != y)}")
for line in before:
    row = json.loads(line)
    print(f"   {row['case']:38s} {row['outcome']:8s} {row['message']}")
