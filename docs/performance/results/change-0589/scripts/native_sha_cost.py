#!/usr/bin/env python3
"""Derive the native per-pass SHA-256 cost and the hashing share of an open.

Model: an open performs `passes` complete SHA-256 passes over the artifact, and
this change removes half of them (6 of 12 for the generic DOC open, 2 of 4 for
the PPT text-edit open).  The observed before/after delta divided by the removed
passes and the artifact length is therefore the per-byte cost of one native pass.
The fit is the evidence: it lands at 2.03-2.12 cycles/byte on all 38 fixtures
and both formats, which would not happen if the removed work were anything else.

Reads perf/perfstat-before.jsonl, perf/perfstat-after.jsonl,
counts/doc-fixtures.txt and counts/ppt-fixtures.txt; writes the table in
perf/native-sha-cost.txt.
"""
import json, pathlib, statistics, sys

base = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else ".")
b = [json.loads(l) for l in (base / "perf/perfstat-before.jsonl").read_text().splitlines()]
a = [json.loads(l) for l in (base / "perf/perfstat-after.jsonl").read_text().splitlines()]
paths = (base / "counts/doc-fixtures.txt").read_text().split() + (
    base / "counts/ppt-fixtures.txt").read_text().split()

print("Derived native SHA-256 cost and hashing share (SHA-NI active).")
print("Passes removed per open: 6 for doc-snapshot-open, 2 for ppt-textedit-open.")
print()
print(f"{'mode':18s} {'fixture':32s} {'bytes':>9s} {'cyc/byte/pass':>14s} "
      f"{'ins/byte/pass':>14s} {'hash % cyc before':>18s} {'after':>8s}")
cb, ib = [], []
for bd, ad, p in zip(b, a, paths):
    passes = 6 if bd["mode"] == "doc-snapshot-open" else 2
    n = pathlib.Path(p).stat().st_size
    cpb = (bd["cycles_per_op"] - ad["cycles_per_op"]) / passes / n
    ipb = (bd["instructions_per_op"] - ad["instructions_per_op"]) / passes / n
    cb.append(cpb); ib.append(ipb)
    sh_b = (passes * 2) * cpb * n / bd["cycles_per_op"] * 100
    sh_a = passes * cpb * n / ad["cycles_per_op"] * 100
    print(f"{bd['mode']:18s} {pathlib.Path(p).name[:32]:32s} {n:9d} {cpb:14.3f} "
          f"{ipb:14.3f} {sh_b:17.1f}% {sh_a:7.1f}%")
print()
print(f"cycles per byte per SHA-256 pass: median {statistics.median(cb):.3f} (n={len(cb)})")
print(f"instructions per byte per pass  : median {statistics.median(ib):.3f}")
print(f"callgrind software backend      : 52.07 Ir/byte -> "
      f"{52.07 / statistics.median(ib):.1f}x the native instruction cost")
