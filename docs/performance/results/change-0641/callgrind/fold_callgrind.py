#!/usr/bin/env python3
"""Difference two callgrind profiles of LOW and HIGH samples, per symbol.

Change 0641. `callgrind_annotate` reports self cost per function for a whole
process; the harness's staging, oracle and reporting are identical between a
LOW-sample and a HIGH-sample run, so the difference divided by HIGH-LOW is one
operation's self cost per symbol, free of the fixed cost.
"""
import re, subprocess, sys, os
S = "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0641"
LOW, HIGH = 2, 12
M = HIGH - LOW

def annotate(path):
    out = subprocess.run(["callgrind_annotate", "--threshold=99.9", "--auto=no", path],
                         capture_output=True, text=True).stdout
    totals = {}
    for line in out.splitlines():
        m = re.match(r"^\s*([\d,]+)\s+(\S.*)$", line)
        if not m:
            continue
        name = m.group(2).strip()
        if ":" not in name:
            continue
        name = re.sub(r"\s*\[.*\]$", "", name.split(":", 1)[1].strip())
        totals[name] = totals.get(name, 0) + int(m.group(1).replace(",", ""))
    return totals

def per_op(label, leg):
    lo = annotate(f"{S}/cg/{label}-{leg}-{LOW}.out")
    hi = annotate(f"{S}/cg/{label}-{leg}-{HIGH}.out")
    return {n: (hi.get(n, 0) - lo.get(n, 0)) / M for n in set(lo) | set(hi)
            if hi.get(n, 0) - lo.get(n, 0)}

def total(path):
    for line in open(path, errors="ignore"):
        if line.startswith("summary:"):
            return int(line.split(":")[1].strip().split()[0])

for label in sys.argv[1:]:
    if not os.path.exists(f"{S}/cg/{label}-after-{HIGH}.out"):
        print(f"{label}: not captured"); continue
    gb = (total(f"{S}/cg/{label}-before-{HIGH}.out") - total(f"{S}/cg/{label}-before-{LOW}.out")) / M
    ga = (total(f"{S}/cg/{label}-after-{HIGH}.out") - total(f"{S}/cg/{label}-after-{LOW}.out")) / M
    sb, sa = per_op(label, "before"), per_op(label, "after")
    print(f"### {label}: {gb:,.0f} -> {ga:,.0f} Ir per operation ({100.0*(ga-gb)/gb:+.2f}%)")
    moved = sorted(set(sb) | set(sa), key=lambda n: sa.get(n, 0) - sb.get(n, 0))
    for name in moved[:10] + moved[-6:]:
        b, a = sb.get(name, 0.0), sa.get(name, 0.0)
        if abs(a - b) < max(1500, 0.0008 * gb):
            continue
        print("    {:<88} {:>12,.0f} -> {:>12,.0f} ({:+,.0f})".format(name[:88], b, a, a - b))
    print()
