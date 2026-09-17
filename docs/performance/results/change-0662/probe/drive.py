#!/usr/bin/env python3
"""Paired-leg driver for change 0662: runs legs in A1 B1 B2 A2 order and
reports p50/mean/p95/p99 per leg."""
import subprocess, sys, statistics, json

BIN = "/home/zhuhe/code/litchi-worktrees/targets/0662-probe/release/probe0662"
def _quietest(count=8):
    """Pins to CPU 17 (this change's assigned CPU) plus the quietest others.

    One CPU cannot measure a width-8 wave, and the host's other CPUs are shared
    with twelve concurrent agents, so the set is chosen from measured idleness
    immediately before the run and is reported with the result."""
    import time
    def snap():
        out = {}
        for line in open("/proc/stat"):
            if line.startswith("cpu") and line[3].isdigit():
                f = line.split()
                vals = [int(x) for x in f[1:]]
                out[int(f[0][3:])] = (sum(vals), vals[3] + vals[4])
        return out
    a = snap(); time.sleep(2); b = snap()
    util = {}
    for cpu in a:
        dt = b[cpu][0] - a[cpu][0]
        util[cpu] = 100 * (1 - (b[cpu][1] - a[cpu][1]) / dt) if dt else 100.0
    chosen = [17] + [c for c, _ in sorted(((c, u) for c, u in util.items() if c != 17),
                                          key=lambda kv: kv[1])][: count - 1]
    return ",".join(str(c) for c in sorted(chosen)), {c: round(util[c], 1) for c in sorted(chosen)}

CPU, CPU_UTIL = _quietest()
import sys as _sys
print(f"# cpu set {CPU} idle-measured util {CPU_UTIL}", file=_sys.stderr)

def run(args, samples):
    out = subprocess.run(["taskset", "-c", CPU, BIN] + [str(a) for a in args],
                         capture_output=True, text=True)
    if out.returncode != 0:
        raise SystemExit(f"probe failed: {out.stderr[-2000:]}")
    values = [int(line) for line in out.stdout.split() if line.strip()]
    if len(values) != samples:
        raise SystemExit(f"expected {samples} samples, got {len(values)}")
    return values, out.stderr.strip()

def stats(values):
    values = sorted(values)
    n = len(values)
    def pct(p):
        return values[min(n - 1, int(round(p * (n - 1))))]
    return {"n": n, "p50": pct(0.50), "mean": round(statistics.fmean(values)),
            "p95": pct(0.95), "p99": pct(0.99)}

def paired(args_a, args_b, samples, warmup):
    """A1 B1 B2 A2."""
    a1, _ = run(args_a + [warmup, samples], samples)
    b1, note = run(args_b + [warmup, samples], samples)
    b2, _ = run(args_b + [warmup, samples], samples)
    a2, _ = run(args_a + [warmup, samples], samples)
    return stats(a1 + a2), stats(b1 + b2), note
