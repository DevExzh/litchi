#!/usr/bin/env python3
"""Compare a change-0568 before-capture summary against the change-0565
after-capture summary retained in the repository.

Change 0568 touches only the worksheet scan, so its BEFORE state must be exactly
change 0565's AFTER state: open and list at 53 logical reads, one-cell at 319.
Any divergence is reported loudly and sets a non-zero exit code.

Usage: check_control.py --summary DIR/summary.json [--reference PATH]
"""
from __future__ import annotations

import argparse
import json
import pathlib
import statistics

REFERENCE = "/home/zhuhe/code/litchi/docs/performance/results/change-0565/counters/after-summary.json"
LOGICAL = ("read_calls", "read_bytes", "version_calls", "len_calls", "seek_calls")
ISO = ("pread64_per_op", "statx_per_op")
# The numbers change 0565 recorded and this before-capture must reproduce.
EXPECTED_READS = {"open": 53, "list": 53, "one-cell": 319}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--summary", required=True)
    ap.add_argument("--reference", default=REFERENCE)
    ap.add_argument("--output")
    args = ap.parse_args()
    got = json.loads(pathlib.Path(args.summary).read_text())["operations"]
    ref = json.loads(pathlib.Path(args.reference).read_text())["operations"]

    rows, mismatches = [], []
    for key in sorted(set(got) | set(ref)):
        g, r = got.get(key, {}), ref.get(key, {})
        gl, rl = g.get("logical") or {}, r.get("logical") or {}
        gi, ri = g.get("isolation") or {}, r.get("isolation") or {}
        row = {"operation": key}
        for name in LOGICAL:
            row[name] = {"before_0568": gl.get(name), "after_0565": rl.get(name),
                         "match": gl.get(name) == rl.get(name)}
            if gl.get(name) != rl.get(name):
                mismatches.append(f"{key}.{name}: 0568-before={gl.get(name)} vs 0565-after={rl.get(name)}")
        for name in ISO:
            row[name] = {"before_0568": gi.get(name), "after_0565": ri.get(name),
                         "match": gi.get(name) == ri.get(name)}
            if gi.get(name) != ri.get(name):
                mismatches.append(f"{key}.{name}: 0568-before={gi.get(name)} vs 0565-after={ri.get(name)}")
        op = key.split("/", 1)[1]
        if op in EXPECTED_READS and gl.get("read_calls") != EXPECTED_READS[op]:
            mismatches.append(f"{key}.read_calls={gl.get('read_calls')} but change 0565 recorded {EXPECTED_READS[op]}")
        samples = gl.get("elapsed_ns_samples") or []
        row["attribution_p50_ns"] = statistics.median(samples) if samples else None
        rows.append(row)

    result = {"schema_version": 1, "check_kind": "litchi-perf-change-0568-control-check",
              "summary": args.summary, "reference": args.reference,
              "expected_read_calls": EXPECTED_READS,
              "mismatches": mismatches, "verdict": "reproduced" if not mismatches else "DIVERGED",
              "rows": rows}
    if args.output:
        pathlib.Path(args.output).write_text(json.dumps(result, indent=2) + "\n")
    print(f"{'operation':24s} {'reads':>7s} {'ref':>7s} {'bytes':>9s} {'ref':>9s} {'ver':>5s} {'ref':>5s} {'pread':>7s} {'statx':>7s}")
    for row in rows:
        print(f"{row['operation']:24s} {str(row['read_calls']['before_0568']):>7s} {str(row['read_calls']['after_0565']):>7s} "
              f"{str(row['read_bytes']['before_0568']):>9s} {str(row['read_bytes']['after_0565']):>9s} "
              f"{str(row['version_calls']['before_0568']):>5s} {str(row['version_calls']['after_0565']):>5s} "
              f"{str(row['pread64_per_op']['before_0568']):>7s} {str(row['statx_per_op']['before_0568']):>7s}")
    if mismatches:
        print("\n*** DIVERGED FROM CHANGE 0565 ***")
        for m in mismatches:
            print("  ", m)
        return 1
    print("\nverdict: reproduced - the 0568 before-state equals the 0565 after-state")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
