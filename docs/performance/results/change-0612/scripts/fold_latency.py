#!/usr/bin/env python3
"""Fold change 0612's A1 B1 B2 A2 wall-clock rounds and reprint every table.

Two modes:
  fold_latency.py --raw <dir> --out <summary.json>
      read the per-sample round captures, write the summary, print the tables.
  fold_latency.py --summary <summary.json>
      print the same tables from the retained summary alone.

The per-sample captures are ~6 MB of JSON per B leg and are not retained; the
summary carries n, p50, mean, p95 and p99 per round together with the SHA-256
of the binary that produced it.

Pure standard library.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import statistics
import sys

PRETTY = {
    "flagship": "ConditionalFormattingSamples.xls",
    "cv": "WithCustomViews.xls",
    "54016": "54016.xls",
}
ORDER = ("flagship", "cv", "54016")
PAT = re.compile(
    r"^lat-(?P<b>[^-]+)-(?P<round>a1|b1|b2|a2)-(?P<stem>flagship|cv|54016)-"
    r"(?P<mode>owned-readat|file-source)-(?P<op>[a-z-]+)\.json$"
)


def quantile(values, q):
    ordered = sorted(values)
    if not ordered:
        return 0.0
    index = min(len(ordered) - 1, max(0, int(round(q * (len(ordered) - 1)))))
    return float(ordered[index])


def fold_raw(directory):
    out = {}
    for name in sorted(os.listdir(directory)):
        m = PAT.match(name)
        if not m:
            continue
        doc = json.load(open(os.path.join(directory, name)))
        samples = doc["elapsed_samples_ns"]
        key = "|".join(
            (m["b"], m["stem"], m["mode"], m["op"], m["round"])
        )
        out[key] = {
            "n": len(samples),
            "p50": quantile(samples, 0.50),
            "mean": statistics.fmean(samples),
            "p95": quantile(samples, 0.95),
            "p99": quantile(samples, 0.99),
            "binary_sha256": doc["binary"]["sha256"],
        }
    return out


def report(summary):
    bnames = sorted({k.split("|")[0] for k in summary})
    ops = sorted({k.split("|")[3] for k in summary})
    for bname in bnames:
        for op in ops:
            print("=" * 78)
            print(f"Paired wall clock, A = base, B = {bname}, operation {op}")
            print("=" * 78)
            print(
                f"{'fixture':<34} {'mode':<13} {'stat':<5} {'A1':>10} {'B1':>10} "
                f"{'B2':>10} {'A2':>10} {'B1/A1':>8} {'B2/A2':>8} {'A2/A1':>8}"
            )
            for stem in ORDER:
                for mode in ("owned-readat", "file-source"):
                    rounds = {}
                    for rnd in ("a1", "b1", "b2", "a2"):
                        rounds[rnd] = summary.get(
                            f"{bname}|{stem}|{mode}|{op}|{rnd}"
                        )
                    if any(v is None for v in rounds.values()):
                        continue
                    for stat in ("p50", "mean", "p95", "p99"):
                        a1 = rounds["a1"][stat]
                        b1 = rounds["b1"][stat]
                        b2 = rounds["b2"][stat]
                        a2 = rounds["a2"][stat]
                        print(
                            f"{PRETTY[stem]:<34} {mode:<13} {stat:<5} "
                            f"{a1:>10,.0f} {b1:>10,.0f} {b2:>10,.0f} {a2:>10,.0f} "
                            f"{100 * (b1 / a1 - 1):>+7.2f}% {100 * (b2 / a2 - 1):>+7.2f}% "
                            f"{100 * (a2 / a1 - 1):>+7.2f}%"
                        )
                    print()
            print()
            print("  B1/A1 and B2/A2 are the paired deltas in the two directions;")
            print("  A2/A1 is the same binary against itself in the same window, the floor.")
            print()


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw")
    ap.add_argument("--summary")
    ap.add_argument("--out")
    args = ap.parse_args(argv)
    if args.raw:
        summary = fold_raw(args.raw)
        if args.out:
            existing = {}
            if os.path.exists(args.out):
                existing = json.load(open(args.out))
            existing.update(summary)
            summary = existing
            with open(args.out, "w") as fh:
                fh.write(json.dumps(summary, indent=2, sort_keys=True) + "\n")
    else:
        summary = json.load(open(args.summary))
    report(summary)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
