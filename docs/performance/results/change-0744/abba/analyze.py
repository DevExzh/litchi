#!/usr/bin/env python3
"""Change 0744 ABBA analysis.

For every (group, case, shape): per-process p50/p95/mean, the median of the
process p50s per arm, paired after/before ratios over the adjacent ABBA pairs
(A1,B1) (A2,B2) (A3,B3) (A4,B4), a percentile bootstrap 95% CI of the mean
paired p50 ratio over those pairs, same-arm drift, and >5% regression flags.
"""
import glob
import json
import os
import random
import statistics
import sys

RAW = sys.argv[1] if len(sys.argv) > 1 else "raw"
OUT = sys.argv[2] if len(sys.argv) > 2 else "analysis.json"
PAIRS = [("A1", "B1"), ("A2", "B2"), ("A3", "B3"), ("A4", "B4")]


def load(path):
    with open(path) as handle:
        data = json.load(handle)
    rows = {}
    for result in data["results"]:
        key = (result["case"], result["corpus"]["shape"])
        elapsed = result["elapsed_ns"]
        rows[key] = {
            "p50": elapsed["p50"],
            "p95": elapsed["p95"],
            "mean": elapsed["mean"],
            "samples": len(elapsed["samples"]),
            "output_sha256": result.get("output_sha256"),
            "accepted_bytes": (result.get("sink") or {}).get("accepted_bytes"),
            "archive_sha256": result["corpus"].get("archive_sha256"),
        }
    return data, rows


def bootstrap(ratios, iterations=20000, seed=744):
    generator = random.Random(seed)
    means = []
    for _ in range(iterations):
        sample = [generator.choice(ratios) for _ in ratios]
        means.append(sum(sample) / len(sample))
    means.sort()
    return means[int(0.025 * iterations)], means[int(0.975 * iterations) - 1]


def main():
    groups = sorted({os.path.basename(path).split("-")[0] for path in glob.glob(f"{RAW}/*.json")})
    report = {"method": __doc__.strip(), "groups": {}}
    flags = []
    for group in groups:
        processes = {}
        environment = None
        for slot in ["A1", "B1", "B2", "A2", "A3", "B3", "B4", "A4"]:
            path = f"{RAW}/{group}-{slot}.json"
            if not os.path.exists(path):
                continue
            data, rows = load(path)
            processes[slot] = rows
            environment = environment or data.get("environment")
        keys = sorted({key for rows in processes.values() for key in rows})
        entries = []
        for key in keys:
            per = {slot: rows[key] for slot, rows in processes.items() if key in rows}
            a = [per[s]["p50"] for s in per if s.startswith("A")]
            b = [per[s]["p50"] for s in per if s.startswith("B")]
            ratios = [per[bs]["p50"] / per[as_]["p50"] for as_, bs in PAIRS if as_ in per and bs in per]
            mean_ratios = [per[bs]["mean"] / per[as_]["mean"] for as_, bs in PAIRS if as_ in per and bs in per]
            low, high = bootstrap(ratios) if ratios else (None, None)
            entry = {
                "case": key[0],
                "shape": key[1],
                "processes": {slot: per[slot] for slot in sorted(per)},
                "before_median_p50_ns": statistics.median(a) if a else None,
                "after_median_p50_ns": statistics.median(b) if b else None,
                "paired_p50_ratios": ratios,
                "paired_mean_ratios": mean_ratios,
                "mean_paired_p50_ratio": sum(ratios) / len(ratios) if ratios else None,
                "bootstrap95_mean_paired_p50_ratio": [low, high],
                "before_p50_drift": (max(a) / min(a) - 1) if a else None,
                "after_p50_drift": (max(b) / min(b) - 1) if b else None,
                "outputs_identical": len({per[s]["output_sha256"] for s in per}) == 1
                and len({per[s]["accepted_bytes"] for s in per}) == 1,
                "corpus_identical": len({per[s]["archive_sha256"] for s in per}) == 1,
            }
            if ratios and max(ratios) > 1.05:
                flags.append({"group": group, "case": key[0], "shape": key[1], "max_paired_p50_ratio": max(ratios)})
            entries.append(entry)
        report["groups"][group] = {"environment": environment, "entries": entries}
    report["regression_flags_over_5pct"] = flags
    with open(OUT, "w") as handle:
        json.dump(report, handle, indent=1)
    print(f"{'case':46s} {'shape':11s} {'before p50':>12s} {'after p50':>12s} {'ratio':>7s} {'CI95':>17s} {'driftA':>7s} {'driftB':>7s} same-out")
    for group, value in report["groups"].items():
        for entry in value["entries"]:
            before = entry["before_median_p50_ns"] / 1e6
            after = entry["after_median_p50_ns"] / 1e6
            low, high = entry["bootstrap95_mean_paired_p50_ratio"]
            print(
                f"{entry['case']:46s} {entry['shape']:11s} {before:12.4f} {after:12.4f} "
                f"{entry['mean_paired_p50_ratio']:7.4f} [{low:.4f},{high:.4f}] "
                f"{entry['before_p50_drift']*100:6.2f}% {entry['after_p50_drift']*100:6.2f}% "
                f"{entry['outputs_identical'] and entry['corpus_identical']}"
            )
    print("flags:", flags)


if __name__ == "__main__":
    main()
