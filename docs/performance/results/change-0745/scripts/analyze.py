#!/usr/bin/env python3
"""Analyze a change 0745 process matrix produced by run.py.

Per process: n, p50 (midpoint median), p95 (nearest rank), mean, min, max.
Per case and arm: median of process p50 / mean / p95.
Paired comparisons (B vs A, C vs B, C vs A): for every round (one heap layout,
all arms), the percentage change of process p50, mean and p95; the median
paired change, its min/max over rounds, and a percentile bootstrap 95%
interval of the median over rounds (10,000 resamples, seed 745). Every paired
round change above +5% and every median change above +5% is listed as a
regression flag.

Oracles: every PPT/DOC probe process must publish the same output digest (and
durable digest) in every arm; every 0734-probe sample must pass its own oracle
and match its expected output digest.

Usage: analyze.py MATRIX_DIR > analysis.json
"""

import glob
import json
import math
import os
import random
import statistics
import sys

COMPARISONS = (("B", "A"), ("C", "B"), ("C", "A"))
STATS = ("p50_ns", "mean_ns", "p95_ns")


def p95(values):
    ordered = sorted(values)
    return ordered[max(1, math.ceil(0.95 * len(ordered))) - 1]


def summarize(values):
    return {
        "n": len(values),
        "p50_ns": statistics.median(values),
        "p95_ns": p95(values),
        "mean_ns": statistics.fmean(values),
        "min_ns": min(values),
        "max_ns": max(values),
    }


def load_probe(path):
    with open(path) as handle:
        report = json.load(handle)
    oracle = {
        "output_sha256": report["output_sha256"],
        "output_bytes": report["output_bytes"],
        "durable_sha256": report.get("durable_sha256"),
        "durable_bytes": report.get("durable_bytes"),
    }
    return {"": (report["elapsed_ns"], oracle)}


def load_p0734(path):
    with open(path) as handle:
        report = json.load(handle)
    expected = report["expected_output_sha256"]
    samples = report["samples"]
    if not report["expected_oracle"]["oracle_ok"] or not all(
        sample["output_sha256"] == expected and sample["oracle"]["oracle_ok"]
        for sample in samples
    ):
        raise SystemExit(f"0734 oracle failed in {path}")
    return {"": ([sample["phase_ns"]["whole_ns"] for sample in samples], {"output_sha256": expected})}


def load_harness(path):
    with open(path) as handle:
        report = json.load(handle)
    return {
        f"{entry['case']}/{entry['corpus']['shape']}": (
            entry["elapsed_ns"]["samples"],
            {"corpus": entry["corpus"]["shape"]},
        )
        for entry in report["results"]
    }


def loader(case):
    if case.startswith("p0734-"):
        return load_p0734
    if case.startswith("harness-"):
        return load_harness
    return load_probe


def bootstrap(changes, seed=745, resamples=10_000):
    generator = random.Random(seed)
    medians = sorted(
        statistics.median([generator.choice(changes) for _ in changes])
        for _ in range(resamples)
    )
    return [medians[int(0.025 * resamples)], medians[int(0.975 * resamples) - 1]]


def main():
    matrix = sys.argv[1]
    analysis = {"matrix": os.path.basename(os.path.abspath(matrix)), "cases": {}, "flags": []}
    for case_dir in sorted(glob.glob(f"{matrix}/*/")):
        case = os.path.basename(os.path.normpath(case_dir))
        load = loader(case)
        per_key = {}
        for path in sorted(glob.glob(f"{case_dir}/r*-[ABC].json")):
            stem = os.path.basename(path)[:-5]
            round_index, arm = int(stem[1:3]), stem[-1]
            for key, (values, oracle) in load(path).items():
                per_key.setdefault(key, {}).setdefault(round_index, {})[arm] = (summarize(values), oracle)
        for key, rounds in per_key.items():
            label = f"{case}{'/' + key if key else ''}"
            arms = sorted({arm for by_arm in rounds.values() for arm in by_arm})
            oracles = {json.dumps(by_arm[arm][1], sort_keys=True) for by_arm in rounds.values() for arm in by_arm}
            entry = {
                "arms": arms,
                "rounds": sorted(rounds),
                "oracle_values": sorted(json.loads(value) for value in oracles),
                "oracle_identical_across_arms": len(oracles) == 1,
                "processes": {arm: {str(r): rounds[r][arm][0] for r in sorted(rounds) if arm in rounds[r]} for arm in arms},
                "arm_medians": {
                    arm: {stat: statistics.median(rounds[r][arm][0][stat] for r in rounds if arm in rounds[r]) for stat in STATS}
                    for arm in arms
                },
                "comparisons": {},
            }
            for after, before in COMPARISONS:
                if after not in arms or before not in arms:
                    continue
                paired = sorted(r for r in rounds if after in rounds[r] and before in rounds[r])
                comparison = {"rounds": len(paired)}
                for stat in STATS:
                    changes = [
                        (rounds[r][after][0][stat] / rounds[r][before][0][stat] - 1.0) * 100.0
                        for r in paired
                    ]
                    comparison[stat] = {
                        "median_change_pct": statistics.median(changes),
                        "min_change_pct": min(changes),
                        "max_change_pct": max(changes),
                        "bootstrap95_pct": bootstrap(changes),
                        "paired_changes_pct": [round(change, 3) for change in changes],
                    }
                    for r, change in zip(paired, changes):
                        if change > 5.0:
                            analysis["flags"].append({"case": label, "comparison": f"{after}/{before}", "stat": stat, "round": r, "change_pct": round(change, 2)})
                    if comparison[stat]["median_change_pct"] > 5.0:
                        analysis["flags"].append({"case": label, "comparison": f"{after}/{before}", "stat": stat, "round": "median", "change_pct": round(comparison[stat]["median_change_pct"], 2)})
                entry["comparisons"][f"{after}/{before}"] = comparison
            analysis["cases"][label] = entry
    json.dump(analysis, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
