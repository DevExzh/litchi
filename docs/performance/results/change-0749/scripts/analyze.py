#!/usr/bin/env python3
"""Statistics for the change 0749 two-arm matrix.

Per process: p50 (midpoint median), p95 (nearest rank) and mean of the
measured samples. Per case and arm: the median, minimum and maximum of the
process p50s and the median of the process means. Per case: the per-round
paired change cand/base - 1 of p50, mean and p95 (both arms of a round share
one argv[0] heap layout), the median of those paired changes, and a
percentile bootstrap interval of that median (10,000 resamples of rounds,
seed 749). Every per-round change above +5% is listed under `flags`.

Probe cases must publish one output digest across every process of both
arms (the change must not alter a single output byte); the harness validates
its own outputs against its expected bytes and fails the process otherwise.

Usage: analyze.py MATRIX_DIR > analysis.json
"""

import json
import math
import os
import random
import statistics
import sys


def p50(values):
    ordered = sorted(values)
    count = len(ordered)
    return (ordered[(count - 1) // 2] + ordered[count // 2]) / 2


def p95(values):
    ordered = sorted(values)
    return ordered[max(0, math.ceil(0.95 * len(ordered)) - 1)]


def summary(values):
    return {
        "samples": len(values),
        "p50_ns": p50(values),
        "p95_ns": p95(values),
        "mean_ns": statistics.fmean(values),
        "min_ns": min(values),
        "max_ns": max(values),
    }


def load(matrix):
    """{label: {(round, arm): (summary, digest)}}"""
    table = {}
    for case in sorted(os.listdir(matrix)):
        directory = os.path.join(matrix, case)
        if not os.path.isdir(directory):
            continue
        for name in sorted(os.listdir(directory)):
            if not name.endswith(".json"):
                continue
            round_index = int(name[1:3])
            arm = name[4:-5]
            with open(os.path.join(directory, name)) as handle:
                report = json.load(handle)
            if case.startswith("harness-"):
                for result in report["results"]:
                    label = f"{result['case']}/{result['corpus']['shape']}"
                    samples = result["elapsed_ns"]["samples"]
                    table.setdefault(label, {})[(round_index, arm)] = (summary(samples), None)
            else:
                digest = (report["output_sha256"], report["output_bytes"])
                table.setdefault(case, {})[(round_index, arm)] = (
                    summary(report["elapsed_ns"]),
                    digest,
                )
    return table


def bootstrap_median(values, seed=749, resamples=10_000):
    generator = random.Random(seed)
    medians = []
    for _ in range(resamples):
        draw = [values[generator.randrange(len(values))] for _ in values]
        medians.append(statistics.median(draw))
    medians.sort()
    return [medians[int(0.025 * resamples)], medians[int(0.975 * resamples) - 1]]


def main():
    matrix = sys.argv[1]
    table = load(matrix)
    cases = {}
    flags = []
    for label, processes in sorted(table.items()):
        arms = {}
        for arm in ("base", "cand"):
            rows = [value[0] for key, value in sorted(processes.items()) if key[1] == arm]
            if not rows:
                continue
            p50s = [row["p50_ns"] for row in rows]
            arms[arm] = {
                "processes": len(rows),
                "median_process_p50_ns": statistics.median(p50s),
                "min_process_p50_ns": min(p50s),
                "max_process_p50_ns": max(p50s),
                "median_process_mean_ns": statistics.median(row["mean_ns"] for row in rows),
                "median_process_p95_ns": statistics.median(row["p95_ns"] for row in rows),
            }
        rounds = sorted({key[0] for key in processes})
        paired = {"p50": [], "mean": [], "p95": []}
        per_round = []
        for round_index in rounds:
            base = processes.get((round_index, "base"))
            cand = processes.get((round_index, "cand"))
            if not base or not cand:
                continue
            entry = {"round": round_index}
            for metric in paired:
                change = 100.0 * (cand[0][f"{metric}_ns"] / base[0][f"{metric}_ns"] - 1.0)
                paired[metric].append(change)
                entry[f"{metric}_change_pct"] = round(change, 3)
                if change > 5.0:
                    flags.append({"case": label, "round": round_index, "metric": metric,
                                  "change_pct": round(change, 3),
                                  "base_ns": base[0][f"{metric}_ns"], "cand_ns": cand[0][f"{metric}_ns"]})
            per_round.append(entry)
        digests = sorted({value[1] for value in processes.values() if value[1] is not None})
        cases[label] = {
            "arms": arms,
            "paired_rounds": len(paired["p50"]),
            "median_paired_change_pct": {
                metric: round(statistics.median(values), 3) for metric, values in paired.items() if values
            },
            "bootstrap95_median_paired_change_pct": {
                metric: [round(bound, 3) for bound in bootstrap_median(values)]
                for metric, values in paired.items() if values
            },
            "per_round": per_round,
            "output_digests": [list(digest) for digest in digests],
            "single_output_digest": len(digests) <= 1,
        }
    report = {
        "schema": "0749-analysis-v1",
        "cases": cases,
        "flags_above_plus_5_pct": flags,
        "all_probe_outputs_identical_across_arms": all(
            case["single_output_digest"] for case in cases.values()
        ),
    }
    json.dump(report, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
