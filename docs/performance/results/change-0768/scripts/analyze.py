#!/usr/bin/env python3
"""0768 ABBA analysis.

Reads the raw reports written by abba.py and prints/writes, per case (and per
writer shape for the harness): per-process p50/p95/mean, the median of the
process p50s per arm, the paired after/before ratio of each round's process
p50s, the geometric mean of those paired ratios with a percentile bootstrap
95% CI over rounds, and the same for whole-process perf instructions/cycles.

Usage: analyze.py RUN_DIR OUT_JSON
"""

import csv
import glob
import json
import math
import os
import random
import statistics
import sys


def percentile(values, fraction):
    ordered = sorted(values)
    if not ordered:
        return math.nan
    rank = fraction * (len(ordered) - 1)
    low = math.floor(rank)
    high = math.ceil(rank)
    return ordered[low] + (ordered[high] - ordered[low]) * (rank - low)


def summary(samples):
    return {
        "n": len(samples),
        "p50": percentile(samples, 0.5),
        "p95": percentile(samples, 0.95),
        "mean": statistics.fmean(samples),
        "min": min(samples),
        "max": max(samples),
    }


def perf_counts(path):
    counts = {}
    with open(path) as handle:
        for row in csv.reader(line for line in handle if line.strip() and not line.startswith("#")):
            if len(row) >= 3 and row[0] not in ("<not counted>", "<not supported>"):
                counts[row[2]] = int(row[0])
    return counts


def bootstrap_geomean(ratios, draws=20000, seed=768):
    rng = random.Random(seed)
    logs = [math.log(ratio) for ratio in ratios]
    means = []
    for _ in range(draws):
        sample = [rng.choice(logs) for _ in logs]
        means.append(statistics.fmean(sample))
    return {
        "geomean": math.exp(statistics.fmean(logs)),
        "ci95_low": math.exp(percentile(means, 0.025)),
        "ci95_high": math.exp(percentile(means, 0.975)),
    }


def paired(per_round):
    rounds = sorted(r for r in per_round if "A" in per_round[r] and "B" in per_round[r])
    ratios = [per_round[r]["B"] / per_round[r]["A"] for r in rounds]
    result = {"rounds": rounds, "ratios_after_over_before": ratios}
    result["median_ratio"] = statistics.median(ratios)
    result.update(bootstrap_geomean(ratios))
    return result


def main():
    run_dir, out_json = sys.argv[1], sys.argv[2]
    schedule = json.load(open(os.path.join(run_dir, "schedule.json")))
    series = {}  # (case, shape) -> {"processes": [...]}
    for entry in schedule:
        case, rnd, leg = entry["case"], entry["round"], entry["leg"]
        perf = perf_counts(os.path.join(run_dir, f"perf-{case}-r{rnd}-{leg}.csv"))
        if case.startswith("harness"):
            report = json.load(open(os.path.join(run_dir, f"harness-r{rnd}-{leg}.json")))
            for result in report["results"]:
                samples = result["elapsed_ns"]["samples"]
                key = f"{case}/{result['corpus']['shape']}"
                series.setdefault(key, []).append({
                    "round": rnd, "leg": leg, **summary(samples),
                    "archive_sha256": result["corpus"]["archive_sha256"],
                    "perf": perf,
                })
        else:
            name = "docnohf" if "docnohf" in case else "docfloat"
            report = json.load(open(os.path.join(run_dir, f"probe-{name}-r{rnd}-{leg}.json")))
            samples = [sample["phase_ns"]["whole_ns"] for sample in report["samples"]]
            outputs_ok = all(
                sample["output_sha256"] == report["expected_output_sha256"]
                for sample in report["samples"]
            )
            series.setdefault(case, []).append({
                "round": rnd, "leg": leg, **summary(samples),
                "expected_output_sha256": report["expected_output_sha256"],
                "source_sha256": report["source_sha256"],
                "all_outputs_match_expected": outputs_ok,
                "perf": perf,
            })
    analysis = {}
    for key, processes in sorted(series.items()):
        arms = {}
        for leg in ("A", "B"):
            chosen = [p for p in processes if p["leg"] == leg]
            arms[leg] = {
                "processes": len(chosen),
                "median_of_process_p50_ns": statistics.median(p["p50"] for p in chosen),
                "median_of_process_mean_ns": statistics.median(p["mean"] for p in chosen),
                "median_of_process_p95_ns": statistics.median(p["p95"] for p in chosen),
                "median_instructions": statistics.median(p["perf"]["instructions"] for p in chosen),
                "median_cycles": statistics.median(p["perf"]["cycles"] for p in chosen),
            }
        by_round = lambda field: {
            p["round"]: {**{}, **{}} for p in processes
        }
        p50 = {}
        mean = {}
        instructions = {}
        cycles = {}
        for p in processes:
            p50.setdefault(p["round"], {})[p["leg"]] = p["p50"]
            mean.setdefault(p["round"], {})[p["leg"]] = p["mean"]
            instructions.setdefault(p["round"], {})[p["leg"]] = p["perf"]["instructions"]
            cycles.setdefault(p["round"], {})[p["leg"]] = p["perf"]["cycles"]
        analysis[key] = {
            "arms": arms,
            "paired_p50": paired(p50),
            "paired_mean": paired(mean),
            "paired_process_instructions": paired(instructions),
            "paired_process_cycles": paired(cycles),
            "processes": processes,
        }
    with open(out_json, "w") as handle:
        json.dump(analysis, handle, indent=1)
    for key, value in analysis.items():
        a, b = value["arms"]["A"], value["arms"]["B"]
        print(f"== {key}")
        print(
            f"   p50 median A {a['median_of_process_p50_ns']:.0f} ns  B {b['median_of_process_p50_ns']:.0f} ns"
            f" | paired p50 ratio geomean {value['paired_p50']['geomean']:.4f}"
            f" [{value['paired_p50']['ci95_low']:.4f}, {value['paired_p50']['ci95_high']:.4f}]"
            f" median {value['paired_p50']['median_ratio']:.4f}"
        )
        print(
            f"   mean paired ratio geomean {value['paired_mean']['geomean']:.4f}"
            f" [{value['paired_mean']['ci95_low']:.4f}, {value['paired_mean']['ci95_high']:.4f}]"
        )
        print(
            f"   process instructions A {a['median_instructions']:.4e} B {b['median_instructions']:.4e}"
            f" ratio {value['paired_process_instructions']['geomean']:.5f}"
            f" | cycles ratio {value['paired_process_cycles']['geomean']:.4f}"
            f" [{value['paired_process_cycles']['ci95_low']:.4f}, {value['paired_process_cycles']['ci95_high']:.4f}]"
        )
        flags = [
            f"r{p['round']}"
            for p in value["processes"]
            if p["leg"] == "B"
        ]
        ratios = value["paired_p50"]["ratios_after_over_before"]
        worst = max(ratios)
        print(f"   per-round p50 ratios: {', '.join(f'{r:.3f}' for r in ratios)} (max {worst:.3f})")


if __name__ == "__main__":
    main()
