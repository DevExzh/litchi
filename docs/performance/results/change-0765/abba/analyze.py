#!/usr/bin/env python3
"""Change 0765 control ABBA analysis.

For every (group, case, shape): per-process p50/p95/mean and sample count, the
median of the process p50s per arm, paired after/before p50 ratios over the
adjacent ABBA pairs (A1,B1) (A2,B2) (A3,B3) (A4,B4), a percentile bootstrap
95% CI of the mean paired p50 ratio (20,000 resamples, seed 765), same-arm
drift, whole-process user-space instructions and cycles from `perf stat`
(including untimed corpus construction) with their paired ratios, output and
corpus identity, and >5% regression flags on any paired p50 ratio.
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
SLOTS = ["A1", "B1", "B2", "A2", "A3", "B3", "B4", "A4"]


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


def load_perf(path):
    counters = {}
    if not os.path.exists(path):
        return counters
    with open(path) as handle:
        for line in handle:
            fields = line.strip().split(",")
            if len(fields) > 2 and fields[0].isdigit():
                counters[fields[2]] = int(fields[0])
    return counters


def bootstrap(ratios, iterations=20000, seed=765):
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
        perf = {}
        environment = None
        for slot in SLOTS:
            path = f"{RAW}/{group}-{slot}.json"
            if not os.path.exists(path):
                continue
            data, rows = load(path)
            processes[slot] = rows
            perf[slot] = load_perf(f"{RAW}/{group}-{slot}.perf")
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
            instructions = {s: perf[s].get("instructions:u") for s in per}
            cycles = {s: perf[s].get("cycles:u") for s in per}
            instruction_ratios = [
                instructions[bs] / instructions[as_]
                for as_, bs in PAIRS
                if instructions.get(as_) and instructions.get(bs)
            ]
            cycle_ratios = [
                cycles[bs] / cycles[as_] for as_, bs in PAIRS if cycles.get(as_) and cycles.get(bs)
            ]
            entry = {
                "case": key[0],
                "shape": key[1],
                "processes": {
                    slot: dict(per[slot], instructions_u=instructions[slot], cycles_u=cycles[slot])
                    for slot in sorted(per)
                },
                "before_median_p50_ns": statistics.median(a) if a else None,
                "after_median_p50_ns": statistics.median(b) if b else None,
                "paired_p50_ratios": ratios,
                "paired_mean_ratios": mean_ratios,
                "mean_paired_p50_ratio": sum(ratios) / len(ratios) if ratios else None,
                "bootstrap95_mean_paired_p50_ratio": [low, high],
                "before_p50_drift": (max(a) / min(a) - 1) if a else None,
                "after_p50_drift": (max(b) / min(b) - 1) if b else None,
                "paired_process_instruction_ratios": instruction_ratios,
                "paired_process_cycle_ratios": cycle_ratios,
                "outputs_identical": len({per[s]["output_sha256"] for s in per}) == 1
                and len({per[s]["accepted_bytes"] for s in per}) == 1,
                "corpus_identical": len({per[s]["archive_sha256"] for s in per}) == 1,
            }
            if ratios and max(ratios) > 1.05:
                flags.append(
                    {"group": group, "case": key[0], "shape": key[1], "max_paired_p50_ratio": max(ratios)}
                )
            entries.append(entry)
        report["groups"][group] = {"environment": environment, "entries": entries}
    report["regression_flags_over_5pct"] = flags
    with open(OUT, "w") as handle:
        json.dump(report, handle, indent=1)
    print(
        f"{'case':30s} {'shape':10s} {'before p50':>11s} {'after p50':>11s} {'ratio':>7s} "
        f"{'CI95':>17s} {'driftA':>7s} {'driftB':>7s} {'instr':>7s} {'cycles':>7s} same-out"
    )
    for group, value in report["groups"].items():
        for entry in value["entries"]:
            before = entry["before_median_p50_ns"] / 1e6
            after = entry["after_median_p50_ns"] / 1e6
            low, high = entry["bootstrap95_mean_paired_p50_ratio"]
            instr = entry["paired_process_instruction_ratios"]
            cyc = entry["paired_process_cycle_ratios"]
            instr_mean = sum(instr) / len(instr) if instr else float("nan")
            cyc_mean = sum(cyc) / len(cyc) if cyc else float("nan")
            print(
                f"{entry['case']:30s} {entry['shape']:10s} {before:11.4f} {after:11.4f} "
                f"{entry['mean_paired_p50_ratio']:7.4f} [{low:.4f},{high:.4f}] "
                f"{entry['before_p50_drift']*100:6.2f}% {entry['after_p50_drift']*100:6.2f}% "
                f"{instr_mean:7.4f} {cyc_mean:7.4f} "
                f"{entry['outputs_identical'] and entry['corpus_identical']}"
            )
    print("flags:", flags)


if __name__ == "__main__":
    main()
