#!/usr/bin/env python3
"""Analyze the change-0558 evidence packet.

Reads the retained ABBA reports and strace summaries, validates that the
deterministic counters and semantic projections agree, and emits one JSON
result. It computes nothing that is not present in the retained children.
"""

from __future__ import annotations

import argparse
import json
import pathlib
import re
import statistics
import sys

MODES = ("file-source", "owned-readat")
OPERATIONS = ("open", "list", "one-cell")
REPEATS = ("R1", "R2")
STAT_KEYS = ("p50", "mean", "p95", "p99")


def quantile(values: list[int], fraction: float) -> float:
    ordered = sorted(values)
    if not ordered:
        raise ValueError("no samples")
    if len(ordered) == 1:
        return float(ordered[0])
    position = fraction * (len(ordered) - 1)
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    weight = position - low
    return ordered[low] * (1.0 - weight) + ordered[high] * weight


def summarize(samples: list[int]) -> dict:
    return {
        "samples": len(samples),
        "p50": quantile(samples, 0.50),
        "mean": statistics.fmean(samples),
        "p95": quantile(samples, 0.95),
        "p99": quantile(samples, 0.99),
        "min": float(min(samples)),
    }


def load_cell(path: pathlib.Path) -> dict:
    report = json.loads(path.read_text())
    records = report["records"]
    metrics = [record["metrics"] for record in records]
    first = metrics[0]
    for other in metrics[1:]:
        for key in ("version_calls", "read_calls", "read_bytes", "len_calls"):
            if other[key] != first[key]:
                raise SystemExit(f"{path.name}: {key} varies across samples in one child")
    return {
        "elapsed": summarize(report["elapsed_samples_ns"]),
        "counters": {key: first[key] for key in ("version_calls", "read_calls", "read_bytes", "len_calls")},
        "oracle": report["semantic_oracle"]["source_implementation_projection"],
        "mode": report["mode"],
        "operation": report["operation"],
        "binary_sha256": report["binary"]["sha256"],
        "input_sha256": report["input"]["sha256"],
    }


def percent(new: float, old: float) -> float:
    return (new - old) / old * 100.0


def parse_strace(path: pathlib.Path) -> dict:
    counts: dict[str, int] = {}
    for line in path.read_text().splitlines():
        fields = line.split()
        if len(fields) < 4 or fields[0].startswith(("%", "-")):
            continue
        name = fields[-1]
        if not re.fullmatch(r"[a-z0-9_]+", name):
            continue
        try:
            calls = int(fields[3])
        except ValueError:
            continue
        counts[name] = counts.get(name, 0) + calls
    return counts


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", default="docs/performance/results/change-0558")
    parser.add_argument("--output", default="docs/performance/results/change-0558/analysis.json")
    args = parser.parse_args()

    root = pathlib.Path(args.packet)
    latency = root / "latency"
    cells: dict[tuple[str, str, str, str], dict] = {}
    for repeat in REPEATS:
        for leg, stage in (("A1", "baseline"), ("B1", "candidate"), ("B2", "candidate"), ("A2", "baseline")):
            for mode in MODES:
                for operation in OPERATIONS:
                    path = latency / f"{repeat}-{leg}-{stage}-{mode}-{operation}.json"
                    cells[(repeat, leg, mode, operation)] = load_cell(path)

    # Every child must agree on the semantic projection, the input, and the
    # logical read work. Those are correctness gates, not statistics.
    oracles_per_cell: dict[tuple[str, str], set[str]] = {}
    for key, cell in cells.items():
        oracles_per_cell.setdefault((key[2], key[3]), set()).add(
            json.dumps(cell["oracle"], sort_keys=True)
        )
    inputs = {cell["input_sha256"] for cell in cells.values()}
    read_work = {}
    for key, cell in cells.items():
        read_work.setdefault((key[2], key[3]), set()).add(
            (cell["counters"]["read_calls"], cell["counters"]["read_bytes"])
        )
    binaries = {}
    for key, cell in cells.items():
        stage = "baseline" if key[1].startswith("A") else "candidate"
        binaries.setdefault(stage, set()).add(cell["binary_sha256"])

    gates = {
        "one_semantic_projection_per_cell": all(
            len(values) == 1 for values in oracles_per_cell.values()
        ),
        "one_input_identity": len(inputs) == 1,
        "read_work_identical_per_cell": all(len(values) == 1 for values in read_work.values()),
        "one_binary_per_stage": all(len(values) == 1 for values in binaries.values()),
        "stages_differ": binaries.get("baseline") != binaries.get("candidate"),
    }

    counters = []
    for mode in MODES:
        for operation in OPERATIONS:
            base = cells[("R1", "A1", mode, operation)]["counters"]
            cand = cells[("R1", "B1", mode, operation)]["counters"]
            counters.append(
                {
                    "mode": mode,
                    "operation": operation,
                    "baseline_version_calls": base["version_calls"],
                    "candidate_version_calls": cand["version_calls"],
                    "version_calls_change_percent": percent(cand["version_calls"], base["version_calls"]),
                    "read_calls": base["read_calls"],
                    "read_calls_unchanged": base["read_calls"] == cand["read_calls"],
                    "read_bytes": base["read_bytes"],
                    "read_bytes_unchanged": base["read_bytes"] == cand["read_bytes"],
                }
            )

    comparisons = []
    for repeat in REPEATS:
        for mode in MODES:
            for operation in OPERATIONS:
                row = {"repeat": repeat, "mode": mode, "operation": operation, "directions": {}}
                for direction, (a_leg, b_leg) in (("first", ("A1", "B1")), ("second", ("A2", "B2"))):
                    a = cells[(repeat, a_leg, mode, operation)]["elapsed"]
                    b = cells[(repeat, b_leg, mode, operation)]["elapsed"]
                    row["directions"][direction] = {
                        stat: {
                            "baseline_ns": a[stat],
                            "candidate_ns": b[stat],
                            "change_percent": percent(b[stat], a[stat]),
                        }
                        for stat in STAT_KEYS
                    }
                row["improves_both_directions"] = {
                    stat: row["directions"]["first"][stat]["change_percent"] < 0.0
                    and row["directions"]["second"][stat]["change_percent"] < 0.0
                    for stat in STAT_KEYS
                }
                comparisons.append(row)

    stability = []
    for repeat in REPEATS:
        for stage, legs in (("baseline", ("A1", "A2")), ("candidate", ("B1", "B2"))):
            for mode in MODES:
                for operation in OPERATIONS:
                    first = cells[(repeat, legs[0], mode, operation)]["elapsed"]
                    second = cells[(repeat, legs[1], mode, operation)]["elapsed"]
                    stability.append(
                        {
                            "repeat": repeat,
                            "stage": stage,
                            "mode": mode,
                            "operation": operation,
                            "same_binary_drift_percent": {
                                stat: percent(second[stat], first[stat]) for stat in STAT_KEYS
                            },
                        }
                    )
    cross_repeat = []
    for stage, leg in (("baseline", "A1"), ("candidate", "B1")):
        for mode in MODES:
            for operation in OPERATIONS:
                first = cells[("R1", leg, mode, operation)]["elapsed"]
                second = cells[("R2", leg, mode, operation)]["elapsed"]
                cross_repeat.append(
                    {
                        "stage": stage,
                        "mode": mode,
                        "operation": operation,
                        "drift_percent": {stat: percent(second[stat], first[stat]) for stat in STAT_KEYS},
                    }
                )

    syscalls = []
    syscall_dir = root / "syscalls"
    if syscall_dir.is_dir():
        for mode in MODES:
            for operation in OPERATIONS:
                base_path = syscall_dir / f"baseline-{mode}-{operation}.strace.txt"
                cand_path = syscall_dir / f"candidate-{mode}-{operation}.strace.txt"
                if not (base_path.exists() and cand_path.exists()):
                    continue
                base = parse_strace(base_path)
                cand = parse_strace(cand_path)
                syscalls.append(
                    {
                        "mode": mode,
                        "operation": operation,
                        "baseline": {k: base.get(k, 0) for k in ("statx", "pread64", "read", "openat")},
                        "candidate": {k: cand.get(k, 0) for k in ("statx", "pread64", "read", "openat")},
                        "statx_change_percent": percent(cand.get("statx", 0), base["statx"]) if base.get("statx") else None,
                        "pread64_unchanged": base.get("pread64") == cand.get("pread64"),
                    }
                )

    result = {
        "schema_version": 1,
        "analysis_kind": "litchi-perf-change-0558-analysis",
        "gates": gates,
        "deterministic_counters": counters,
        "syscalls": syscalls,
        "latency_comparisons": comparisons,
        "same_binary_stability": stability,
        "cross_repeat_drift": cross_repeat,
    }
    pathlib.Path(args.output).write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps({"gates": gates}, indent=2))
    return 0 if all(gates.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
