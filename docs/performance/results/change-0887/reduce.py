#!/usr/bin/env python3
"""Reduce retained 0887 children into paired descriptive evidence.

All ratios are formed from whole-child statistics.  Samples are never pooled
across children, and allocator counters remain a separate primary workload
observation.  The reducer reopens raw reports and receipts so a cached
summary cannot stand in for the captured evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import random
import statistics
from typing import Any

import experiment


SCHEMA = "litchi.xlsx-mce-inline4-reduction-0887.v1"
BOOTSTRAP_DRAWS = 10_000
BOOTSTRAP_LOW_INDEX = 249
BOOTSTRAP_HIGH_INDEX = 9_749
CONTROL_CASES = experiment.CASES[1:]


def _read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _bootstrap(ratios: list[float]) -> dict[str, Any]:
    require(ratios, "cannot bootstrap an empty ratio vector")
    rng = random.Random(experiment.BOOTSTRAP_SEED)
    medians = []
    for _ in range(BOOTSTRAP_DRAWS):
        draw = [ratios[rng.randrange(len(ratios))] for _ in ratios]
        medians.append(statistics.median(draw))
    ordered = sorted(medians)
    return {
        "seed": experiment.BOOTSTRAP_SEED,
        "draws": BOOTSTRAP_DRAWS,
        "low_index": BOOTSTRAP_LOW_INDEX,
        "high_index": BOOTSTRAP_HIGH_INDEX,
        "median_ratio": statistics.median(ratios),
        "ci95_low": ordered[BOOTSTRAP_LOW_INDEX],
        "ci95_high": ordered[BOOTSTRAP_HIGH_INDEX],
        "ratios": ratios,
    }


def _artifact_sha(receipt: dict[str, Any], name: str, path: Path) -> None:
    record = receipt.get("artifacts", {}).get(name)
    require(isinstance(record, dict) and record.get("sha256") == _sha(path), f"{name} descriptor mismatch")


def _allocation(report: dict[str, Any]) -> dict[str, list[int]]:
    result_rows = report.get("results")
    require(type(result_rows) is list and len(result_rows) == 1, "allocator report result count mismatch")
    allocation = result_rows[0].get("operation_metrics", {}).get("allocation")
    require(isinstance(allocation, dict) and allocation.get("status") == "measured", "allocator metrics are missing")
    result = {}
    for key in ("allocation_calls", "allocated_bytes", "region_peak_live_bytes"):
        vector = allocation.get(key, {}).get("values")
        require(type(vector) is list and vector and all(type(value) is int and value > 0 for value in vector), f"allocator vector {key} is invalid")
        result[key] = vector
    failed = allocation.get("failed_allocation_calls", {}).get("values")
    require(failed == [0] * len(result["allocation_calls"]), "allocator failed allocation calls are nonzero")
    return result


def _load_rows(plan_path: Path, experiment_root: Path) -> list[dict[str, Any]]:
    plan = experiment.load_plan(plan_path)
    root = experiment_root.resolve()
    receipt_path = root / "experiment.json"
    require(receipt_path.is_file(), "experiment.json is missing")
    experiment_receipt = _read(receipt_path)
    require(experiment_receipt.get("schema") == experiment.EXPERIMENT_SCHEMA and experiment_receipt.get("status") == "pass", "experiment receipt is not passing")
    require(experiment_receipt.get("plan_sha256") == _sha(plan_path), "experiment plan binding mismatch")
    rows = experiment._schedule()
    require(experiment_receipt.get("children") == len(rows) == experiment.TOTAL_CHILDREN, "experiment population mismatch")
    expected = experiment._load_expected_corpora()
    binaries = {(role, mode): experiment._binary(plan, role, mode) for role in ("before", "after") for mode in ("native", "allocator")}
    observed = []
    previous_finished = None
    for row in rows:
        folder = root / "runs" / f"{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}"
        receipt_path = folder / "receipt.json"
        require(receipt_path.is_file(), f"missing child receipt: {row['ordinal']}")
        receipt = _read(receipt_path)
        require(receipt.get("schema") == experiment.RECEIPT_SCHEMA, f"child receipt schema mismatch: {row['ordinal']}")
        require(receipt.get("row") == row and receipt.get("exit_code") == 0, f"child receipt identity/status mismatch: {row['ordinal']}")
        require(receipt.get("plan_sha256") == _sha(plan_path), f"child plan binding mismatch: {row['ordinal']}")
        mode = "native" if row["lane"] == "native" else "allocator"
        binary = binaries[(row["role"], mode)]
        require(receipt.get("binary") == binary, f"child binary binding mismatch: {row['ordinal']}")
        command_path = folder / "command.json"
        require(command_path.is_file(), f"child command missing: {row['ordinal']}")
        command = _read(command_path)
        require(command.get("argv") == receipt.get("argv") and command.get("row") == row, f"child command mismatch: {row['ordinal']}")
        require(receipt.get("command", {}).get("sha256") == _sha(command_path), f"child command descriptor mismatch: {row['ordinal']}")
        for key in ("cwd", "environment", "binary", "plan_sha256", "started_utc", "finished_utc", "exit_code"):
            require(command.get(key) == receipt.get(key), f"child command/receipt {key} mismatch: {row['ordinal']}")
        for name in ("stdout.txt", "stderr.txt", "report.json", "peak-rss-kib.txt"):
            path = folder / name
            require(path.is_file(), f"child {name} missing: {row['ordinal']}")
            _artifact_sha(receipt, name, path)
        report_path = folder / "report.json"
        producer_path = folder / "producer.json"
        report = _read(report_path)
        producer = _read(producer_path) if producer_path.is_file() else None
        if producer_path.is_file():
            _artifact_sha(receipt, "producer.json", producer_path)
        checked = experiment._validate_child(report, producer, row["case"], mode, row["warmup"], row["samples"], expected[row["case"]], binary)
        peak = experiment._read_peak(folder / "peak-rss-kib.txt")
        started = experiment._parse_time(receipt["started_utc"])
        finished = experiment._parse_time(receipt["finished_utc"])
        require(finished >= started, f"child interval invalid: {row['ordinal']}")
        if previous_finished is not None:
            require(started >= previous_finished, f"child order is not sequential: {row['ordinal']}")
        previous_finished = finished
        observed.append({"row": row, "receipt": receipt, "report": report, "producer": producer, "checked": checked, "peak_rss_kib": peak, "folder": str(folder)})
    return observed


def _pairs(rows: list[dict[str, Any]], lane: str, case: str) -> list[dict[str, Any]]:
    selected = [item for item in rows if item["row"]["lane"] == lane and item["row"]["case"] == case]
    groups: dict[str, dict[str, dict[str, Any]]] = {}
    for item in selected:
        pair = item["row"]["pair"]
        role = item["row"]["role"]
        require(role in ("before", "after"), f"unknown pair role {role}")
        require(role not in groups.setdefault(pair, {}), f"duplicate pair role {pair}/{role}")
        groups[pair][role] = item
    expected_count = experiment.NATIVE_BLOCKS if lane == "native" else experiment.ALLOCATOR_BLOCKS
    require(len(groups) == expected_count, f"{lane}/{case} pair count mismatch")
    pairs = []
    for pair_name in sorted(groups):
        pair = groups[pair_name]
        require(set(pair) == {"before", "after"}, f"incomplete pair {pair_name}")
        pairs.append({"pair": pair_name, "before": pair["before"], "after": pair["after"]})
    return pairs


def _elapsed_metric(item: dict[str, Any], metric: str) -> int:
    return item["checked"]["stats"][metric]


def _native_summary(pairs: list[dict[str, Any]], metric: str) -> dict[str, Any]:
    values = []
    for pair in pairs:
        before = _elapsed_metric(pair["before"], metric)
        after = _elapsed_metric(pair["after"], metric)
        require(before > 0 and after > 0, f"nonpositive elapsed metric {metric}")
        values.append({"pair": pair["pair"], "before": before, "after": after, "ratio": after / before})
    ratios = [item["ratio"] for item in values]
    result = _bootstrap(ratios)
    result["pairs"] = values
    return result


def _rss_summary(pairs: list[dict[str, Any]]) -> dict[str, Any]:
    values = []
    for pair in pairs:
        before = pair["before"]["peak_rss_kib"]
        after = pair["after"]["peak_rss_kib"]
        require(before > 0 and after > 0, "nonpositive peak RSS")
        values.append({"pair": pair["pair"], "before": before, "after": after, "ratio": after / before})
    result = _bootstrap([item["ratio"] for item in values])
    result["pairs"] = values
    return result


def _spread(rows: list[dict[str, Any]], lane: str, case: str, role: str) -> dict[str, Any]:
    selected = [item for item in rows if item["row"]["lane"] == lane and item["row"]["case"] == case and item["row"]["role"] == role]
    values = [_elapsed_metric(item, "p50") for item in selected]
    require(values, f"missing spread values for {lane}/{case}/{role}")
    ratio = max(values) / min(values)
    return {"lane": lane, "case": case, "role": role, "values": values, "max_over_min": ratio, "flag": ratio > 1.20}


def _allocator_summary(pairs: list[dict[str, Any]]) -> dict[str, Any]:
    result = {}
    for metric in ("allocation_calls", "allocated_bytes", "region_peak_live_bytes"):
        values = []
        for pair in pairs:
            before_vector = _allocation(pair["before"]["report"])[metric]
            after_vector = _allocation(pair["after"]["report"])[metric]
            before = statistics.median(before_vector)
            after = statistics.median(after_vector)
            values.append({"pair": pair["pair"], "before": before, "after": after, "ratio": after / before})
        summary = _bootstrap([item["ratio"] for item in values])
        summary["pairs"] = values
        result[metric] = summary
    return result


def reduce(plan_path: Path, experiment_root: Path, output: Path | None = None) -> dict[str, Any]:
    rows = _load_rows(plan_path, experiment_root)
    native: dict[str, Any] = {}
    for case in experiment.CASES:
        pairs = _pairs(rows, "native", case)
        native[case] = {
            "p50_ns": _native_summary(pairs, "p50"),
            "p95_ns": _native_summary(pairs, "p95"),
            "p99_ns": _native_summary(pairs, "p99"),
            "peak_rss_kib": _rss_summary(pairs),
        }
    allocator_pairs = _pairs(rows, "allocator", experiment.PRIMARY_CASE)
    allocator = _allocator_summary(allocator_pairs)
    spread_flags = []
    for case in experiment.CASES:
        for role in ("before", "after"):
            spread = _spread(rows, "native", case, role)
            if spread["flag"]:
                spread_flags.append(spread)
    control_flags = []
    for case in CONTROL_CASES:
        for metric in ("p50_ns", "p95_ns", "p99_ns", "peak_rss_kib"):
            point = native[case][metric]
            if any(pair["ratio"] > 1.05 for pair in point["pairs"]):
                control_flags.append({"case": case, "metric": metric, "reason": "control pair ratio above 1.05", "summary": point})
    primary = native[experiment.PRIMARY_CASE]["p50_ns"]
    primary_flags = []
    for metric in ("p50_ns", "p95_ns", "p99_ns", "peak_rss_kib"):
        point = native[experiment.PRIMARY_CASE][metric]
        if any(pair["ratio"] > 1.05 for pair in point["pairs"]):
            primary_flags.append({"case": experiment.PRIMARY_CASE, "metric": metric, "reason": "primary pair ratio above 1.05", "summary": point})
    primary_gate = {"median_ratio_at_most": 0.98, "ci95_upper_below": 1.0, "median_ratio": primary["median_ratio"], "ci95_high": primary["ci95_high"], "pass": primary["median_ratio"] <= 0.98 and primary["ci95_high"] < 1.0}
    result = {
        "schema": SCHEMA,
        "status": "pass",
        "plan": str(plan_path.resolve()),
        "plan_sha256": _sha(plan_path),
        "experiment": str((experiment_root / "experiment.json").resolve()),
        "experiment_sha256": _sha(experiment_root / "experiment.json"),
        "native": native,
        "allocator_primary": allocator,
        "primary_native_p50_gate": primary_gate,
        "primary_review_flags": primary_flags,
        "control_review_flags": control_flags,
        "native_p50_spread_flags": spread_flags,
        "allocator_adoption": {"status": "not_evaluated", "reason": "allocator counters are a separate primary observation; no automatic adoption"},
        "bootstrap": {"seed": experiment.BOOTSTRAP_SEED, "draws": BOOTSTRAP_DRAWS, "ci95_indices": [BOOTSTRAP_LOW_INDEX, BOOTSTRAP_HIGH_INDEX], "unit": "whole paired child ratios"},
        "pooling": "none; each ratio uses one child aggregate statistic and pairs are summarized by median",
        "performance_claim": None,
    }
    if output is not None:
        with output.open("x", encoding="utf-8") as handle:
            json.dump(result, handle, indent=2, sort_keys=True)
            handle.write("\n")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--experiment-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = reduce(args.plan.resolve(), args.experiment_root.resolve(), args.output.resolve())
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
