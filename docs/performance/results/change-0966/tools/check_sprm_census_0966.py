#!/usr/bin/env python3
"""Independent structural and repeat audit for the 0966 DOC SPRM census.

This checker never opens a DOC and never invokes the probe. It validates only
the bounded JSON reports emitted by fresh children. It deliberately does not
turn aggregate counters into a timing or allocation claim.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


SCHEMA = "litchi.performance.0966.doc_sprm_census.v1"
CASES = {
    "docnohf": (26112, "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48"),
    "docfloat": (335360, "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d"),
}
EDIT_COUNTS = {1, 2, 8, 32}
RETENTIONS = {"zero", "default_8mib", "release_8mib"}
PHASES = ("snapshot_open", "edit_open", "staged_replacement", "final_commit", "output_copy", "other")
HISTOGRAM_KEYS = (
    "input_length_histogram",
    "output_count_histogram",
)
CORRECTNESS_KEYS = (
    "all_stages_change_text",
    "exact_cross_route_outputs",
    "final_paragraph_semantic_readback",
    "final_public_reader_reopen",
    "exact_noop_source",
    "noop_shares_source_allocation",
    "forward_patch_exact_output",
    "inverse_patch_exact_source",
    "foreign_patch_refused",
    "stale_patch_refused",
    "same_value_preserves_retention",
    "operation_limit_refusal_atomic",
    "retained_capacities_within_ceiling",
    "zero_route_clears_retention",
    "release_route_clears_every_stage",
    "untouched_streams_preserved",
    "directory_metadata_preserved",
    "semantic_oracles_passed",
    "semantic_change_count_matches_edit_count",
    "allocator_retained_matches_output_capacity",
)
BUCKET_COUNTERS = (
    "calls",
    "successes",
    "errors",
    "input_empty",
    "input_nonempty",
    "input_bytes_total",
    "output_count_total",
    "output_capacity_total",
    "logical_vec_capacity_bytes",
)
BOUND_COUNTERS = (
    "input_bytes_min",
    "input_bytes_max",
    "output_capacity_min",
    "output_capacity_max",
)


def fail(message: str) -> None:
    raise SystemExit(f"FAIL: {message}")


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def artifact(path: Path) -> dict[str, Any]:
    try:
        with path.open(encoding="utf-8") as handle:
            value = json.load(handle)
    except (OSError, json.JSONDecodeError) as error:
        fail(f"cannot read {path}: {error}")
    require(isinstance(value, dict), f"{path} is not a JSON object")
    return value


def sum_values(values: Any, label: str, expected_length: int) -> int:
    require(isinstance(values, list), f"{label} is not an array")
    require(len(values) == expected_length, f"{label} has length {len(values)}, expected {expected_length}")
    require(all(type(value) is int and value >= 0 for value in values), f"{label} has invalid counters")
    return sum(values)


def check_bucket(bucket: dict[str, Any], label: str, sprm_size: int) -> None:
    for key in BUCKET_COUNTERS + BOUND_COUNTERS:
        require(type(bucket.get(key)) is int and bucket[key] >= 0, f"{label}.{key}: invalid counter")
    input_histogram_total = sum_values(
        bucket.get("input_length_histogram"), f"{label}.input_length_histogram", 10
    )
    output_histogram_total = sum_values(
        bucket.get("output_count_histogram"), f"{label}.output_count_histogram", 10
    )
    error_histogram_total = sum_values(bucket.get("errors_by_kind"), f"{label}.errors_by_kind", 5)
    require(bucket.get("calls") == bucket.get("successes", 0) + bucket.get("errors", 0), f"{label}: call conservation")
    require(bucket.get("input_empty", 0) + bucket.get("input_nonempty", 0) == bucket.get("calls"), f"{label}: empty/nonempty conservation")
    require(input_histogram_total == bucket.get("calls"), f"{label}: input histogram conservation")
    require(error_histogram_total == bucket.get("errors"), f"{label}: error-kind conservation")
    successes = bucket.get("successes")
    require(output_histogram_total == successes, f"{label}: output histogram conservation")
    require(bucket["logical_vec_capacity_bytes"] == bucket["output_capacity_total"] * sprm_size, f"{label}: logical capacity lower bound")
    for minimum, maximum, name in (("input_bytes_min", "input_bytes_max", "input bytes"), ("output_capacity_min", "output_capacity_max", "output capacity")):
        if bucket["calls"] == 0 or (name == "output capacity" and bucket["successes"] == 0):
            require(bucket[minimum] == 0 and bucket[maximum] == 0, f"{label}: empty {name} bounds")
        else:
            require(0 <= bucket[minimum] <= bucket[maximum], f"{label}: invalid {name} bounds")
    if bucket["calls"]:
        require(bucket["input_bytes_min"] <= bucket["input_bytes_max"], f"{label}: invalid input byte bounds")


def canonical_report(report: dict[str, Any]) -> dict[str, Any]:
    sample = report["samples"][0]
    return {
        "source_sha256": report["source_sha256"],
        "source_bytes": report["source_bytes"],
        "output_sha256": report["output_sha256"],
        "output_bytes": report["output_bytes"],
        "output_capacity": report["output_capacity"],
        "exact_output_equal_reference": report["exact_output_equal_reference"],
        "correctness": report["correctness"],
        "sprm_census": sample["sprm_census"],
    }


def check_report(path: Path) -> tuple[tuple[str, int, str], dict[str, Any]]:
    report = artifact(path)
    require(report.get("schema") == SCHEMA, f"{path}: schema")
    require(report.get("change") == "0966", f"{path}: change")
    require(report.get("timing_claim") is False, f"{path}: timing claim must be false")
    require(report.get("warmups") == 0 and report.get("samples_requested") == 1, f"{path}: one fresh sample")
    require(isinstance(report.get("samples"), list) and len(report["samples"]) == 1, f"{path}: sample count")
    case = report.get("case")
    edits = report.get("edits")
    retention = report.get("retention")
    require(case in CASES and type(edits) is int and edits in EDIT_COUNTS and retention in RETENTIONS, f"{path}: case cell")
    expected_bytes, expected_sha256 = CASES[case]
    require(type(report.get("source_bytes")) is int and report.get("source_bytes") == expected_bytes and report.get("source_sha256") == expected_sha256, f"{path}: fixture identity")
    require(report.get("exact_output_equal_reference") is True, f"{path}: output parity")
    for key in ("output_bytes", "output_capacity"):
        require(type(report.get(key)) is int and report[key] >= 0, f"{path}: top {key}")
    correctness = report.get("correctness")
    require(isinstance(correctness, dict), f"{path}: correctness object")
    require(set(correctness) == set(CORRECTNESS_KEYS), f"{path}: correctness key set")
    require(all(type(correctness.get(key)) is bool and correctness[key] is True for key in CORRECTNESS_KEYS), f"{path}: correctness gate")
    sample = report["samples"][0]
    for key in ("output_bytes", "output_capacity"):
        require(type(sample.get(key)) is int and sample[key] >= 0, f"{path}: sample {key}")
    for key in ("output_sha256", "output_bytes", "output_capacity"):
        require(sample.get(key) == report.get(key), f"{path}: sample/top {key} mismatch")
    require(sample.get("output_sha256") == report.get("reference_output_sha256"), f"{path}: sample/reference hash mismatch")
    require(sample.get("output_bytes") == report.get("reference_output_bytes"), f"{path}: sample/reference byte count mismatch")
    require(type(report.get("reference_output_bytes")) is int and report["reference_output_bytes"] >= 0, f"{path}: reference byte count")
    route_outputs = report.get("route_outputs")
    require(isinstance(route_outputs, list) and len(route_outputs) == 3, f"{path}: route output shape")
    require(isinstance(route_outputs[0], dict), f"{path}: reference route shape")
    require(report.get("reference_output_sha256") == route_outputs[0].get("output_sha256"), f"{path}: reference hash mismatch")
    require(report.get("reference_output_bytes") == route_outputs[0].get("output_bytes"), f"{path}: reference byte count mismatch")
    census = sample.get("sprm_census")
    require(isinstance(census, dict) and census.get("schema_version") == 1, f"{path}: census object")
    require(census.get("saturated") is False, f"{path}: saturated counters")
    sprm_size = census.get("sprm_size_bytes")
    require(type(sprm_size) is int and sprm_size > 0, f"{path}: SPRM size")
    total = census.get("total")
    require(isinstance(total, dict), f"{path}: total bucket")
    phase_buckets = []
    for phase in PHASES:
        bucket = census.get(phase)
        require(isinstance(bucket, dict), f"{path}: missing phase {phase}")
        check_bucket(bucket, f"{path}:{phase}", sprm_size)
        phase_buckets.append(bucket)
    check_bucket(total, f"{path}:total", sprm_size)
    scalar_fields = (
        "calls", "successes", "errors", "input_empty", "input_nonempty",
        "input_bytes_total", "output_count_total", "output_capacity_total",
        "logical_vec_capacity_bytes",
    )
    for field in scalar_fields:
        require(sum(bucket[field] for bucket in phase_buckets) == total[field], f"{path}: phase conservation for {field}")
    for field in HISTOGRAM_KEYS + ("errors_by_kind",):
        require([sum(bucket[field][index] for bucket in phase_buckets) for index in range(len(total[field]))] == total[field], f"{path}: phase conservation for {field}")
    for field in ("input_bytes_min", "input_bytes_max", "output_capacity_min", "output_capacity_max"):
        # Bounds are not additive; the independent checker checks only their shape.
        require(isinstance(total[field], int) and total[field] >= 0, f"{path}: total {field}")
    require(census.get("other", {}).get("calls") == 0, f"{path}: unexpected unclassified phase calls")
    key = (case, edits, retention)
    return key, canonical_report(report)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("reports", type=Path, help="directory containing exactly 48 child JSON reports")
    parser.add_argument("--expected-count", type=int, default=48)
    args = parser.parse_args()
    paths = sorted(args.reports.glob("*.json"))
    require(len(paths) == args.expected_count, f"expected {args.expected_count} reports, found {len(paths)}")
    groups: dict[tuple[str, int, str], list[dict[str, Any]]] = {}
    for path in paths:
        key, identity = check_report(path)
        groups.setdefault(key, []).append(identity)
    expected_cells = {(case, edits, retention) for case in CASES for edits in EDIT_COUNTS for retention in RETENTIONS}
    require(set(groups) == expected_cells, "the 24-cell matrix is incomplete or has an extra cell")
    for key, identities in sorted(groups.items()):
        require(len(identities) == 2, f"{key}: expected two fresh-child repeats")
        require(identities[0] == identities[1], f"{key}: repeat reports differ")
    print(json.dumps({"schema": SCHEMA, "reports": len(paths), "cells": len(groups), "repeat_equal": True}, sort_keys=True))


if __name__ == "__main__":
    main()
