"""Admission checks for single-case ordinary-save allocation report vectors.

This reads existing JSON only. It does not run a workload, summarize timings,
validate artifact semantics, or establish comparability between two reports.
Capture drivers must additionally bind the source, binary, corpus, outputs,
and protocol. Schema 1 and serialized_region_peak_v3 are intentionally pinned.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from typing import Any


ALLOCATION_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
SCOPE = "operation_global_system_allocator"
ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"
REVISION = "serialized_region_peak_v3"
U64_MAX = (1 << 64) - 1


class AllocationSchemaError(ValueError):
    """A report cannot be admitted under the requested capture contract."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AllocationSchemaError(message)


def object_at(value: Any, path: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: expected object")
    return value


def uint(value: Any, path: str, *, positive: bool = False) -> int:
    require(type(value) is int and int(positive) <= value <= U64_MAX,
            f"{path}: expected {'positive' if positive else 'nonnegative'} u64 integer")
    return value


def validate_allocation(value: Any, *, samples: int, measured: bool) -> None:
    """Check every v3 metric and each sample independently, without mutation.

    Raw counters are unsigned. Net-live differences are signed and may be
    negative; neither allocation count equality nor nonnegative net-live is
    required. The process lifetime peak is distinct from the region peak.
    Measured reports with failed allocation calls are excluded from this
    successful-operation admission contract.
    """
    uint(samples, "samples", positive=True)
    require(type(measured) is bool, "measured: expected bool")
    allocation = object_at(value, "allocation")
    require(set(allocation) == {"status", "scope", *ALLOCATION_FIELDS},
            "allocation: missing or unknown fields")
    status = "measured" if measured else "unavailable"
    require(allocation["status"] == status, f"allocation.status: expected {status}")
    require(allocation["scope"] == SCOPE, "allocation.scope: unexpected scope")
    vectors: dict[str, list[int]] = {}
    for name in ALLOCATION_FIELDS:
        path = f"allocation.{name}"
        metric = object_at(allocation[name], path)
        require(set(metric) == ({"status", "scope", "values"} if measured else {"status", "scope"}),
                f"{path}: missing or unknown fields")
        require(metric["status"] == status, f"{path}.status: expected {status}")
        require(metric["scope"] == SCOPE, f"{path}.scope: unexpected scope")
        if measured:
            values = metric["values"]
            require(isinstance(values, list) and len(values) == samples,
                    f"{path}.values: expected {samples} samples")
            vectors[name] = [uint(x, f"{path}.values[{i}]") for i, x in enumerate(values)]
    if not measured:
        return
    for index in range(samples):
        row = {name: values[index] for name, values in vectors.items()}
        label = f"allocation sample {index}"
        require(row["failed_allocation_calls"] == 0, f"{label}: failed allocation calls")
        require(row["live_bytes_after"] - row["live_bytes_before"] ==
                row["allocated_bytes"] - row["deallocated_bytes"],
                f"{label}: signed live-byte conservation failed")
        require(row["peak_live_bytes_before"] >= row["live_bytes_before"],
                f"{label}: lifetime peak before is below live bytes before")
        require(row["peak_live_bytes_after"] >= row["peak_live_bytes_before"],
                f"{label}: lifetime peak decreased")
        require(max(row["live_bytes_before"], row["live_bytes_after"]) <=
                row["region_peak_live_bytes"] <= row["peak_live_bytes_after"],
                f"{label}: region peak is outside live/lifetime bounds")


def validate_report(report: Any, *, expected_case: str, expected_samples: int,
                    expected_warmup: int, mode: str) -> None:
    """Validate a single-case capture's identity, alignment, and allocations.

    The caller supplies the frozen case/sample/warmup/mode contract. Those
    expectations must not be learned from the untrusted report being checked.
    Process counters, artifact oracles, and statistical summaries need their
    own validators; this function deliberately makes no claims about them.
    """
    require(mode in ("native", "observer"), "mode: expected native or observer")
    require(isinstance(expected_case, str) and bool(expected_case), "expected_case: empty or invalid")
    uint(expected_samples, "expected_samples", positive=True)
    uint(expected_warmup, "expected_warmup")
    report = object_at(report, "report")
    require(type(report.get("schema_version")) is int and report["schema_version"] == 1,
            "report.schema_version: expected 1")
    tool = object_at(report.get("tool"), "tool")
    require(tool.get("name") == "litchi-perf-baseline", "tool.name: unexpected harness")
    observer = mode == "observer"
    require(tool.get("binary") == ("litchi-perf-baseline-alloc" if observer else "litchi-perf-baseline"),
            "tool.binary: does not match requested mode")
    if observer:
        require(tool.get("instrumentation") in (
            "system_allocator_operation_scoped",
            "ordinary_save_procfs_and_system_allocator_operation_scoped",
        ), "tool.instrumentation: expected allocator observer")
        require(tool.get("allocator_counter_revision") == REVISION,
                "tool.allocator_counter_revision: expected serialized_region_peak_v3")
    else:
        require(tool.get("instrumentation") == "none", "tool.instrumentation: expected none")
        require("allocator_counter_revision" not in tool,
                "tool.allocator_counter_revision: native report must omit allocator revision")
    config = object_at(report.get("configuration"), "configuration")
    require(uint(config.get("samples_per_case"), "configuration.samples_per_case", positive=True) == expected_samples,
            "configuration.samples_per_case: does not match frozen expectation")
    require(uint(config.get("warmup_iterations_per_case"), "configuration.warmup_iterations_per_case") == expected_warmup,
            "configuration.warmup_iterations_per_case: does not match frozen expectation")
    require(config.get("cases") == [expected_case], "configuration.cases: does not match frozen case")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1, "results: expected one case")
    result = object_at(results[0], "results[0]")
    require(result.get("case") == expected_case, "results[0].case: does not match frozen case")
    elapsed = object_at(result.get("elapsed_ns"), "elapsed_ns")
    require(elapsed.get("unit") == "ns", "elapsed_ns.unit: expected ns")
    values = elapsed.get("samples")
    require(isinstance(values, list) and len(values) == expected_samples,
            "elapsed_ns.samples: wrong sample count")
    for i, value in enumerate(values):
        uint(value, f"elapsed_ns.samples[{i}]")
    require(values == sorted(values), "elapsed_ns.samples: expected elapsed-sorted samples")
    order = elapsed.get("sample_order")
    require(isinstance(order, list) and len(order) == expected_samples,
            "elapsed_ns.sample_order: wrong sample count")
    for i, index in enumerate(order):
        uint(index, f"elapsed_ns.sample_order[{i}]")
    require(sorted(order) == list(range(expected_samples)),
            "elapsed_ns.sample_order: expected a complete permutation")
    require(all(values[i] != values[i + 1] or order[i] < order[i + 1]
                for i in range(expected_samples - 1)), "elapsed_ns.sample_order: tied samples out of order")
    metrics = object_at(result.get("operation_metrics"), "operation_metrics")
    require(uint(metrics.get("sample_count"), "operation_metrics.sample_count", positive=True) == expected_samples,
            "operation_metrics.sample_count: wrong sample count")
    indices = metrics.get("sample_indices")
    require(isinstance(indices, list) and all(type(i) is int for i in indices) and indices == order,
            "operation_metrics.sample_indices: must equal elapsed_ns.sample_order")
    require(metrics.get("alignment") == ALIGNMENT, "operation_metrics.alignment: unexpected alignment")
    require(metrics.get("latency_claim") == (
        "allocator_instrumented_elapsed_not_latency_claim" if observer else "comparable_timed_operation"
    ), "operation_metrics.latency_claim: does not match requested mode")
    validate_allocation(metrics.get("allocation"), samples=expected_samples, measured=observer)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report", type=Path)
    parser.add_argument("--case", required=True)
    parser.add_argument("--samples", required=True, type=int)
    parser.add_argument("--warmup", required=True, type=int)
    parser.add_argument("--mode", required=True, choices=("native", "observer"))
    args = parser.parse_args(argv)
    try:
        validate_report(json.loads(args.report.read_text()), expected_case=args.case,
                        expected_samples=args.samples, expected_warmup=args.warmup, mode=args.mode)
    except (AllocationSchemaError, OSError, ValueError) as error:
        print(f"allocation schema rejected: {error}", file=sys.stderr)
        return 1
    print("allocation schema PASS; source, artifact semantics, and timing comparability require separate admission")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
