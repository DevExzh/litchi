#!/usr/bin/env python3
"""Validate the 0940 DOCX owned-ingress diagnostic reports.

This module is deliberately offline.  It validates the report contract emitted
by ``docx-owned-attribution`` and returns compact scalar/vector evidence for the
capture driver.  It does not start a benchmark, choose a binary, or interpret
the diagnostic as an optimization result.
"""

from __future__ import annotations

import hashlib
import math
import statistics
from typing import Any, Mapping


SCHEMA = "docx_owned_ingress_attribution_v1"
SCENARIO = "docx-owned-ingress-attribution"
SOURCE_GENERATOR = "litchi-docx-section-layout-source-v1"
SOURCE_BYTES = 2_101_427
SOURCE_SHA256 = "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8"
OUTPUT_BYTES = 2_101_446
OUTPUT_SHA256 = "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
ROUTES = ("owned_vec", "owned_reader")
CALIBRATION_ROUTE = "route_independent"
WORKLOAD_PHASES = (
    "preparation_ns",
    "open_ns",
    "query_ns",
    "edit_ns",
    "commit_ns",
    "publication_ns",
)
CALIBRATION_PHASES = ("preparation_ns", "observer_gap_ns", "open_ns")
GATE_FIELDS = (
    "snapshot_shape_verified",
    "changed_nonnoop_columns_verified",
    "patch_forward_inverse_stale_verified",
    "publication_inverse_applicable_verified",
    "output_verified",
    "reopened_output_verified",
    "inverse_source_verified",
    "source_immutability_verified",
    "cache_contract_verified",
    "cache_deterministic_verified",
    "sink_bounds_verified",
    "owned_reader_ingress_verified",
)


class EvidenceError(ValueError):
    """A report does not satisfy the frozen diagnostic contract."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64, f"{label}: digest missing")
    require(all(char in "0123456789abcdef" for char in value), f"{label}: digest is not lowercase hex")
    return value


def _revision(value: Any, label: str) -> str:
    require(isinstance(value, str) and len(value) == 40, f"{label}: revision missing")
    require(all(char in "0123456789abcdef" for char in value), f"{label}: revision is not lowercase hex")
    return value


def _integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(type(value) is int and (value > 0 if positive else value >= 0), f"{label}: invalid integer")
    return int(value)


def _vector(value: Any, length: int, label: str, *, positive: bool = False) -> list[int]:
    require(isinstance(value, list) and len(value) == length, f"{label}: length changed")
    return [_integer(item, f"{label}[{index}]", positive=positive) for index, item in enumerate(value)]


def _bool_vector(value: Any, length: int, label: str, expected: bool | None = None) -> list[bool]:
    require(isinstance(value, list) and len(value) == length, f"{label}: length changed")
    result: list[bool] = []
    for index, item in enumerate(value):
        require(type(item) is bool, f"{label}[{index}]: expected boolean")
        if expected is not None:
            require(item is expected, f"{label}[{index}]: expected {expected}")
        result.append(item)
    return result


def _midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def _nearest_rank(values: list[int], percentile: int) -> int:
    index = min((percentile * len(values) + 99) // 100 - 1, len(values) - 1)
    return values[index]


def _check_statistics(value: Any, samples: int, label: str) -> tuple[list[int], list[int]]:
    require(isinstance(value, dict), f"{label}: statistics missing")
    require(value.get("unit") == "ns", f"{label}: unit changed")
    ordered = _vector(value.get("samples"), samples, f"{label}.samples")
    require(ordered == sorted(ordered), f"{label}.samples: values are not sorted")
    sample_order = _vector(value.get("sample_order"), samples, f"{label}.sample_order")
    require(sorted(sample_order) == list(range(samples)), f"{label}.sample_order: not a permutation")
    expected = {
        "min": ordered[0],
        "p50": _midpoint(ordered[(samples - 1) // 2], ordered[samples // 2]),
        "p95": _nearest_rank(ordered, 95),
        "p99": _nearest_rank(ordered, 99),
        "max": ordered[-1],
    }
    for field, wanted in expected.items():
        require(value.get(field) == wanted, f"{label}.{field}: statistic changed")
    require(isinstance(value.get("mean"), (int, float)) and math.isfinite(float(value["mean"])),
            f"{label}.mean: invalid value")
    require(isinstance(value.get("standard_deviation"), (int, float))
            and math.isfinite(float(value["standard_deviation"]))
            and float(value["standard_deviation"]) >= 0,
            f"{label}.standard_deviation: invalid value")
    interval = value.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{label}.confidence_interval_95: missing")
    require(interval.get("method") == "two-sided Student's t interval for the mean",
            f"{label}.confidence_interval_95.method: changed")
    for field in ("lower", "upper"):
        require(isinstance(interval.get(field), (int, float)) and math.isfinite(float(interval[field])),
                f"{label}.confidence_interval_95.{field}: invalid value")
    require(float(interval["lower"]) <= float(interval["upper"]),
            f"{label}.confidence_interval_95: inverted interval")
    require(float(value["mean"]) >= ordered[0] and float(value["mean"]) <= ordered[-1],
            f"{label}.mean: outside observed sample range")
    return ordered, sample_order


def _check_snapshot(value: Any, label: str) -> None:
    if value is None:
        return
    require(isinstance(value, dict), f"{label}: snapshot is not an object")
    for field in (
        "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes", "syscr", "syscw",
        "minor_faults", "major_faults", "user_cpu_ticks", "system_cpu_ticks",
        "clock_ticks_per_second", "voluntary_context_switches", "nonvoluntary_context_switches",
        "rss_bytes", "peak_rss_bytes",
    ):
        _integer(value.get(field), f"{label}.{field}")


def _check_process_observation(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label}: observation missing")
    for field in ("before_status", "after_status"):
        status = value.get(field)
        require(status == "measured" or (isinstance(status, str) and status.startswith("procfs_"))
                or status == "unavailable", f"{label}.{field}: invalid status")
    before_status = value.get("before_status")
    after_status = value.get("after_status")
    before = value.get("before")
    after = value.get("after")
    require((before_status == "measured") == (before is not None),
            f"{label}.before status/snapshot mismatch")
    require((after_status == "measured") == (after is not None),
            f"{label}.after status/snapshot mismatch")
    _check_snapshot(before, f"{label}.before")
    _check_snapshot(after, f"{label}.after")
    delta = value.get("delta")
    if delta is not None:
        require(isinstance(delta, dict), f"{label}.delta: not an object")
        for field in (
            "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes", "syscr", "syscw",
            "minor_faults", "major_faults", "user_cpu_ticks", "system_cpu_ticks",
            "clock_ticks_per_second", "voluntary_context_switches", "nonvoluntary_context_switches",
            "rss_bytes", "peak_rss_bytes",
        ):
            _integer(delta.get(field), f"{label}.delta.{field}")
    signed = value.get("rss_delta_signed")
    require(signed is None or type(signed) is int, f"{label}.rss_delta_signed: invalid value")
    rss_status = value.get("rss_delta_status")
    require(rss_status in ("measured", "overflow", "unavailable"),
            f"{label}.rss_delta_status: invalid status")
    if before is None or after is None:
        require(value.get("delta") is None and signed is None and rss_status == "unavailable",
                f"{label}: incomplete process observation has derived values")
    else:
        before_rss = before["rss_bytes"]
        after_rss = after["rss_bytes"]
        require(isinstance(value.get("delta"), dict), f"{label}.delta: missing")
        for field in (
            "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes", "syscr", "syscw",
            "minor_faults", "major_faults", "user_cpu_ticks", "system_cpu_ticks",
            "voluntary_context_switches", "nonvoluntary_context_switches", "rss_bytes",
        ):
            require(value["delta"].get(field) == max(0, after[field] - before[field]),
                    f"{label}.delta.{field}: mismatch")
        require(value["delta"].get("clock_ticks_per_second") == after["clock_ticks_per_second"],
                f"{label}.delta.clock_ticks_per_second: mismatch")
        require(value["delta"].get("peak_rss_bytes") == after["peak_rss_bytes"],
                f"{label}.delta.peak_rss_bytes: mismatch")
        expected_signed = after_rss - before_rss
        if -(1 << 63) <= expected_signed < (1 << 63):
            require(signed == expected_signed and rss_status == "measured",
                    f"{label}: signed RSS status mismatch")
        else:
            require(signed is None and rss_status == "overflow",
                    f"{label}: signed RSS overflow status mismatch")


def _check_allocation_sample(value: Any, label: str, *, measured: bool) -> None:
    require(isinstance(value, dict), f"{label}: sample missing")
    require(value.get("scope") == "operation_global_system_allocator", f"{label}.scope: changed")
    status = value.get("status")
    require(status == ("measured" if measured else "unavailable"), f"{label}.status: changed")
    numeric = (
        "allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls",
        "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after",
        "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes",
    )
    if measured:
        for field in numeric:
            _integer(value.get(field), f"{label}.{field}")
        require(value["peak_live_bytes_before"] >= value["live_bytes_before"],
                f"{label}: beginning peak does not cover beginning live bytes")
        require(value["peak_live_bytes_after"] >= value["live_bytes_after"],
                f"{label}: ending peak does not cover ending live bytes")
        require(value["region_peak_live_bytes"] >= max(value["live_bytes_before"], value["live_bytes_after"]),
                f"{label}: region peak does not cover live bytes")
        require(value["region_peak_live_bytes"] <= value["peak_live_bytes_after"],
                f"{label}: region peak exceeds ending process peak")
    else:
        require(all(field not in value for field in numeric), f"{label}: unavailable sample has counters")


def _check_diagnostics(report: Mapping[str, Any], samples: int, binary_identity: str) -> None:
    enabled = report.get("diagnostic_enabled")
    require(type(enabled) is bool, "diagnostic_enabled: invalid value")
    allocation = report.get("allocation")
    process = report.get("process")
    if not enabled:
        require(all(value is None for value in (allocation, process)),
                "diagnostic-off report published instrumentation")
        return
    require(isinstance(allocation, dict), "diagnostic-on allocation report missing")
    require(isinstance(process, dict), "diagnostic-on process report missing")
    measured = binary_identity == "litchi-perf-baseline-alloc"
    allocation_groups: dict[str, list[dict[str, Any]]] = {}
    process_groups: dict[str, list[dict[str, Any]]] = {}
    for phase in ("prep", "observer_gap", "open", "combined_prep_gap_open"):
        samples_value = allocation.get(phase)
        require(isinstance(samples_value, list) and len(samples_value) == samples,
                f"allocation.{phase}: length changed")
        for index, sample in enumerate(samples_value):
            _check_allocation_sample(sample, f"allocation.{phase}[{index}]", measured=measured)
        allocation_groups[phase] = samples_value
        process_value = process.get(phase)
        require(isinstance(process_value, list) and len(process_value) == samples,
                f"process.{phase}: length changed")
        for index, observation in enumerate(process_value):
            _check_process_observation(observation, f"process.{phase}[{index}]")
        process_groups[phase] = process_value
    for index in range(samples):
        alloc_values = [allocation_groups[name][index] for name in ("prep", "observer_gap", "open")]
        combined = allocation_groups["combined_prep_gap_open"][index]
        statuses = {value["status"] for value in alloc_values + [combined]}
        require(len(statuses) == 1, f"allocation status mismatch at sample {index}")
        if measured:
            for field in (
                "allocation_calls", "deallocation_calls", "reallocation_calls",
                "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
            ):
                require(sum(value[field] for value in alloc_values) == combined[field],
                        f"allocation {field} conservation failed at sample {index}")
            require(alloc_values[0]["live_bytes_after"] == alloc_values[1]["live_bytes_before"],
                    f"allocation prep/gap boundary failed at sample {index}")
            require(alloc_values[1]["live_bytes_after"] == alloc_values[2]["live_bytes_before"],
                    f"allocation gap/open boundary failed at sample {index}")
            require(alloc_values[0]["live_bytes_before"] == combined["live_bytes_before"],
                    f"allocation combined start boundary failed at sample {index}")
            require(alloc_values[2]["live_bytes_after"] == combined["live_bytes_after"],
                    f"allocation combined end boundary failed at sample {index}")
            require(combined["region_peak_live_bytes"] == max(value["region_peak_live_bytes"] for value in alloc_values),
                    f"allocation combined peak failed at sample {index}")


def _check_common(report: Mapping[str, Any], route: str, warmups: int, samples: int,
                  binary: Mapping[str, Any] | None) -> tuple[str, list[int]]:
    require(report.get("schema") == SCHEMA, "report schema changed")
    require(report.get("scenario") == SCENARIO, "report scenario changed")
    require(report.get("route") == route, "report route changed")
    require(report.get("source_corpus_generator") == SOURCE_GENERATOR, "source generator changed")
    require(report.get("source_bytes") == SOURCE_BYTES, "source bytes changed")
    require(report.get("source_sha256") == SOURCE_SHA256, "source digest changed")
    _revision(report.get("source_revision"), "source_revision")
    require(report.get("warmup_iterations") == warmups, "warmup count changed")
    require(report.get("sample_count") == samples, "sample count changed")
    identity = report.get("binary_identity")
    require(identity in ("litchi-perf-baseline", "litchi-perf-baseline-alloc"),
            "binary identity changed")
    if binary is not None:
        expected = binary.get("binary_identity", binary.get("identity"))
        if expected is not None:
            require(identity == expected, "report binary does not match bound binary")
        expected_revision = binary.get("source_revision")
        if expected_revision is not None:
            require(report.get("source_revision") == expected_revision, "source revision changed")
    enabled = report.get("diagnostic_enabled")
    require(type(enabled) is bool, "diagnostic_enabled: invalid value")
    expected_arm = {
        ("litchi-perf-baseline", False): "normal_off",
        ("litchi-perf-baseline", True): "normal_on",
        ("litchi-perf-baseline-alloc", True): "allocator_on",
    }.get((identity, enabled))
    require(expected_arm is not None and report.get("diagnostic_arm") == expected_arm,
            "diagnostic arm does not match binary and mode")
    if binary is not None:
        expected_report = binary.get("expected_report")
        require(isinstance(expected_report, Mapping), "bound expected report metadata missing")
        expected_fields = (
            "source_revision", "binary_identity", "diagnostic_arm", "diagnostic_enabled",
            "allocator_identity", "instrumentation_identity", "allocator_counter_revision",
            "timing_scope", "observer_gap_scope",
        )
        for field in expected_fields:
            require(field in expected_report, f"bound expected report field missing: {field}")
            require(report.get(field) == expected_report[field],
                    f"report.{field} does not match bound expected metadata")
        for field in ("expected_allocation_present", "expected_process_present"):
            require(type(binary.get(field)) is bool, f"bound {field} missing")
            require((report.get("allocation") is not None) == binary[field],
                    f"report allocation presence does not match {field}")
            process_present = report.get("process") is not None
            require(process_present == binary["expected_process_present"],
                    f"report process presence does not match expected_process_present")
    if identity == "litchi-perf-baseline-alloc":
        require(report.get("allocator_identity") == "CountingSystemAllocator(std::alloc::System)",
                "allocator identity changed")
        require(report.get("instrumentation_identity") == "system_allocator_operation_scoped",
                "allocator instrumentation identity changed")
        require(report.get("allocator_counter_revision") == "serialized_region_peak_v3",
                "allocator counter revision changed")
    else:
        require(report.get("allocator_identity") == "Rust system allocator",
                "normal allocator identity changed")
        require(report.get("instrumentation_identity") == "none",
                "normal instrumentation identity changed")
        require(report.get("allocator_counter_revision") is None,
                "normal report published allocator counter revision")
    operation, operation_order = _check_statistics(
        report.get("operation_elapsed_ns"), samples, "operation_elapsed_ns"
    )
    wall, wall_order = _check_statistics(report.get("wall_elapsed_ns"), samples, "wall_elapsed_ns")
    require(len(wall) == len(operation), "timing sample counts changed")
    require(report.get("sample_order") == operation_order, "sample order is not bound to operation statistics")
    require(sorted(report["sample_order"]) == list(range(samples)), "sample order is not a permutation")
    _check_diagnostics(report, samples, identity)
    return identity, operation


def _check_phase_vectors(report: Mapping[str, Any], operation: list[int], phases: tuple[str, ...]) -> None:
    values = {name: _vector(report.get(name), len(operation), name) for name in phases}
    for index, expected in enumerate(operation):
        observed = sum(values[name][index] for name in phases)
        require(observed == expected, f"phase sum changed at sample {index}: {observed} != {expected}")


def _check_gates(report: Mapping[str, Any], samples: int) -> None:
    values = report.get("correctness_gates")
    require(isinstance(values, list) and len(values) == samples, "correctness_gates length changed")
    for index, gate in enumerate(values):
        require(isinstance(gate, dict), f"correctness_gates[{index}] is not an object")
        for field in GATE_FIELDS:
            require(gate.get(field) is True, f"correctness_gates[{index}].{field} failed")
        for field in ("cache_hits", "cache_cold_loads", "cache_successful_loads", "cache_retained_entries", "cache_retained_bytes"):
            _integer(gate.get(field), f"correctness_gates[{index}].{field}")
        require((gate["cache_hits"], gate["cache_cold_loads"], gate["cache_successful_loads"],
                 gate["cache_retained_entries"], gate["cache_retained_bytes"]) == (0, 1, 1, 1, 18_855),
                f"correctness_gates[{index}]: cache contract changed")


def _check_workload(report: Mapping[str, Any], route: str, samples: int, operation: list[int]) -> None:
    require(report.get("case") == "workload", "workload case changed")
    require(route in ROUTES, "unknown workload route")
    require(report.get("timing_scope"), "timing scope missing")
    require(report.get("observer_gap_scope"), "observer gap scope missing")
    require(report.get("expected_output_bytes") == OUTPUT_BYTES, "expected output bytes changed")
    require(report.get("expected_output_sha256") == OUTPUT_SHA256, "expected output digest changed")
    outputs = report.get("output_sha256")
    require(isinstance(outputs, list) and len(outputs) == samples, "output digest vector changed")
    require(outputs == [OUTPUT_SHA256] * samples, "output digest changed")
    require(report.get("correctness_status") == "verified", "correctness status changed")
    _bool_vector(report.get("correctness_verified"), samples, "correctness_verified", True)
    _bool_vector(report.get("inverse_source_verified"), samples, "inverse_source_verified", True)
    _check_gates(report, samples)
    sink = report.get("sink")
    require(isinstance(sink, list) and len(sink) == samples, "sink vector changed")
    for index, value in enumerate(sink):
        require(isinstance(value, dict), f"sink[{index}] is not an object")
        require(value.get("accepted_bytes") == OUTPUT_BYTES, f"sink[{index}] byte count changed")
        _integer(value.get("write_calls"), f"sink[{index}].write_calls", positive=True)
        largest = _integer(value.get("largest_write"), f"sink[{index}].largest_write", positive=True)
        require(largest <= 65_536, f"sink[{index}] exceeded bounded write size")
    _vector(report.get("owned_input_bytes"), samples, "owned_input_bytes")
    require(report["owned_input_bytes"] == [SOURCE_BYTES] * samples, "owned input bytes changed")
    calls = report.get("owned_ingress_read_calls")
    bytes_value = report.get("owned_ingress_read_bytes")
    require(isinstance(calls, list) and len(calls) == samples, "owned ingress call vector changed")
    require(isinstance(bytes_value, list) and len(bytes_value) == samples, "owned ingress byte vector changed")
    if route == "owned_vec":
        require(calls == [None] * samples and bytes_value == [None] * samples,
                "owned_vec unexpectedly published reader counters")
    else:
        require(calls == [258] * samples and bytes_value == [SOURCE_BYTES] * samples,
                "owned_reader ingress counters changed")
    _check_phase_vectors(report, operation, WORKLOAD_PHASES)


def _check_calibration(report: Mapping[str, Any], samples: int, operation: list[int]) -> None:
    require(report.get("case") == "calibration", "calibration case changed")
    require(report.get("route") == CALIBRATION_ROUTE, "calibration route changed")
    require(report.get("diagnostic_enabled") is True, "calibration diagnostic mode disabled")
    require(report.get("expected_output_bytes") is None, "calibration published output bytes")
    require(report.get("expected_output_sha256") is None, "calibration published output digest")
    for field in ("output_sha256", "correctness_verified", "inverse_source_verified", "correctness_gates", "sink", "owned_ingress_read_calls", "owned_ingress_read_bytes", "owned_input_bytes"):
        value = report.get(field)
        require(isinstance(value, list) and not value, f"calibration.{field} must be empty")
    require(report.get("correctness_status") == "not_applicable", "calibration correctness status changed")
    _check_phase_vectors(report, operation, ("preparation_ns", "open_ns"))
    _vector(report.get("observer_gap_ns"), samples, "observer_gap_ns")
    for phase in ("query_ns", "edit_ns", "commit_ns", "publication_ns"):
        require(report.get(phase) == [], f"calibration.{phase} must be empty")


def validate_report(
    report: dict[str, Any],
    route: str,
    *,
    warmups: int = 5,
    samples: int = 50,
    label: str = "report",
    binary: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate one workload or calibration report and return compact vectors."""

    try:
        identity, operation = _check_common(report, route, warmups, samples, binary)
        if report.get("case") == "workload":
            _check_workload(report, route, samples, operation)
        elif report.get("case") == "calibration":
            _check_calibration(report, samples, operation)
        else:
            raise EvidenceError("unknown report case")
    except EvidenceError as error:
        raise EvidenceError(f"{label}: {error}") from error
    result: dict[str, Any] = {
        "schema": "litchi.performance.0940.scalar-validation.v1",
        "status": "passed",
        "label": label,
        "case": report["case"],
        "route": route,
        "binary_identity": identity,
        "diagnostic_enabled": report["diagnostic_enabled"],
        "sample_count": samples,
        "elapsed_ns": operation,
        "wall_elapsed_ns": report["wall_elapsed_ns"]["samples"],
        "phase_ns": {phase: report[phase] for phase in WORKLOAD_PHASES if report.get(phase)},
    }
    if report["case"] == "workload":
        result["output_sha256"] = OUTPUT_SHA256
        result["sink_accepted_bytes"] = [item["accepted_bytes"] for item in report["sink"]]
    if report["diagnostic_enabled"]:
        result["allocation"] = report["allocation"]
        result["process"] = report["process"]
    return result


def _statistics(values: list[int]) -> dict[str, Any]:
    ordered = sorted(values)
    return {
        "unit": "ns",
        "samples": ordered,
        "sample_order": sorted(range(len(values)), key=lambda index: (values[index], index)),
        "min": ordered[0],
        "p50": _midpoint(ordered[(len(values) - 1) // 2], ordered[len(values) // 2]),
        "p95": _nearest_rank(ordered, 95),
        "p99": _nearest_rank(ordered, 99),
        "max": ordered[-1],
        "mean": statistics.fmean(ordered),
        "standard_deviation": statistics.stdev(ordered) if len(ordered) > 1 else 0.0,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": 0.0,
            "upper": statistics.fmean(ordered),
        },
    }


def synthetic_report(*, route: str, case: str = "workload", diagnostic: bool = False,
                     warmups: int = 5, samples: int = 50,
                     binary_identity: str = "litchi-perf-baseline") -> dict[str, Any]:
    """Make a deterministic no-workload fixture for driver preflight tests."""

    require(case in ("workload", "calibration"), "unknown synthetic case")
    if case == "calibration":
        route = CALIBRATION_ROUTE
        diagnostic = True
    workload = case == "workload"
    prep = [10 + index for index in range(samples)]
    gap = [1 + index for index in range(samples)]
    opening = [20 + index for index in range(samples)]
    query = [30 + index for index in range(samples)] if case == "workload" else []
    edit = [40 + index for index in range(samples)] if case == "workload" else []
    commit = [50 + index for index in range(samples)] if case == "workload" else []
    publication = [60 + index for index in range(samples)] if case == "workload" else []
    if workload:
        operation_values = [
            sum(values) for values in zip(prep, opening, query, edit, commit, publication)
        ]
    else:
        operation_values = [sum(values) for values in zip(prep, opening)]
    wall_values = [value + 100 for value in operation_values]
    unavailable = {
        "status": "unavailable",
        "scope": "operation_global_system_allocator",
    }
    process_observation = {
        "before": None,
        "after": None,
        "before_status": "procfs_not_found",
        "after_status": "procfs_not_found",
        "delta": None,
        "rss_delta_signed": None,
        "rss_delta_status": "unavailable",
    }
    allocation = None
    process = None
    if diagnostic:
        if binary_identity == "litchi-perf-baseline-alloc":
            def measured(before: int, after: int, peak: int, allocated: int, calls: int = 1) -> dict[str, Any]:
                return {
                    "status": "measured",
                    "scope": "operation_global_system_allocator",
                    "allocation_calls": calls,
                    "deallocation_calls": 0,
                    "reallocation_calls": 0,
                    "failed_allocation_calls": 0,
                    "allocated_bytes": allocated,
                    "deallocated_bytes": 0,
                    "live_bytes_before": before,
                    "live_bytes_after": after,
                    "peak_live_bytes_before": before,
                    "peak_live_bytes_after": peak,
                    "region_peak_live_bytes": peak,
                }
            allocation = {
                "prep": [measured(0, 1, 1, 1) for _ in range(samples)],
                "observer_gap": [measured(1, 2, 2, 1) for _ in range(samples)],
                "open": [measured(2, 3, 3, 1) for _ in range(samples)],
                "combined_prep_gap_open": [measured(0, 3, 3, 3, calls=3) for _ in range(samples)],
            }
        else:
            allocation = {phase: [dict(unavailable) for _ in range(samples)] for phase in ("prep", "observer_gap", "open", "combined_prep_gap_open")}
        process = {phase: [dict(process_observation) for _ in range(samples)] for phase in ("prep", "observer_gap", "open", "combined_prep_gap_open")}
    report: dict[str, Any] = {
        "schema": SCHEMA,
        "scenario": SCENARIO,
        "case": case,
        "route": route,
        "diagnostic_enabled": diagnostic,
        "diagnostic_arm": (
            "allocator_on" if binary_identity == "litchi-perf-baseline-alloc" and diagnostic
            else "normal_on" if diagnostic else "normal_off"
        ),
        "binary_identity": binary_identity,
        "allocator_identity": ("CountingSystemAllocator(std::alloc::System)"
                                if binary_identity == "litchi-perf-baseline-alloc"
                                else "Rust system allocator"),
        "instrumentation_identity": ("system_allocator_operation_scoped"
                                      if binary_identity == "litchi-perf-baseline-alloc" else "none"),
        "allocator_counter_revision": ("serialized_region_peak_v3"
                                        if binary_identity == "litchi-perf-baseline-alloc" else None),
        "source_revision": "3ae856d76424f2f9b58d26dcd20fd074b8e97541",
        "source_corpus_generator": SOURCE_GENERATOR,
        "source_bytes": SOURCE_BYTES,
        "source_sha256": SOURCE_SHA256,
        "expected_output_bytes": OUTPUT_BYTES if workload else None,
        "expected_output_sha256": OUTPUT_SHA256 if workload else None,
        "warmup_iterations": warmups,
        "sample_count": samples,
        "timing_scope": "operation_elapsed_ns includes prep, open, and workload phases; wall_elapsed_ns is the measured outer bracket from before-prep observation through the selected body end (publication for workload, post-open observation for calibration)",
        "observer_gap_scope": "observer_gap_ns starts after prep's first Region::split and ends after the second split; allocation observer_gap contains both procfs reads, allocation combined ends before after-open observation, process combined includes after-open observation, and wall includes the full outer bracket",
        "operation_elapsed_ns": _statistics(operation_values),
        "wall_elapsed_ns": _statistics(wall_values),
        "sample_order": list(range(samples)),
        "preparation_ns": prep,
        "observer_gap_ns": gap,
        "open_ns": opening,
        "query_ns": query,
        "edit_ns": edit,
        "commit_ns": commit,
        "publication_ns": publication,
        "correctness_status": "verified" if workload else "not_applicable",
        "correctness_verified": [True] * samples if workload else [],
        "inverse_source_verified": [True] * samples if workload else [],
        "output_sha256": [OUTPUT_SHA256] * samples if workload else [],
        "correctness_gates": [
            {field: True for field in GATE_FIELDS}
            | {"cache_hits": 0, "cache_cold_loads": 1, "cache_successful_loads": 1,
               "cache_retained_entries": 1, "cache_retained_bytes": 18_855}
            for _ in range(samples)
        ] if workload else [],
        "sink": [{"accepted_bytes": OUTPUT_BYTES, "write_calls": 33, "largest_write": 65_536} for _ in range(samples)] if workload else [],
        "owned_ingress_read_calls": ([None] * samples if route == "owned_vec" and workload else []),
        "owned_ingress_read_bytes": ([None] * samples if route == "owned_vec" and workload else []),
        "owned_input_bytes": ([SOURCE_BYTES] * samples if workload else []),
        "allocation": allocation,
        "process": process,
    }
    if workload and route == "owned_reader":
        report["owned_ingress_read_calls"] = [258] * samples
        report["owned_ingress_read_bytes"] = [SOURCE_BYTES] * samples
    return report
