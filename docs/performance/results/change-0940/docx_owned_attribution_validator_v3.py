#!/usr/bin/env python3
"""Independent semantic validator for the 0940 owned-ingress report.

This module is loaded by ``tools/perf_child_validation.py`` after the child
terminal receipt has been published.  It intentionally validates the report
shape and conservation laws independently of the Rust implementation.
"""

from __future__ import annotations

import math
from typing import Any, Mapping, Sequence


SCHEMA = "docx_owned_ingress_attribution_v1"
SCENARIO = "docx-owned-ingress-attribution"
CORPUS_GENERATOR = "litchi-docx-section-layout-source-v1"
SOURCE_BYTES = 2_101_427
SOURCE_SHA256 = "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8"
EXPECTED_OUTPUT_BYTES = 2_101_446
EXPECTED_OUTPUT_SHA256 = "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
ROUTES = frozenset({"owned_vec", "owned_reader"})
CALIBRATION_ROUTE = "route_independent"
DIGEST_LENGTH = 64
ALLOCATION_STATUSES = frozenset({"measured", "unavailable", "overflow"})
ALLOCATION_NUMERIC = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
PROCESS_SNAPSHOT_FIELDS = (
    "rchar",
    "wchar",
    "read_bytes",
    "write_bytes",
    "cancelled_write_bytes",
    "syscr",
    "syscw",
    "minor_faults",
    "major_faults",
    "user_cpu_ticks",
    "system_cpu_ticks",
    "clock_ticks_per_second",
    "voluntary_context_switches",
    "nonvoluntary_context_switches",
    "rss_bytes",
    "peak_rss_bytes",
)
PROCESS_DELTA_FIELDS = PROCESS_SNAPSHOT_FIELDS
PROCESS_STATUSES = frozenset({
    "disabled",
    "measured",
    "procfs_not_found",
    "procfs_permission_denied",
    "procfs_invalid_data",
    "procfs_read_error",
})
BOUND_REPORT_FIELDS = (
    "source_revision",
    "binary_identity",
    "diagnostic_arm",
    "diagnostic_enabled",
    "allocator_identity",
    "instrumentation_identity",
    "allocator_counter_revision",
    "timing_scope",
    "observer_gap_scope",
)
CORRECTNESS_BOOLEAN_FIELDS = (
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


class ValidationError(ValueError):
    """The report violates the independently checked 0940 contract."""


def _fail(message: str) -> None:
    raise ValidationError(message)


def _map(value: Any, location: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        _fail(f"{location} must be an object")
    return value


def _list(value: Any, location: str) -> list[Any]:
    if not isinstance(value, list):
        _fail(f"{location} must be a list")
    return value


def _int(value: Any, location: str, *, nonnegative: bool = True) -> int:
    if type(value) is not int or (nonnegative and value < 0):
        _fail(f"{location} must be a {'non-negative ' if nonnegative else ''}integer")
    return value


def _bool(value: Any, location: str) -> bool:
    if type(value) is not bool:
        _fail(f"{location} must be boolean")
    return value


def _digest(value: Any, location: str) -> str:
    if (
        not isinstance(value, str)
        or len(value) != DIGEST_LENGTH
        or any(char not in "0123456789abcdef" for char in value)
    ):
        _fail(f"{location} must be lowercase SHA-256 hex")
    return value


def _exact(value: Any, expected: Any, location: str) -> None:
    if value != expected or type(value) is not type(expected):
        _fail(f"{location} does not equal expected value")


def _require_finite(value: Any, location: str) -> float:
    if type(value) not in (int, float) or not math.isfinite(float(value)):
        _fail(f"{location} must be finite")
    return float(value)


def _nearest_rank(values: Sequence[int], percentile: int) -> int:
    index = ((percentile * len(values) + 99) // 100) - 1
    return values[min(index, len(values) - 1)]


def _midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def _permutation(order: Any, count: int, location: str) -> list[int]:
    values = _list(order, location)
    if len(values) != count or sorted(values) != list(range(count)):
        _fail(f"{location} must be a permutation of sample indices")
    return [_int(value, f"{location}[{index}]") for index, value in enumerate(values)]


def _stats(value: Any, count: int, location: str) -> Mapping[str, Any]:
    stats = _map(value, location)
    _exact(stats.get("unit"), "ns", f"{location}.unit")
    samples = [_int(item, f"{location}.samples[{index}]") for index, item in enumerate(_list(stats.get("samples"), f"{location}.samples"))]
    if len(samples) != count or samples != sorted(samples):
        _fail(f"{location}.samples must be sorted and have the sample count")
    order = _permutation(stats.get("sample_order"), count, f"{location}.sample_order")
    _exact(stats.get("min"), samples[0], f"{location}.min")
    _exact(stats.get("p50"), _midpoint(samples[(count - 1) // 2], samples[count // 2]), f"{location}.p50")
    _exact(stats.get("p95"), _nearest_rank(samples, 95), f"{location}.p95")
    _exact(stats.get("p99"), _nearest_rank(samples, 99), f"{location}.p99")
    _exact(stats.get("max"), samples[-1], f"{location}.max")
    _require_finite(stats.get("mean"), f"{location}.mean")
    _require_finite(stats.get("standard_deviation"), f"{location}.standard_deviation")
    if float(stats["standard_deviation"]) < 0:
        _fail(f"{location}.standard_deviation must be nonnegative")
    interval = _map(stats.get("confidence_interval_95"), f"{location}.confidence_interval_95")
    _require_finite(interval.get("lower"), f"{location}.confidence_interval_95.lower")
    _require_finite(interval.get("upper"), f"{location}.confidence_interval_95.upper")
    if float(interval["lower"]) > float(interval["upper"]):
        _fail(f"{location}.confidence_interval_95 is inverted")
    return {"samples": samples, "sample_order": order}


def _inverse_ranked(values: Sequence[Any], order: Sequence[int], location: str) -> list[Any]:
    if len(values) != len(order):
        _fail(f"{location} does not align with sample_order")
    chronological: list[Any] = [None] * len(order)
    for rank, acquisition_index in enumerate(order):
        if chronological[acquisition_index] is not None:
            _fail(f"{location} has duplicate acquisition index")
        chronological[acquisition_index] = values[rank]
    return chronological


def _u64_vector(report: Mapping[str, Any], key: str, count: int) -> list[int]:
    return [
        _int(value, f"{key}[{index}]")
        for index, value in enumerate(_list(report.get(key), key))
    ] if len(_list(report.get(key), key)) == count else _fail(f"{key} has wrong length")


def _phase_vectors(report: Mapping[str, Any], case: str, count: int) -> dict[str, list[int]]:
    names = ["preparation_ns", "observer_gap_ns", "open_ns"]
    if case == "workload":
        names.extend(["query_ns", "edit_ns", "commit_ns", "publication_ns"])
    else:
        for name in ("query_ns", "edit_ns", "commit_ns", "publication_ns"):
            values = _list(report.get(name), name)
            if values:
                _fail(f"calibration {name} must be empty")
    return {name: _u64_vector(report, name, count) for name in names}


def _check_timing(report: Mapping[str, Any], case: str, count: int) -> None:
    operation = _stats(report.get("operation_elapsed_ns"), count, "operation_elapsed_ns")
    wall = _stats(report.get("wall_elapsed_ns"), count, "wall_elapsed_ns")
    top_order = _permutation(report.get("sample_order"), count, "sample_order")
    if top_order != operation["sample_order"]:
        _fail("top-level sample_order differs from operation_elapsed_ns.sample_order")
    phases = _phase_vectors(report, case, count)
    ranked_phase = {key: values for key, values in phases.items()}
    operation_by_acquisition = _inverse_ranked(
        operation["samples"], operation["sample_order"], "operation_elapsed_ns.samples"
    )
    phase_by_acquisition = {
        key: _inverse_ranked(values, top_order, key) for key, values in ranked_phase.items()
    }
    for index, observed in enumerate(operation_by_acquisition):
        expected = sum(
            values[index]
            for key, values in phase_by_acquisition.items()
            if key != "observer_gap_ns"
        )
        if observed != expected:
            _fail(f"operation phase sum mismatch at acquisition index {index}")
    # wall_elapsed_ns is an outer wall-clock bracket. It includes the process
    # probes and boundary bookkeeping described by observer_gap_scope, so it
    # is intentionally not treated as an additive phase sum.


def _check_allocation_sample(value: Any, location: str) -> Mapping[str, Any]:
    sample = _map(value, location)
    status = sample.get("status")
    if status not in ALLOCATION_STATUSES:
        _fail(f"{location}.status is unknown")
    _exact(sample.get("scope"), "operation_global_system_allocator", f"{location}.scope")
    present_numeric = [key for key in ALLOCATION_NUMERIC if key in sample]
    if status != "measured":
        if present_numeric:
            _fail(f"{location} has numeric values despite status {status}")
        return sample
    if len(present_numeric) != len(ALLOCATION_NUMERIC):
        _fail(f"{location} measured sample has incomplete numeric values")
    values: dict[str, Any] = {
        "status": status,
        "scope": sample["scope"],
    }
    values.update(
        {key: _int(sample[key], f"{location}.{key}") for key in ALLOCATION_NUMERIC}
    )
    if values["peak_live_bytes_before"] < values["live_bytes_before"]:
        _fail(f"{location} peak before is below live before")
    if values["peak_live_bytes_after"] < values["peak_live_bytes_before"]:
        _fail(f"{location} peak after regressed")
    if values["peak_live_bytes_after"] < values["live_bytes_after"]:
        _fail(f"{location} peak after is below live after")
    if values["region_peak_live_bytes"] < max(
        values["live_bytes_before"], values["live_bytes_after"]
    ):
        _fail(f"{location} region peak is below live boundary")
    if values["region_peak_live_bytes"] > values["peak_live_bytes_after"]:
        _fail(f"{location} region peak exceeds process peak after")
    return values


def _check_allocation(report: Mapping[str, Any], enabled: bool, count: int) -> None:
    allocation = report.get("allocation")
    if not enabled:
        if allocation is not None:
            _fail("diagnostic-disabled report published allocation metrics")
        return
    groups = _map(allocation, "allocation")
    samples = {
        name: _list(groups.get(name), f"allocation.{name}") for name in ("prep", "observer_gap", "open", "combined_prep_gap_open")
    }
    if any(len(values) != count for values in samples.values()):
        _fail("allocation phase vectors do not have the sample count")
    for name, values in samples.items():
        for index, value in enumerate(values):
            samples[name][index] = _check_allocation_sample(value, f"allocation.{name}[{index}]")
    for index in range(count):
        prep = samples["prep"][index]
        gap = samples["observer_gap"][index]
        opened = samples["open"][index]
        combined = samples["combined_prep_gap_open"][index]
        statuses = {item["status"] for item in (prep, gap, opened, combined)}
        if len(statuses) != 1:
            _fail(f"allocation status mismatch at sample {index}")
        if combined["status"] != "measured":
            continue
        for field in (
            "allocation_calls",
            "deallocation_calls",
            "reallocation_calls",
            "failed_allocation_calls",
            "allocated_bytes",
            "deallocated_bytes",
        ):
            if prep[field] + gap[field] + opened[field] != combined[field]:
                _fail(f"allocation {field} conservation failed at sample {index}")
        if prep["live_bytes_after"] != gap["live_bytes_before"]:
            _fail(f"allocation prep/gap boundary failed at sample {index}")
        if gap["live_bytes_after"] != opened["live_bytes_before"]:
            _fail(f"allocation gap/open boundary failed at sample {index}")
        if prep["live_bytes_before"] != combined["live_bytes_before"]:
            _fail(f"allocation combined start boundary failed at sample {index}")
        if opened["live_bytes_after"] != combined["live_bytes_after"]:
            _fail(f"allocation combined end boundary failed at sample {index}")
        expected_region_peak = max(
            prep["region_peak_live_bytes"],
            gap["region_peak_live_bytes"],
            opened["region_peak_live_bytes"],
        )
        if combined["region_peak_live_bytes"] != expected_region_peak:
            _fail(f"allocation combined region peak is not the segment maximum at sample {index}")


def _check_process_observation(value: Any, location: str) -> bool:
    observation = _map(value, location)
    before = observation.get("before")
    after = observation.get("after")
    delta = observation.get("delta")
    signed_rss = observation.get("rss_delta_signed")
    before_status = observation.get("before_status")
    after_status = observation.get("after_status")
    rss_status = observation.get("rss_delta_status")
    if before_status not in PROCESS_STATUSES:
        _fail(f"{location}.before_status is not a recognized Rust process status")
    if after_status not in PROCESS_STATUSES:
        _fail(f"{location}.after_status is not a recognized Rust process status")
    if before_status == "disabled" or after_status == "disabled":
        _fail(f"{location} diagnostic-enabled process observation cannot be disabled")
    if not isinstance(rss_status, str):
        _fail(f"{location}.rss_delta_status is missing")
    if (before is None) != (before_status != "measured"):
        _fail(f"{location}.before status disagrees with snapshot availability")
    if (after is None) != (after_status != "measured"):
        _fail(f"{location}.after status disagrees with snapshot availability")
    if before is None and after is None and delta is None and signed_rss is None:
        if rss_status != "unavailable":
            _fail(f"{location}.rss_delta_status must be unavailable without snapshots")
        return False
    if not all(isinstance(item, Mapping) for item in (before, after, delta)):
        if delta is not None or signed_rss is not None or rss_status != "unavailable":
            _fail(f"{location} has an inconsistent partial process observation")
        return False
    before = _map(before, f"{location}.before")
    after = _map(after, f"{location}.after")
    delta = _map(delta, f"{location}.delta")
    before_values = {key: _int(before.get(key), f"{location}.before.{key}") for key in PROCESS_SNAPSHOT_FIELDS}
    after_values = {key: _int(after.get(key), f"{location}.after.{key}") for key in PROCESS_SNAPSHOT_FIELDS}
    for key in PROCESS_DELTA_FIELDS:
        if key in {"clock_ticks_per_second", "peak_rss_bytes"}:
            expected = after_values[key]
        else:
            expected = max(0, after_values[key] - before_values[key])
        if delta.get(key) != expected:
            _fail(f"{location}.delta.{key} does not match saturating process delta")
    expected_signed_rss = after_values["rss_bytes"] - before_values["rss_bytes"]
    operands_fit_signed_rss = (
        before_values["rss_bytes"] <= (1 << 63) - 1
        and after_values["rss_bytes"] <= (1 << 63) - 1
    )
    if operands_fit_signed_rss:
        if signed_rss != expected_signed_rss:
            _fail(f"{location}.rss_delta_signed does not preserve RSS direction")
        if rss_status != "measured":
            _fail(f"{location}.rss_delta_status must be measured for representable snapshots")
    else:
        if signed_rss is not None or rss_status != "overflow":
            _fail(f"{location}.rss_delta_status must report overflow for unrepresentable snapshots")
    return True


def _check_process(report: Mapping[str, Any], enabled: bool, count: int) -> None:
    process = report.get("process")
    if not enabled:
        if process is not None:
            _fail("diagnostic-disabled report published process metrics")
        return
    groups = _map(process, "process")
    names = ("prep", "observer_gap", "open", "combined_prep_gap_open")
    observations = {
        name: _list(groups.get(name), f"process.{name}") for name in names
    }
    if any(len(values) != count for values in observations.values()):
        _fail("process phase vectors do not have the sample count")
    availability: dict[str, list[bool]] = {}
    for name, values in observations.items():
        availability[name] = [
            _check_process_observation(value, f"process.{name}[{index}]")
            for index, value in enumerate(values)
        ]
    for index in range(count):
        prep = observations["prep"][index]
        gap = observations["observer_gap"][index]
        opened = observations["open"][index]
        combined = observations["combined_prep_gap_open"][index]
        if not all(availability[name][index] for name in names):
            continue
        prep_before = _map(prep["before"], f"process.prep[{index}].before")
        gap_before = _map(gap["before"], f"process.observer_gap[{index}].before")
        gap_after = _map(gap["after"], f"process.observer_gap[{index}].after")
        open_before = _map(opened["before"], f"process.open[{index}].before")
        combined_before = _map(combined["before"], f"process.combined[{index}].before")
        open_after = _map(opened["after"], f"process.open[{index}].after")
        combined_after = _map(combined["after"], f"process.combined[{index}].after")
        if prep_before != combined_before or open_after != combined_after:
            _fail(f"process combined boundary mismatch at sample {index}")
        if gap_before != _map(prep["after"], f"process.prep[{index}].after"):
            _fail(f"process prep/gap boundary mismatch at sample {index}")
        if open_before != gap_after:
            _fail(f"process gap/open boundary mismatch at sample {index}")


def _check_identity(report: Mapping[str, Any], route: str, warmups: int, samples: int) -> str:
    _exact(report.get("schema"), SCHEMA, "schema")
    _exact(report.get("scenario"), SCENARIO, "scenario")
    _exact(report.get("route"), route, "route")
    _exact(report.get("warmup_iterations"), warmups, "warmup_iterations")
    _exact(report.get("sample_count"), samples, "sample_count")
    source_revision = report.get("source_revision")
    if (
        not isinstance(source_revision, str)
        or len(source_revision) != 40
        or any(char not in "0123456789abcdefABCDEF" for char in source_revision)
    ):
        _fail("source_revision must be a 40-character revision")
    source_bytes = _int(report.get("source_bytes"), "source_bytes")
    _exact(source_bytes, SOURCE_BYTES, "source_bytes")
    _exact(report.get("source_sha256"), SOURCE_SHA256, "source_sha256")
    _exact(report.get("source_corpus_generator"), CORPUS_GENERATOR, "source_corpus_generator")
    diagnostic_enabled = _bool(report.get("diagnostic_enabled"), "diagnostic_enabled")
    binary_identity = report.get("binary_identity")
    if diagnostic_enabled:
        expected_arm = {
            "litchi-perf-baseline": "normal_on",
            "litchi-perf-baseline-alloc": "allocator_on",
        }.get(binary_identity)
        if expected_arm is None:
            _fail("diagnostic-enabled report has unknown binary identity")
    else:
        if binary_identity != "litchi-perf-baseline":
            _fail("diagnostic-off report must use the normal binary")
        expected_arm = "normal_off"
    _exact(report.get("diagnostic_arm"), expected_arm, "diagnostic_arm")
    if binary_identity == "litchi-perf-baseline-alloc":
        _exact(report.get("allocator_identity"), "CountingSystemAllocator(std::alloc::System)", "allocator_identity")
        _exact(report.get("instrumentation_identity"), "system_allocator_operation_scoped", "instrumentation_identity")
        _exact(report.get("allocator_counter_revision"), "serialized_region_peak_v3", "allocator_counter_revision")
    else:
        _exact(report.get("allocator_identity"), "Rust system allocator", "allocator_identity")
        _exact(report.get("instrumentation_identity"), "none", "instrumentation_identity")
        if report.get("allocator_counter_revision") is not None:
            _fail("normal report must not publish allocator counter revision")
    case = report.get("case")
    if case not in {"workload", "calibration"}:
        _fail("case must be workload or calibration")
    return case


def _check_correctness_gates(value: Mapping[str, Any], case: str, samples: int) -> None:
    gates = _list(value.get("correctness_gates"), "correctness_gates")
    if case == "calibration":
        if gates:
            _fail("calibration correctness_gates must be empty")
        return
    if len(gates) != samples:
        _fail("workload correctness_gates has wrong length")
    expected_cache: tuple[int, int, int, int, int] | None = None
    for index, item in enumerate(gates):
        gate = _map(item, f"correctness_gates[{index}]")
        for key in CORRECTNESS_BOOLEAN_FIELDS:
            if gate.get(key) is not True:
                _fail(f"correctness_gates[{index}].{key} is not true")
        for key in (
            "cache_hits",
            "cache_cold_loads",
            "cache_successful_loads",
            "cache_retained_entries",
            "cache_retained_bytes",
        ):
            _int(gate.get(key), f"correctness_gates[{index}].{key}")
        if gate["cache_successful_loads"] != 1 or gate["cache_cold_loads"] != 1:
            _fail(f"correctness_gates[{index}] cache load counts are not one")
        if gate["cache_retained_entries"] != 1 or gate["cache_retained_bytes"] == 0:
            _fail(f"correctness_gates[{index}] cache retention is absent")
        cache = tuple(gate[key] for key in (
            "cache_hits",
            "cache_cold_loads",
            "cache_successful_loads",
            "cache_retained_entries",
            "cache_retained_bytes",
        ))
        if expected_cache is None:
            expected_cache = cache
        elif cache != expected_cache:
            _fail("workload cache gate counts are not deterministic across samples")


def validate_report(
    report: dict[str, Any],
    route: str,
    *,
    warmups: int,
    samples: int,
    label: str,
    binary: Mapping[str, Any] | None,
) -> Mapping[str, Any]:
    del label
    value = _map(report, "report")
    case = _check_identity(value, route, warmups, samples)
    if not isinstance(binary, Mapping):
        _fail("DOCX owned report requires bound binary metadata")
    expected_report = binary.get("expected_report")
    if not isinstance(expected_report, Mapping):
        _fail("bound binary metadata requires expected_report mapping")
    for key in BOUND_REPORT_FIELDS:
        if key not in expected_report:
            _fail(f"bound binary metadata is missing expected_report.{key}")
        _exact(value.get(key), expected_report[key], f"report.{key} bound metadata")
    for key in ("expected_allocation_present", "expected_process_present"):
        expected_presence = binary.get(key)
        if type(expected_presence) is not bool:
            _fail(f"bound binary metadata {key} must be boolean")
        actual_key = "allocation" if key == "expected_allocation_present" else "process"
        _exact(value.get(actual_key) is not None, expected_presence, f"report.{key} bound metadata")
    if case == "calibration":
        _exact(route, CALIBRATION_ROUTE, "calibration route")
        if value.get("diagnostic_enabled") is not True:
            _fail("calibration must enable diagnostics")
    elif route not in ROUTES:
        _fail("workload route is not an owned route")
    _check_timing(value, case, samples)
    _check_correctness_gates(value, case, samples)
    _check_allocation(value, bool(value["diagnostic_enabled"]), samples)
    _check_process(value, bool(value["diagnostic_enabled"]), samples)
    if case == "workload":
        _exact(value.get("correctness_status"), "verified", "correctness_status")
        correctness = _list(value.get("correctness_verified"), "correctness_verified")
        inverse = _list(value.get("inverse_source_verified"), "inverse_source_verified")
        if correctness != [True] * samples or inverse != [True] * samples:
            _fail("workload correctness/inverse vectors are not all true")
        output_values = _list(value.get("output_sha256"), "output_sha256")
        if len(output_values) != samples:
            _fail("workload output_sha256 has wrong length")
        for index, output in enumerate(output_values):
            _digest(output, f"output_sha256[{index}]")
        expected = _digest(value.get("expected_output_sha256"), "expected_output_sha256")
        _exact(value.get("expected_output_bytes"), EXPECTED_OUTPUT_BYTES, "expected_output_bytes")
        _exact(expected, EXPECTED_OUTPUT_SHA256, "expected_output_sha256")
        if any(output != expected for output in output_values):
            _fail("workload output digest differs from expected output")
        sink = _list(value.get("sink"), "sink")
        if len(sink) != samples:
            _fail("workload sink vector has wrong length")
        sink_summaries: list[tuple[int, int, int]] = []
        for index, item in enumerate(sink):
            sink_item = _map(item, f"sink[{index}]")
            accepted = _int(sink_item.get("accepted_bytes"), f"sink[{index}].accepted_bytes")
            writes = _int(sink_item.get("write_calls"), f"sink[{index}].write_calls")
            largest = _int(sink_item.get("largest_write"), f"sink[{index}].largest_write")
            expected_bytes = _int(value.get("expected_output_bytes"), "expected_output_bytes")
            if (
                accepted != expected_bytes
                or accepted == 0
                or writes == 0
                or largest > accepted
                or largest > 64 * 1024
            ):
                _fail(f"sink[{index}] has inconsistent write accounting")
            sink_summaries.append((accepted, writes, largest))
        if len(set(sink_summaries)) != 1:
            _fail("workload sink summary is not deterministic across samples")
        for key in ("owned_input_bytes",):
            values = [_int(item, f"{key}[{index}]") for index, item in enumerate(_list(value.get(key), key))]
            if len(values) != samples or any(item != value["source_bytes"] for item in values):
                _fail(f"{key} does not bind every workload sample to source_bytes")
        for key in ("owned_ingress_read_calls", "owned_ingress_read_bytes"):
            values = _list(value.get(key), key)
            if len(values) != samples:
                _fail(f"{key} has wrong length")
            if route == "owned_vec" and any(item is not None for item in values):
                _fail(f"{key} must be unavailable for owned_vec")
            if route == "owned_reader":
                if any(type(item) is not int or item <= 0 for item in values):
                    _fail(f"{key} must contain positive reader counters")
        return {"elapsed_ns": _stats(value["operation_elapsed_ns"], samples, "operation_elapsed_ns")["samples"], "output_sha256": expected}
    _exact(value.get("correctness_status"), "not_applicable", "correctness_status")
    for key in ("correctness_verified", "inverse_source_verified", "output_sha256", "sink", "owned_input_bytes", "owned_ingress_read_calls", "owned_ingress_read_bytes"):
        if _list(value.get(key), key):
            _fail(f"calibration {key} must be empty")
    if value.get("expected_output_bytes") is not None:
        _fail("calibration must not publish expected output bytes")
    if value.get("expected_output_sha256") is not None:
        _fail("calibration must not publish an expected output digest")
    return {"elapsed_ns": _stats(value["operation_elapsed_ns"], samples, "operation_elapsed_ns")["samples"]}
