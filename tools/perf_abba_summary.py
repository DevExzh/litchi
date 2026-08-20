#!/usr/bin/env python3
"""Build a deterministic, fail-closed summary from four ABBA reports.

The input reports are the JSON files emitted by ``tools/perf-baseline``.  The
four legs are ordered ``A1, B1, B2, A2``: A is the control implementation and
B is the candidate implementation.  This module deliberately has no
third-party dependencies so that it can be used from a clean checkout.

The summary is descriptive evidence.  A statistic is marked accepted only
when both candidate directions are lower and both same-implementation drift
values are within the configured ceilings.  No speedup claim is inferred from
an accepted statistic.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

if __package__:
    from . import perf_compare
else:  # pragma: no cover - exercised by the direct CLI entry point
    import perf_compare


SCHEMA_VERSION = 1
HARNESS_SCHEMA_VERSION = 1
HARNESS_TOOL_NAME = "litchi-perf-baseline"
TOOL_NAME = "litchi-perf-abba-summary"
TOOL_VERSION = "0.1.0"
LEG_ORDER = ("a1", "b1", "b2", "a2")
STATISTICS = ("p50", "mean", "p95", "p99")
U64_MAX = (1 << 64) - 1
MIN_RETAINED_SAMPLES = 15
ENVIRONMENT_VARIANTS = frozenset(("git_revision",))
REQUIRED_TOOL_FIELDS = ("name", "version", "profile", "target_os", "target_arch")
REQUIRED_CONFIGURATION_FIELDS = (
    "samples_per_case",
    "warmup_iterations_per_case",
    "cases",
    "corpus_shapes",
)
REQUIRED_ENVIRONMENT_FIELDS = (
    "rustc_version",
    "git_revision",
    "git_worktree_dirty",
    "logical_cpus_available",
    "allocator",
    "rustflags",
    "cargo_build_target",
    "perf_event_paranoid",
)
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
STUDENT_T_CRITICAL_95 = (
    12.706,
    4.303,
    3.182,
    2.776,
    2.571,
    2.447,
    2.365,
    2.306,
    2.262,
    2.228,
    2.201,
    2.179,
    2.160,
    2.145,
    2.131,
    2.120,
    2.110,
    2.101,
    2.093,
    2.086,
    2.080,
    2.074,
    2.069,
    2.064,
    2.060,
    2.056,
    2.052,
    2.048,
    2.045,
    2.042,
)
STUDENT_T_Z_975 = 1.959_963_984_540_054
DEFAULT_DRIFT_CEILINGS: dict[str, float] = {
    "p50": 5.0,
    "mean": 5.0,
    "p95": 10.0,
    "p99": 15.0,
}


class AbbaSummaryInputError(ValueError):
    """Raised when ABBA reports are not safely comparable."""


def _validate_json_tree(value: Any, location: str) -> None:
    """Reject values that cannot be represented by strict JSON."""

    if value is None or isinstance(value, (bool, str, int)):
        return
    if isinstance(value, float):
        if not math.isfinite(value):
            raise AbbaSummaryInputError(f"{location} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _validate_json_tree(item, f"{location}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(key, str):
                raise AbbaSummaryInputError(f"{location} has a non-string object key")
            _validate_json_tree(item, f"{location}.{key}")
        return
    raise AbbaSummaryInputError(
        f"{location} contains unsupported JSON value {type(value).__name__}"
    )


def _canonical_json(value: Any, location: str) -> str:
    """Return compact canonical JSON, rejecting values JSON cannot preserve."""

    _validate_json_tree(value, location)
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError, OverflowError) as error:
        raise AbbaSummaryInputError(f"{location} is not canonical JSON: {error}") from error


def _require_object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise AbbaSummaryInputError(f"{location} must be an object")
    return value


def _finite_number(value: Any, location: str, *, positive: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise AbbaSummaryInputError(f"{location} must be a number")
    try:
        number = float(value)
    except (OverflowError, ValueError) as error:
        raise AbbaSummaryInputError(f"{location} is outside the finite range") from error
    if not math.isfinite(number):
        raise AbbaSummaryInputError(f"{location} must be finite")
    if positive and number <= 0:
        raise AbbaSummaryInputError(f"{location} must be positive")
    return number


def _u64(value: Any, location: str, *, positive: bool = False) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise AbbaSummaryInputError(f"{location} must be an unsigned 64-bit integer")
    if value < (1 if positive else 0) or value > U64_MAX:
        requirement = "positive " if positive else ""
        raise AbbaSummaryInputError(
            f"{location} must be a {requirement}unsigned 64-bit integer"
        )
    return value


def _student_t_critical_95(degrees_of_freedom: int) -> float:
    if degrees_of_freedom == 0:
        return 0.0
    if degrees_of_freedom <= len(STUDENT_T_CRITICAL_95):
        return STUDENT_T_CRITICAL_95[degrees_of_freedom - 1]
    degrees = float(degrees_of_freedom)
    z = STUDENT_T_Z_975
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    return (
        z
        + (z3 + z) / (4.0 * degrees)
        + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * degrees * degrees)
        + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
        / (384.0 * degrees * degrees * degrees)
    )


def _float_close(reported: float, expected: float, location: str) -> None:
    tolerance = max(1e-12, abs(expected) * 1e-12)
    if not math.isfinite(reported) or abs(reported - expected) > tolerance:
        raise AbbaSummaryInputError(
            f"{location}={reported} disagrees with samples ({expected})"
        )


def recompute_statistics(elapsed: Any, location: str) -> dict[str, Any]:
    """Validate an ``elapsed_ns`` object using the Rust harness formula."""

    elapsed_object = _require_object(elapsed, location)
    if elapsed_object.get("unit") != "ns":
        raise AbbaSummaryInputError(f"{location}.unit must be 'ns'")
    samples_raw = elapsed_object.get("samples")
    if not isinstance(samples_raw, list) or len(samples_raw) < MIN_RETAINED_SAMPLES:
        raise AbbaSummaryInputError(
            f"{location}.samples must retain at least {MIN_RETAINED_SAMPLES} samples"
        )
    samples = [
        _u64(value, f"{location}.samples[{index}]")
        for index, value in enumerate(samples_raw)
    ]
    if samples != sorted(samples):
        raise AbbaSummaryInputError(f"{location}.samples must be sorted ascending")

    mean = 0.0
    squared_deviation_sum = 0.0
    for index, value in enumerate(samples):
        value_as_float = float(value)
        next_count = float(index + 1)
        delta = value_as_float - mean
        next_mean = mean + delta / next_count
        squared_deviation_sum += delta * (value_as_float - next_mean)
        mean = next_mean
    standard_deviation = (
        math.sqrt(squared_deviation_sum / float(len(samples) - 1))
        if len(samples) > 1
        else 0.0
    )
    margin = (
        _student_t_critical_95(len(samples) - 1)
        * standard_deviation
        / math.sqrt(float(len(samples)))
        if len(samples) > 1
        else 0.0
    )
    left = samples[(len(samples) - 1) // 2]
    right = samples[len(samples) // 2]
    p50 = left // 2 + right // 2 + (left % 2 + right % 2) // 2

    def nearest_rank(percentile: int) -> int:
        index = ((percentile * len(samples) + 99) // 100) - 1
        return samples[min(index, len(samples) - 1)]

    computed_u64 = {
        "min": samples[0],
        "p50": p50,
        "p95": nearest_rank(95),
        "p99": nearest_rank(99),
        "max": samples[-1],
    }
    computed_float = {
        "mean": mean,
        "standard_deviation": standard_deviation,
        "confidence_interval_95.lower": max(mean - margin, 0.0),
        "confidence_interval_95.upper": mean + margin,
    }
    for name, expected in computed_u64.items():
        if name not in elapsed_object:
            raise AbbaSummaryInputError(f"{location}.{name} is required")
        reported = _u64(elapsed_object[name], f"{location}.{name}")
        if reported != expected:
            raise AbbaSummaryInputError(
                f"{location}.{name}={reported} disagrees with samples ({expected})"
            )
    for name, expected in computed_float.items():
        field = name.split(".")[-1] if "." in name else name
        if "." in name:
            confidence = _require_object(
                elapsed_object.get("confidence_interval_95"),
                f"{location}.confidence_interval_95",
            )
            value = confidence.get(field)
        else:
            if name not in elapsed_object:
                raise AbbaSummaryInputError(f"{location}.{name} is required")
            value = elapsed_object[name]
        reported = _finite_number(value, f"{location}.{name}")
        _float_close(reported, expected, f"{location}.{name}")
    confidence = _require_object(
        elapsed_object.get("confidence_interval_95"),
        f"{location}.confidence_interval_95",
    )
    if confidence.get("method") != "two-sided Student's t interval for the mean":
        raise AbbaSummaryInputError(
            f"{location}.confidence_interval_95.method does not match the harness"
        )
    return {
        "sample_count": len(samples),
        **computed_u64,
        **{
            "mean": mean,
            "standard_deviation": standard_deviation,
            "confidence_interval_95": {
                "method": confidence["method"],
                "lower": computed_float["confidence_interval_95.lower"],
                "upper": computed_float["confidence_interval_95.upper"],
            },
        },
    }


def _result_key(result: Any, location: str) -> tuple[str, str, dict[str, Any]]:
    row = _require_object(result, location)
    case = row.get("case")
    if not isinstance(case, str) or not case:
        raise AbbaSummaryInputError(f"{location}.case must be a non-empty string")
    corpus = _require_object(row.get("corpus"), f"{location}.corpus")
    if not corpus:
        raise AbbaSummaryInputError(f"{location}.corpus must not be empty")
    corpus_identity = _canonical_json(corpus, f"{location}.corpus")
    return case, corpus_identity, row


def _index_results(report: dict[str, Any], label: str) -> dict[tuple[str, str], dict[str, Any]]:
    results = report.get("results")
    if not isinstance(results, list) or not results:
        raise AbbaSummaryInputError(f"{label}.results must be a non-empty list")
    indexed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, result in enumerate(results):
        case, corpus_identity, row = _result_key(result, f"{label}.results[{index}]")
        key = (case, corpus_identity)
        if key in indexed:
            raise AbbaSummaryInputError(
                f"{label}.results contains duplicate case/corpus identity for {case!r}"
            )
        indexed[key] = row
    return indexed


def _required_nonempty_string(
    value: Any, location: str, field: str, *, allow_none: bool = False
) -> str | None:
    if allow_none and value is None:
        return None
    if not isinstance(value, str) or not value:
        suffix = " or null" if allow_none else ""
        raise AbbaSummaryInputError(f"{location}.{field} must be a non-empty string{suffix}")
    return value


def _optional_string(value: Any, location: str, field: str) -> str | None:
    if value is not None and not isinstance(value, str):
        raise AbbaSummaryInputError(f"{location}.{field} must be a string or null")
    return value


def _required_bool(value: Any, location: str, field: str) -> bool:
    if not isinstance(value, bool):
        raise AbbaSummaryInputError(f"{location}.{field} must be a boolean")
    return value


def _required_positive_integer(value: Any, location: str, field: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise AbbaSummaryInputError(f"{location}.{field} must be a positive integer")
    return value


def _required_string_list(value: Any, location: str, field: str) -> list[str]:
    if (
        not isinstance(value, list)
        or not value
        or any(not isinstance(item, str) or not item for item in value)
        or len(value) != len(set(value))
    ):
        raise AbbaSummaryInputError(
            f"{location}.{field} must be a non-empty list of unique strings"
        )
    return value


def _validate_tool(tool: dict[str, Any], label: str) -> None:
    for field in REQUIRED_TOOL_FIELDS:
        _required_nonempty_string(tool.get(field), f"{label}.tool", field)
    if tool["name"] != HARNESS_TOOL_NAME:
        raise AbbaSummaryInputError(
            f"{label}.tool.name must be {HARNESS_TOOL_NAME!r}"
        )


def _validate_environment(environment: dict[str, Any], label: str) -> None:
    location = f"{label}.environment"
    for field in REQUIRED_ENVIRONMENT_FIELDS:
        if field not in environment:
            raise AbbaSummaryInputError(f"{location}.{field} is required")
    _required_nonempty_string(environment["rustc_version"], location, "rustc_version")
    revision = _required_nonempty_string(environment["git_revision"], location, "git_revision")
    if revision is None:
        raise AbbaSummaryInputError(f"{location}.git_revision must not be null")
    if _required_bool(environment["git_worktree_dirty"], location, "git_worktree_dirty"):
        raise AbbaSummaryInputError(f"{location}.git_worktree_dirty must be false")
    _required_positive_integer(
        environment["logical_cpus_available"], location, "logical_cpus_available"
    )
    _required_nonempty_string(environment["allocator"], location, "allocator")
    for field in ("rustflags", "cargo_build_target", "perf_event_paranoid"):
        _optional_string(environment[field], location, field)
    for field in (
        "os",
        "kernel",
        "cpu_model",
        "filesystem_type",
        "cpu_affinity",
        "storage_identifier",
    ):
        if field in environment:
            _optional_string(environment[field], location, field)
    for field in ("total_memory_bytes", "page_size_bytes"):
        if field in environment and environment[field] is not None:
            _u64(environment[field], f"{location}.{field}")
    if "source_destination_same_device" in environment and environment[
        "source_destination_same_device"
    ] is not None:
        _required_bool(
            environment["source_destination_same_device"],
            location,
            "source_destination_same_device",
        )


def _validate_configuration(configuration: dict[str, Any], label: str) -> None:
    location = f"{label}.configuration"
    for field in REQUIRED_CONFIGURATION_FIELDS:
        if field not in configuration:
            raise AbbaSummaryInputError(f"{location}.{field} is required")
    samples = _required_positive_integer(
        configuration["samples_per_case"], location, "samples_per_case"
    )
    if samples < MIN_RETAINED_SAMPLES:
        raise AbbaSummaryInputError(
            f"{location}.samples_per_case must be at least {MIN_RETAINED_SAMPLES}"
        )
    _required_positive_integer(
        configuration["warmup_iterations_per_case"],
        location,
        "warmup_iterations_per_case",
    )
    _required_string_list(configuration["cases"], location, "cases")
    _required_string_list(configuration["corpus_shapes"], location, "corpus_shapes")
    for field, value in configuration.items():
        if field.endswith("_shapes") and field != "corpus_shapes":
            _required_string_list(value, location, field)
    if "filesystem_root_selected" in configuration:
        _required_bool(
            configuration["filesystem_root_selected"],
            location,
            "filesystem_root_selected",
        )


def validate_parallel_metrics(report: Any, label: str = "report") -> None:
    """Validate an emitted descriptive parallel-metrics envelope."""

    try:
        perf_compare.validate_parallel_metrics(report, label)
    except perf_compare.ComparisonInputError as error:
        raise AbbaSummaryInputError(str(error)) from error


def _validate_report(
    report: Any, label: str
) -> tuple[
    int,
    dict[str, Any],
    dict[str, Any],
    dict[str, Any],
    dict[tuple[str, str], dict[str, Any]],
    str,
    frozenset[str],
]:
    root = _require_object(report, label)
    canonical_report = _canonical_json(root, f"{label}.report")
    report_sha256 = hashlib.sha256(canonical_report.encode("utf-8")).hexdigest()
    schema_version = root.get("schema_version")
    if isinstance(schema_version, bool) or not isinstance(schema_version, int):
        raise AbbaSummaryInputError(f"{label}.schema_version must be an integer")
    if schema_version != HARNESS_SCHEMA_VERSION:
        raise AbbaSummaryInputError(
            f"{label}.schema_version must be {HARNESS_SCHEMA_VERSION!r}"
        )
    validate_parallel_metrics(root, label)
    tool = _require_object(root.get("tool"), f"{label}.tool")
    if not tool:
        raise AbbaSummaryInputError(f"{label}.tool must not be empty")
    _validate_tool(tool, label)
    environment = _require_object(root.get("environment"), f"{label}.environment")
    if not environment:
        raise AbbaSummaryInputError(f"{label}.environment must not be empty")
    _validate_environment(environment, label)
    configuration = _require_object(root.get("configuration"), f"{label}.configuration")
    if not configuration:
        raise AbbaSummaryInputError(f"{label}.configuration must not be empty")
    _validate_configuration(configuration, label)
    filesystem_evidence = root.get("filesystem_evidence")
    filesystem_shapes: set[str] = set()
    if filesystem_evidence is not None and not isinstance(filesystem_evidence, list):
        raise AbbaSummaryInputError(f"{label}.filesystem_evidence must be a list")
    if isinstance(filesystem_evidence, list):
        for index, evidence in enumerate(filesystem_evidence):
            evidence_object = _require_object(
                evidence, f"{label}.filesystem_evidence[{index}]"
            )
            corpus = _require_object(
                evidence_object.get("corpus"),
                f"{label}.filesystem_evidence[{index}].corpus",
            )
            shape = corpus.get("shape")
            if not isinstance(shape, str) or not shape:
                raise AbbaSummaryInputError(
                    f"{label}.filesystem_evidence[{index}].corpus.shape "
                    "must be a non-empty string"
                )
            filesystem_shapes.add(shape)
    return (
        schema_version,
        tool,
        environment,
        configuration,
        _index_results(root, label),
        report_sha256,
        frozenset(filesystem_shapes),
    )


def _stable_environment(environment: dict[str, Any]) -> dict[str, Any]:
    """Return environment facts expected to remain fixed across ABBA legs."""

    return {
        key: value for key, value in environment.items() if key not in ENVIRONMENT_VARIANTS
    }


def _validate_configuration_rows(
    configuration: dict[str, Any],
    indexed: Mapping[tuple[str, str], dict[str, Any]],
    label: str,
    *,
    filesystem_shapes: Iterable[str] = (),
) -> None:
    """Check optional harness cardinality/selector declarations when present."""

    cases = configuration["cases"]
    actual_cases = {case for case, _ in indexed}
    if actual_cases != set(cases):
        raise AbbaSummaryInputError(
            f"{label}.configuration.cases does not match result cases"
        )
    declared_shapes = {
        shape
        for field, values in configuration.items()
        if field.endswith("_shapes")
        for shape in values
    }
    actual_shapes: set[str] = set()
    for case, corpus_identity in indexed:
        corpus = json.loads(corpus_identity)
        shape = corpus.get("shape")
        if not isinstance(shape, str) or not shape:
            raise AbbaSummaryInputError(
                f"{label}.results case {case!r} corpus.shape must be a non-empty string"
            )
        actual_shapes.add(shape)
    filesystem_shape_set = set(filesystem_shapes)
    filesystem_exception = bool(filesystem_shape_set) and actual_shapes.issubset(
        filesystem_shape_set
    )
    if not actual_shapes.issubset(declared_shapes) and not filesystem_exception:
        raise AbbaSummaryInputError(
            f"{label}.configuration shape declarations do not cover result shapes"
        )
    samples_per_case = configuration.get("samples_per_case")
    for key, row in indexed.items():
        case = key[0]
        elapsed = _require_object(row.get("elapsed_ns"), f"{label}.{case}.elapsed_ns")
        samples = elapsed.get("samples")
        if not isinstance(samples, list) or len(samples) != samples_per_case:
            raise AbbaSummaryInputError(
                f"{label}.{case}.elapsed_ns.samples does not match samples_per_case"
            )


def _validate_drift_ceilings(value: Mapping[str, Any] | None) -> dict[str, float]:
    if value is None:
        return dict(DEFAULT_DRIFT_CEILINGS)
    if not isinstance(value, Mapping):
        raise AbbaSummaryInputError("drift ceilings must be an object")
    if set(value) != set(STATISTICS):
        raise AbbaSummaryInputError(
            f"drift ceilings must contain exactly {list(STATISTICS)!r}"
        )
    ceilings = {
        name: _finite_number(value[name], f"drift ceilings.{name}")
        for name in STATISTICS
    }
    if any(ceiling < 0 for ceiling in ceilings.values()):
        raise AbbaSummaryInputError("drift ceilings must be non-negative")
    return ceilings


def _identity_value(row: dict[str, Any], field: str, location: str) -> tuple[bool, str]:
    # Harness Option fields may be serialized as explicit null.  Null carries
    # no identity, so it is absence rather than a verified-equal payload.
    present = field in row and row[field] is not None
    value = row[field] if present else None
    return present, _canonical_json(value, f"{location}.{field}")


def _compare_row_identity(
    rows: Mapping[str, dict[str, Any]],
    field: str,
    location: str,
) -> tuple[str, bool, str | None, Any]:
    identities = {
        label: _identity_value(rows[label], field, f"{label}.{location}")
        for label in LEG_ORDER
    }
    expected_presence = identities["a1"][0]
    if any(identity[0] != expected_presence for identity in identities.values()):
        raise AbbaSummaryInputError(
            f"{location} {field} presence differs between ABBA legs"
        )
    if not expected_presence:
        return "consistently_absent", False, None, None
    expected_identity = identities["a1"][1]
    if any(identity[1] != expected_identity for identity in identities.values()):
        raise AbbaSummaryInputError(
            f"{location} {field} identity differs between ABBA legs"
        )
    value = rows["a1"][field]
    return "verified_equal", True, expected_identity, value


def _validate_output_sha256(value: Any, location: str) -> str:
    if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
        raise AbbaSummaryInputError(f"{location} must be a lowercase SHA-256 hex string")
    return value


def _coerce_reports(
    a1: Mapping[str, Any] | Sequence[Mapping[str, Any]],
    b1: Mapping[str, Any] | None,
    b2: Mapping[str, Any] | None,
    a2: Mapping[str, Any] | None,
) -> tuple[Mapping[str, Any], Mapping[str, Any], Mapping[str, Any], Mapping[str, Any]]:
    if b1 is None and b2 is None and a2 is None:
        if isinstance(a1, Mapping):
            for labels in (LEG_ORDER, tuple(label.upper() for label in LEG_ORDER)):
                if all(label in a1 for label in labels):
                    return tuple(a1[label] for label in labels)  # type: ignore[return-value]
            raise AbbaSummaryInputError(
                f"reports mapping must contain {list(LEG_ORDER)!r}"
            )
        if isinstance(a1, Sequence) and not isinstance(a1, (str, bytes)) and len(a1) == 4:
            return tuple(a1)  # type: ignore[return-value]
        raise AbbaSummaryInputError("reports must be four reports in A1,B1,B2,A2 order")
    if b1 is None or b2 is None or a2 is None:
        raise AbbaSummaryInputError("all four ABBA reports are required")
    return a1, b1, b2, a2


def _parse_selectors(values: Iterable[str] | None) -> set[str] | None:
    if values is None:
        return None
    if isinstance(values, str):
        values = (values,)
    selectors: set[str] = set()
    for value in values:
        if not isinstance(value, str):
            raise AbbaSummaryInputError("selectors must contain only strings")
        for selector in value.split(","):
            selector = selector.strip()
            if selector:
                selectors.add(selector)
    if not selectors:
        raise AbbaSummaryInputError("selectors must contain at least one non-empty value")
    return selectors


def _delta_percent(control: float, candidate: float) -> float:
    if control <= 0:
        raise AbbaSummaryInputError("elapsed statistics must have a positive control value")
    delta = (control - candidate) / control * 100.0
    if not math.isfinite(delta):
        raise AbbaSummaryInputError("candidate reduction is not finite")
    return delta


def _drift_percent(first: float, second: float) -> float:
    if first <= 0:
        raise AbbaSummaryInputError("elapsed statistics must have positive drift baselines")
    drift = (second - first) / first * 100.0
    if not math.isfinite(drift):
        raise AbbaSummaryInputError("same-implementation drift is not finite")
    return drift


def _result_summary(
    rows: Mapping[str, dict[str, Any]],
    *,
    case: str,
    corpus_identity: str,
    drift_ceilings: Mapping[str, float],
) -> dict[str, Any]:
    source_status, source_present, source_identity, source_value = _compare_row_identity(
        rows, "source", f"{case}[{corpus_identity}]"
    )
    sink_status, sink_present, sink_identity, sink_value = _compare_row_identity(
        rows, "sink", f"{case}[{corpus_identity}]"
    )
    for label in LEG_ORDER:
        if "output_sha256" in rows[label] and rows[label]["output_sha256"] is not None:
            _validate_output_sha256(
                rows[label]["output_sha256"], f"{label}.{case}.output_sha256"
            )
    output_status, output_present, output_identity, output_value = _compare_row_identity(
        rows, "output_sha256", f"{case}[{corpus_identity}]"
    )

    elapsed: dict[str, dict[str, Any]] = {}
    for label in LEG_ORDER:
        elapsed[label] = recompute_statistics(
            rows[label].get("elapsed_ns"), f"{label}.{case}.elapsed_ns"
        )
    sample_counts = {elapsed[label]["sample_count"] for label in LEG_ORDER}
    if len(sample_counts) != 1:
        raise AbbaSummaryInputError(
            f"{case}[{corpus_identity}] ABBA legs have different sample counts"
        )

    candidate_reduction = {
        "a1_to_b1": {
            name: _delta_percent(elapsed["a1"][name], elapsed["b1"][name])
            for name in STATISTICS
        },
        "a2_to_b2": {
            name: _delta_percent(elapsed["a2"][name], elapsed["b2"][name])
            for name in STATISTICS
        },
    }
    adverse_both: list[str] = []
    drift = {
        "control": {
            name: _drift_percent(elapsed["a1"][name], elapsed["a2"][name])
            for name in STATISTICS
        },
        "candidate": {
            name: _drift_percent(elapsed["b1"][name], elapsed["b2"][name])
            for name in STATISTICS
        },
    }
    drift_within_ceiling = {
        implementation: {
            name: abs(drift[implementation][name]) <= drift_ceilings[name]
            for name in STATISTICS
        }
        for implementation in ("control", "candidate")
    }
    accepted: list[str] = []
    rejected: dict[str, str] = {}
    for name in STATISTICS:
        reasons: list[str] = []
        first_reduction = candidate_reduction["a1_to_b1"][name]
        second_reduction = candidate_reduction["a2_to_b2"][name]
        if first_reduction < 0 and second_reduction < 0:
            adverse_both.append(name)
            reasons.append("candidate is not lower in both paired directions")
        elif first_reduction <= 0 or second_reduction <= 0:
            if (first_reduction < 0 < second_reduction) or (
                second_reduction < 0 < first_reduction
            ):
                reasons.append("paired directions disagree")
            else:
                reasons.append("candidate is not lower in both paired directions")
        for implementation in ("control", "candidate"):
            if not drift_within_ceiling[implementation][name]:
                reasons.append(
                    f"{implementation} drift {drift[implementation][name]:+.6f}% "
                    f"exceeds {drift_ceilings[name]:g}% ceiling"
                )
        if reasons:
            rejected[name] = "; ".join(reasons)
        else:
            accepted.append(name)

    corpus = json.loads(corpus_identity)
    return {
        "case": case,
        "shape": corpus.get("shape"),
        "corpus": corpus,
        "source": source_value,
        "sink": sink_value,
        "output_sha256": output_value,
        "identity": {
            "corpus": corpus_identity,
            "source_present": source_present,
            "source_canonical_json": source_identity,
            "source_status": source_status,
            "sink_present": sink_present,
            "sink_canonical_json": sink_identity,
            "sink_status": sink_status,
            "output_sha256_status": output_status,
            "output_sha256_canonical_json": output_identity,
        },
        "elapsed_ns": {
            "sample_count": next(iter(sample_counts)),
            "legs_ns": elapsed,
            "candidate_reduction_percent": candidate_reduction,
            "same_implementation_drift_percent": drift,
            "drift_ceiling_percent": dict(drift_ceilings),
            "drift_within_ceiling": drift_within_ceiling,
            "adverse_both_statistics": adverse_both,
            "accepted_statistics": accepted,
            "rejected_statistics": rejected,
        },
    }


def summarize_reports(
    a1: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
    b1: Mapping[str, Any] | None = None,
    b2: Mapping[str, Any] | None = None,
    a2: Mapping[str, Any] | None = None,
    *,
    drift_ceilings: Mapping[str, Any] | None = None,
    cases: Iterable[str] | None = None,
    shapes: Iterable[str] | None = None,
    reports: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
    ceilings: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate four reports and return a deterministic machine-readable summary.

    ``a1`` may be a four-item sequence, a mapping keyed by ``a1``, ``b1``,
    ``b2`` and ``a2``, or the first report when the remaining three reports are
    passed as positional arguments.
    """

    if reports is not None:
        if a1 is not None or b1 is not None or b2 is not None or a2 is not None:
            raise AbbaSummaryInputError("reports cannot be combined with positional ABBA reports")
        a1 = reports
    if a1 is None:
        raise AbbaSummaryInputError("four ABBA reports are required")
    if drift_ceilings is not None and ceilings is not None:
        raise AbbaSummaryInputError("use drift_ceilings or ceilings, not both")
    ceiling_values = _validate_drift_ceilings(
        drift_ceilings if drift_ceilings is not None else ceilings
    )
    report_values = _coerce_reports(a1, b1, b2, a2)
    validated = {
        label: _validate_report(report, label)
        for label, report in zip(LEG_ORDER, report_values)
    }
    schema_versions = {item[0] for item in validated.values()}
    if len(schema_versions) != 1:
        raise AbbaSummaryInputError("harness schema_version differs between ABBA legs")
    tool_identities = {
        _canonical_json(item[1], f"{label}.tool") for label, item in validated.items()
    }
    if len(tool_identities) != 1:
        raise AbbaSummaryInputError("harness tool identity differs between ABBA legs")
    environments_by_leg = {label: item[2] for label, item in validated.items()}
    control_revision = environments_by_leg["a1"]["git_revision"]
    if control_revision != environments_by_leg["a2"]["git_revision"]:
        raise AbbaSummaryInputError("control A1/A2 git_revision differs")
    candidate_revision = environments_by_leg["b1"]["git_revision"]
    if candidate_revision != environments_by_leg["b2"]["git_revision"]:
        raise AbbaSummaryInputError("candidate B1/B2 git_revision differs")
    if control_revision == candidate_revision:
        raise AbbaSummaryInputError(
            "control and candidate git_revision must be distinct"
        )
    environment_identities = {
        label: _canonical_json(item[2], f"{label}.environment")
        for label, item in validated.items()
    }
    stable_environment_identities = {
        label: _canonical_json(_stable_environment(item[2]), f"{label}.environment")
        for label, item in validated.items()
    }
    if len(set(stable_environment_identities.values())) != 1:
        raise AbbaSummaryInputError(
            "stable environment identity differs between ABBA legs"
        )
    configurations = {
        _canonical_json(item[3], f"{label}.configuration") for label, item in validated.items()
    }
    if len(configurations) != 1:
        raise AbbaSummaryInputError("harness configuration differs between ABBA legs")

    for label, (_, _, _, configuration, indexed, _, filesystem_shapes) in validated.items():
        _validate_configuration_rows(
            configuration,
            indexed,
            label,
            filesystem_shapes=filesystem_shapes,
        )

    result_sets = {frozenset(item[4]) for item in validated.values()}
    if len(result_sets) != 1:
        raise AbbaSummaryInputError("case/corpus result identities differ between ABBA legs")
    selected_cases = _parse_selectors(cases)
    selected_shapes = _parse_selectors(shapes)
    first_index = validated["a1"][4]
    selected_keys = []
    for case, corpus_identity in sorted(first_index):
        corpus = json.loads(corpus_identity)
        if selected_cases is not None and case not in selected_cases:
            continue
        shape = corpus.get("shape")
        if selected_shapes is not None and not isinstance(shape, str):
            raise AbbaSummaryInputError(
                f"{case}[{corpus_identity}] cannot be selected by shape without a string shape"
            )
        if selected_shapes is not None and shape not in selected_shapes:
            continue
        selected_keys.append((case, corpus_identity))
    if not selected_keys:
        raise AbbaSummaryInputError("selectors did not match any case/corpus result")

    # Validate every row before applying selectors.  A selector controls what
    # is emitted; it must not hide an unsafe mismatch in another report row.
    all_summaries: dict[tuple[str, str], dict[str, Any]] = {}
    for case, corpus_identity in sorted(first_index):
        rows = {label: validated[label][4][(case, corpus_identity)] for label in LEG_ORDER}
        all_summaries[(case, corpus_identity)] = _result_summary(
            rows,
            case=case,
            corpus_identity=corpus_identity,
            drift_ceilings=ceiling_values,
        )

    results = [all_summaries[key] for key in selected_keys]

    tool = json.loads(next(iter(tool_identities)))
    configuration = json.loads(next(iter(configurations)))
    environments = {
        label: json.loads(environment_identities[label]) for label in LEG_ORDER
    }
    stable_environment = json.loads(stable_environment_identities["a1"])
    report_identities = {
        label: {"canonical_sha256": validated[label][5]} for label in LEG_ORDER
    }
    status_counts = {
        field: {
            status: sum(
                result["identity"][f"{field}_status"] == status for result in results
            )
            for status in ("verified_equal", "consistently_absent")
        }
        for field in ("source", "sink")
    }
    output_status_counts = {
        status: sum(
            result["identity"]["output_sha256_status"] == status for result in results
        )
        for status in ("verified_equal", "consistently_absent")
    }
    return {
        "schema_version": SCHEMA_VERSION,
        "tool": {"name": TOOL_NAME, "version": TOOL_VERSION},
        "protocol": {
            "order": ["a1_control", "b1_candidate", "b2_candidate", "a2_control"],
            "statistics": list(STATISTICS),
            "drift_ceiling_percent": ceiling_values,
            "percentiles": (
                "p50 = Rust u64 floor midpoint; p95/p99 = integer nearest-rank"
            ),
            "dispersion": "sample standard deviation via Rust Welford update",
            "uncertainty": (
                "95% two-sided Student's t interval for the mean; embedded harness "
                "critical-value table with Cornish-Fisher tail"
            ),
        },
        "harness_identity": {
            "schema_version": next(iter(schema_versions)),
            "tool": tool,
            "configuration": configuration,
        },
        "environment": {
            "stable": stable_environment,
            "legs": environments,
        },
        "implementation_identity": {
            "control": {"git_revision": control_revision, "legs": ["a1", "a2"]},
            "candidate": {
                "git_revision": candidate_revision,
                "legs": ["b1", "b2"],
            },
            "distinct": True,
        },
        "report_identity": report_identities,
        "results": results,
        "verification": {
            "result_count": len(results),
            "tool_identity_verified": True,
            "configuration_identity_verified": True,
            "environment_stable_identity_verified": True,
            "environment_legs_recorded": True,
            "source_identity": status_counts["source"],
            "sink_identity": status_counts["sink"],
            "output_sha256_identity": output_status_counts,
            "source_identity_verified": status_counts["source"]["verified_equal"] == len(results),
            "sink_identity_verified": status_counts["sink"]["verified_equal"] == len(results),
            "output_sha256_identity_verified": output_status_counts["verified_equal"]
            == len(results),
            "case_corpus_identity_verified": True,
            "statistics_recomputed_from_samples": True,
        },
    }


build_summary = summarize_reports
summarize_abba = summarize_reports


def load_report(path: Path) -> dict[str, Any]:
    def reject_nonfinite(value: str) -> None:
        raise ValueError(f"non-finite JSON value {value!r}")

    def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        value: dict[str, Any] = {}
        for key, item in pairs:
            if key in value:
                raise ValueError(f"duplicate JSON object key {key!r}")
            value[key] = item
        return value

    try:
        with path.open(encoding="utf-8") as handle:
            value = json.load(
                handle,
                object_pairs_hook=reject_duplicate_keys,
                parse_constant=reject_nonfinite,
            )
    except (OSError, ValueError) as error:
        raise AbbaSummaryInputError(f"cannot read {path}: {error}") from error
    return _require_object(value, str(path))


def _parse_ceiling_argument(value: str) -> dict[str, float]:
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError:
        parsed = {}
        for item in value.split(","):
            name, separator, number = item.partition("=")
            if not separator:
                raise AbbaSummaryInputError(
                    "--drift-ceilings must be JSON or comma-separated name=value pairs"
                )
            parsed[name.strip()] = float(number)
    if not isinstance(parsed, dict):
        raise AbbaSummaryInputError("--drift-ceilings must describe an object")
    return _validate_drift_ceilings(parsed)


def _write_json(path: Path, value: dict[str, Any]) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
    except FileExistsError as error:
        raise AbbaSummaryInputError(f"output already exists: {path}") from error
    except OSError as error:
        raise AbbaSummaryInputError(f"cannot write {path}: {error}") from error


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "reports",
        nargs="*",
        type=Path,
        metavar="REPORT",
        help="four reports in A1,B1,B2,A2 order",
    )
    for label in LEG_ORDER:
        aliases = {
            "a1": ("--a1", "--control-a", "--before-a"),
            "b1": ("--b1", "--candidate-a", "--after-a"),
            "b2": ("--b2", "--candidate-b", "--after-b"),
            "a2": ("--a2", "--control-b", "--before-b"),
        }[label]
        parser.add_argument(*aliases, type=Path, help=f"{label.upper()} harness report")
    parser.add_argument("--json-out", "--output", type=Path)
    parser.add_argument(
        "--drift-ceilings",
        help="JSON object or comma-separated p50=5,mean=5,p95=10,p99=15",
    )
    parser.add_argument("--case", dest="cases", action="append")
    parser.add_argument("--shape", dest="shapes", action="append")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    try:
        named = [getattr(arguments, label) for label in LEG_ORDER]
        if arguments.reports and any(path is not None for path in named):
            raise AbbaSummaryInputError("use positional reports or --a1/--b1/--b2/--a2, not both")
        if arguments.reports:
            if len(arguments.reports) != 4:
                raise AbbaSummaryInputError("exactly four positional reports are required")
            paths = arguments.reports
        else:
            if any(path is None for path in named):
                raise AbbaSummaryInputError("all four --a1/--b1/--b2/--a2 reports are required")
            paths = named
        reports = [load_report(path) for path in paths]
        ceilings = (
            _parse_ceiling_argument(arguments.drift_ceilings)
            if arguments.drift_ceilings is not None
            else None
        )
        summary = summarize_reports(
            reports,
            drift_ceilings=ceilings,
            cases=arguments.cases,
            shapes=arguments.shapes,
        )
        if arguments.json_out is not None:
            _write_json(arguments.json_out, summary)
        json.dump(summary, sys.stdout, indent=2, sort_keys=True, allow_nan=False)
        sys.stdout.write("\n")
        return 0
    except (OSError, AbbaSummaryInputError, ValueError) as error:
        print(f"{TOOL_NAME}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
