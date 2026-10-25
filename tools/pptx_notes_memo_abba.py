#!/usr/bin/env python3
"""Fail-closed ABBA validator for the opt-in PPTX notes-memo probe.

The regular ``perf_abba_summary.py`` intentionally has a closed catalog of
the default 213-row matrix.  This wrapper owns the small local probe schema,
but keeps the same four-leg order: ``A1, B1, B2, A2``.  It validates source
and executable identity, deterministic fixture identity, native/allocator
metric applicability, exact sample ordering, and every untimed fidelity/raw
oracle gate before writing a descriptive summary.  It never computes a
speedup or accepts a report with fewer than four samples per round and four
complete rounds.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path
from typing import Any, Mapping


SCHEMA = "pptx_notes_memo_probe_v1"
SUMMARY_SCHEMA = "pptx_notes_memo_abba_summary_v1"
LEG_ORDER = ("a1", "b1", "b2", "a2")
CASES = (
    "pptx_notes_memo_cold_capture",
    "pptx_notes_memo_warm_notes_commit",
    "pptx_notes_memo_changed_slide_commit",
)
OPERATIONS = {
    CASES[0]: "cold_capture",
    CASES[1]: "warm_notes_commit",
    CASES[2]: "changed_slide_commit",
}
TIMING_SCOPES = {
    CASES[0]: "opened_presentation over a pre-parsed package; Package::from_bytes fixture setup and all fidelity checks are outside",
    CASES[1]: "opened Snapshot::edit, set_notes_text(1), and commit; source package setup, serialization, and fidelity checks are outside",
    CASES[2]: "opened Snapshot::edit, set_shape_text(0, 0), and commit; source package setup, serialization, and fidelity checks are outside",
}
SHA256 = re.compile(r"[0-9a-f]{64}\Z")
REVISION = re.compile(r"[0-9a-f]{40}\Z")
SAMPLES_PER_ROUND = 4
MIN_ROUNDS = 4
REQUIRED_SAMPLE_COUNT = SAMPLES_PER_ROUND * MIN_ROUNDS
MAX_U64 = (1 << 64) - 1
NATIVE_ALLOCATOR = "Rust system allocator"
INSTRUMENTED_ALLOCATOR = "CountingSystemAllocator(std::alloc::System)"
ALLOCATOR_COUNTER_REVISION = "serialized_region_peak_v3"
ALLOCATION_SCOPE = "operation_global_system_allocator"
REPORT_ALLOCATION_SCOPE = "one non-nested operation-scoped global-allocator region around each timed API operation"
MEMO_SCOPE = "source-audited SlideRootMemo eligibility; runtime proved-root counters remain internal and are covered by the focused crate test"
FIDELITY_SCOPE = "untimed semantic reopen plus exact three-slide/notes and raw ZIP member, relationship, master/theme, and root oracle after every retained operation"
FIXTURE_RECIPE = "three authored slides; title notes slide {i}, one body text box, speaker notes {i}"
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
ALLOCATION_NUMERIC_FIELDS = (
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
REPORT_KEYS = frozenset(
    {
        "schema",
        "tool",
        "binary_identity",
        "source_revision",
        "instrumentation",
        "allocator",
        "allocator_counter_revision",
        "timing_claim",
        "allocation_scope",
        "fixture",
        "samples",
        "warmup",
        "checked_iteration_count",
        "memo_scope",
        "fidelity_scope",
        "results",
    }
)
BINARY_KEYS = frozenset({"path", "binary_sha256", "binary_bytes", "mode_bits", "executable", "profile"})
FIXTURE_KEYS = frozenset({"generator", "slides", "notes_slides", "authored_shapes_per_slide", "bytes", "sha256", "recipe"})
STATISTICS_KEYS = frozenset(
    {
        "unit",
        "samples",
        "sample_order",
        "min",
        "p50",
        "p95",
        "p99",
        "max",
        "mean",
        "standard_deviation",
        "confidence_interval_95",
    }
)
CONFIDENCE_INTERVAL_KEYS = frozenset({"method", "lower", "upper"})
ROW_KEYS = frozenset({"case", "operation", "timing_scope", "elapsed_ns", "observations"})
OBSERVATION_KEYS = frozenset(
    {"sample_index", "elapsed_ns", "output_bytes", "output_sha256", "fidelity", "allocation"}
)
SEMANTIC_KEYS = frozenset(
    {"verified", "slide_count", "notes_slide_count", "slide_text_sha256", "notes_text_sha256"}
)
FIDELITY_KEYS = frozenset(
    {
        "semantic",
        "changed_commit_reported",
        "output_reopened",
        "changed_slide_index",
        "untouched_slide_root_count",
        "raw",
    }
)
FIDELITY_EDIT_KEYS = FIDELITY_KEYS | {"source_immutable"}
RAW_KEYS = frozenset(
    {
        "source_member_count",
        "output_member_count",
        "source_payload_sha256",
        "output_payload_sha256",
        "member_names_equal",
        "untouched_payloads_equal",
        "relationship_member_count",
        "relationship_names_equal",
        "relationship_payloads_equal",
        "master_theme_member_count",
        "master_theme_payloads_equal",
        "changed_members",
        "expected_changed_members",
        "exact_changed_member_set",
        "changed_payloads_nonempty",
        "slide_root_count",
        "notes_root_count",
        "changed_slide_index",
        "changed_notes_slide_index",
        "untouched_slide_root_count",
    }
)
ALLOCATION_KEYS = frozenset({"status", "scope"}) | frozenset(ALLOCATION_NUMERIC_FIELDS)


class ValidationError(ValueError):
    """A report failed a protocol identity, metric, or fidelity gate."""


def _strict_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    """Build JSON objects without silently accepting duplicate member names."""

    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValidationError(f"duplicate JSON object member {key!r}")
        result[key] = value
    return result


def _reject_nonfinite_json_constant(value: str) -> None:
    raise ValidationError(f"non-finite JSON constant {value!r} is not permitted")


def load_strict_json(text: str, path: str) -> Any:
    """Parse protocol JSON with duplicate-key and non-finite-number rejection."""

    try:
        return json.loads(
            text,
            object_pairs_hook=_strict_object,
            parse_constant=_reject_nonfinite_json_constant,
        )
    except ValidationError:
        raise
    except (UnicodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot parse {path}: {error}") from error


def _require_mapping(value: Any, path: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise ValidationError(f"{path} must be an object")
    return value


def _exact_keys(value: Mapping[str, Any], expected: frozenset[str], path: str) -> None:
    actual = set(value)
    missing = sorted(expected - actual)
    unknown = sorted(actual - expected)
    if missing or unknown:
        details = []
        if missing:
            details.append(f"missing={missing}")
        if unknown:
            details.append(f"unknown={unknown}")
        raise ValidationError(f"{path} has an invalid key set ({'; '.join(details)})")


def _required(value: Mapping[str, Any], key: str, path: str) -> Any:
    if key not in value:
        raise ValidationError(f"{path}.{key} is required")
    return value[key]


def _string(value: Any, path: str, pattern: re.Pattern[str] | None = None) -> str:
    if not isinstance(value, str) or not value:
        raise ValidationError(f"{path} must be a non-empty string")
    if pattern is not None and pattern.fullmatch(value) is None:
        raise ValidationError(f"{path} has the wrong identity format")
    return value


def _integer(value: Any, path: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ValidationError(f"{path} must be an integer >= {minimum}")
    return value


def _u64(value: Any, path: str, minimum: int = 0) -> int:
    result = _integer(value, path, minimum)
    if result > MAX_U64:
        raise ValidationError(f"{path} must fit in an unsigned 64-bit counter")
    return result


def _boolean(value: Any, path: str) -> bool:
    if not isinstance(value, bool):
        raise ValidationError(f"{path} must be boolean")
    return value


def _validate_binary(report: Mapping[str, Any], path: str) -> dict[str, Any]:
    identity = _require_mapping(_required(report, "binary_identity", path), f"{path}.binary_identity")
    _exact_keys(identity, BINARY_KEYS, f"{path}.binary_identity")
    result = {
        "path": _string(_required(identity, "path", f"{path}.binary_identity"), f"{path}.binary_identity.path"),
        "binary_sha256": _string(
            _required(identity, "binary_sha256", f"{path}.binary_identity"),
            f"{path}.binary_identity.binary_sha256",
            SHA256,
        ),
        "binary_bytes": _u64(
            _required(identity, "binary_bytes", f"{path}.binary_identity"),
            f"{path}.binary_identity.binary_bytes",
            1,
        ),
        "mode_bits": _required(identity, "mode_bits", f"{path}.binary_identity"),
        "executable": _boolean(
            _required(identity, "executable", f"{path}.binary_identity"),
            f"{path}.binary_identity.executable",
        ),
        "profile": _string(
            _required(identity, "profile", f"{path}.binary_identity"),
            f"{path}.binary_identity.profile",
        ),
    }
    if result["mode_bits"] is not None:
        mode_bits = _integer(result["mode_bits"], f"{path}.binary_identity.mode_bits")
        if mode_bits > 0o7777:
            raise ValidationError(f"{path}.binary_identity.mode_bits is outside the executable mode range")
    if not Path(result["path"]).is_absolute():
        raise ValidationError(f"{path}.binary_identity.path must be absolute")
    if not result["executable"] or result["profile"] != "release":
        raise ValidationError(f"{path}.binary_identity must be an executable release binary")
    return result


def _validate_fixture(report: Mapping[str, Any], path: str) -> dict[str, Any]:
    fixture = _require_mapping(_required(report, "fixture", path), f"{path}.fixture")
    _exact_keys(fixture, FIXTURE_KEYS, f"{path}.fixture")
    result = {
        "generator": _string(_required(fixture, "generator", f"{path}.fixture"), f"{path}.fixture.generator"),
        "slides": _integer(_required(fixture, "slides", f"{path}.fixture"), f"{path}.fixture.slides"),
        "notes_slides": _integer(
            _required(fixture, "notes_slides", f"{path}.fixture"), f"{path}.fixture.notes_slides"
        ),
        "authored_shapes_per_slide": _integer(
            _required(fixture, "authored_shapes_per_slide", f"{path}.fixture"),
            f"{path}.fixture.authored_shapes_per_slide",
        ),
        "bytes": _integer(_required(fixture, "bytes", f"{path}.fixture"), f"{path}.fixture.bytes", 1),
        "sha256": _string(_required(fixture, "sha256", f"{path}.fixture"), f"{path}.fixture.sha256", SHA256),
        "recipe": _string(_required(fixture, "recipe", f"{path}.fixture"), f"{path}.fixture.recipe"),
    }
    if result["generator"] != "pptx-notes-memo-three-slide-v1":
        raise ValidationError(f"{path}.fixture.generator is not the pinned three-slide fixture")
    if result["slides"] != 3 or result["notes_slides"] != 3 or result["authored_shapes_per_slide"] != 1:
        raise ValidationError(f"{path}.fixture shape/count identity is not the pinned recipe")
    if result["recipe"] != FIXTURE_RECIPE:
        raise ValidationError(f"{path}.fixture.recipe is not the pinned fixture recipe")
    return result


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


def _float_close(reported: Any, expected: float, path: str) -> None:
    if isinstance(reported, bool) or not isinstance(reported, (int, float)) or not math.isfinite(reported):
        raise ValidationError(f"{path} must be finite numeric evidence")
    tolerance = max(1e-12, abs(expected) * 1e-12)
    if abs(float(reported) - expected) > tolerance:
        raise ValidationError(f"{path} does not match the samples")


def _validate_statistics(value: Any, path: str, sample_count: int) -> tuple[list[int], list[int]]:
    stats = _require_mapping(value, path)
    _exact_keys(stats, STATISTICS_KEYS, path)
    if _required(stats, "unit", path) != "ns":
        raise ValidationError(f"{path}.unit must be ns")
    samples_value = _required(stats, "samples", path)
    if not isinstance(samples_value, list) or len(samples_value) != sample_count:
        raise ValidationError(f"{path}.samples must contain exactly {sample_count} values")
    samples = [_u64(item, f"{path}.samples[{index}]", 1) for index, item in enumerate(samples_value)]
    order_value = _required(stats, "sample_order", path)
    if not isinstance(order_value, list) or len(order_value) != sample_count:
        raise ValidationError(f"{path}.sample_order must contain exactly {sample_count} values")
    order = [_integer(item, f"{path}.sample_order[{index}]") for index, item in enumerate(order_value)]
    if sorted(order) != list(range(sample_count)):
        raise ValidationError(f"{path}.sample_order is not a permutation")
    if samples != sorted(samples):
        raise ValidationError(f"{path}.samples is not sorted")
    mean = 0.0
    squared_deviation_sum = 0.0
    for index, sample in enumerate(samples):
        value = float(sample)
        next_count = float(index + 1)
        delta = value - mean
        next_mean = mean + delta / next_count
        squared_deviation_sum += delta * (value - next_mean)
        mean = next_mean
    standard_deviation = (
        math.sqrt(squared_deviation_sum / float(sample_count - 1))
        if sample_count > 1
        else 0.0
    )
    margin = (
        _student_t_critical_95(sample_count - 1)
        * standard_deviation
        / math.sqrt(float(sample_count))
        if sample_count > 1
        else 0.0
    )
    left = samples[(sample_count - 1) // 2]
    right = samples[sample_count // 2]
    expected = {
        "min": samples[0],
        "max": samples[-1],
        "p50": left // 2 + right // 2 + (left % 2 + right % 2) // 2,
        "p95": samples[max(1, (95 * sample_count + 99) // 100) - 1],
        "p99": samples[max(1, (99 * sample_count + 99) // 100) - 1],
    }
    for key, expected_value in expected.items():
        if _required(stats, key, path) != expected_value:
            raise ValidationError(f"{path}.{key} does not match the samples")
    _float_close(_required(stats, "mean", path), mean, f"{path}.mean")
    _float_close(
        _required(stats, "standard_deviation", path),
        standard_deviation,
        f"{path}.standard_deviation",
    )
    interval = _require_mapping(_required(stats, "confidence_interval_95", path), f"{path}.confidence_interval_95")
    _exact_keys(interval, CONFIDENCE_INTERVAL_KEYS, f"{path}.confidence_interval_95")
    method = _string(
        _required(interval, "method", f"{path}.confidence_interval_95"),
        f"{path}.confidence_interval_95.method",
    )
    if method != "two-sided Student's t interval for the mean":
        raise ValidationError(f"{path}.confidence_interval_95.method does not match the harness")
    _float_close(
        _required(interval, "lower", f"{path}.confidence_interval_95"),
        max(mean - margin, 0.0),
        f"{path}.confidence_interval_95.lower",
    )
    _float_close(
        _required(interval, "upper", f"{path}.confidence_interval_95"),
        mean + margin,
        f"{path}.confidence_interval_95.upper",
    )
    return samples, order


def _validate_allocation(value: Any, path: str, mode: str) -> str:
    allocation = _require_mapping(value, path)
    status = _string(_required(allocation, "status", path), f"{path}.status")
    scope = _string(_required(allocation, "scope", path), f"{path}.scope")
    if status == "unavailable":
        _exact_keys(allocation, frozenset({"status", "scope"}), path)
    elif status == "measured":
        _exact_keys(allocation, ALLOCATION_KEYS, path)
    else:
        raise ValidationError(f"{path}.status is not a supported allocation status")
    if mode == "native":
        if (
            status != "unavailable"
            or scope != ALLOCATION_SCOPE
            or any(field in allocation for field in ALLOCATION_NUMERIC_FIELDS)
        ):
            raise ValidationError(
                f"{path} must be unavailable with the pinned scope and without numeric allocation evidence"
            )
    else:
        if status != "measured" or scope != ALLOCATION_SCOPE:
            raise ValidationError(f"{path} must be measured system-allocator evidence")
        numeric = {
            field: _u64(_required(allocation, field, path), f"{path}.{field}")
            for field in ALLOCATION_NUMERIC_FIELDS
        }
        if numeric["allocation_calls"] < numeric["reallocation_calls"]:
            raise ValidationError(f"{path}.allocation_calls is smaller than reallocation_calls")
        live_before = numeric["live_bytes_before"]
        live_after = numeric["live_bytes_after"]
        allocated = numeric["allocated_bytes"]
        deallocated = numeric["deallocated_bytes"]
        if allocated > MAX_U64 - live_before or live_after != live_before + allocated - deallocated:
            raise ValidationError(f"{path} live-byte arithmetic is inconsistent")
        peak_before = numeric["peak_live_bytes_before"]
        peak_after = numeric["peak_live_bytes_after"]
        region_peak = numeric["region_peak_live_bytes"]
        if peak_before < live_before or peak_after < peak_before or peak_after < live_after:
            raise ValidationError(f"{path} peak-live bounds are inconsistent")
        if region_peak < max(live_before, live_after) or region_peak > peak_after:
            raise ValidationError(f"{path}.region_peak_live_bytes is outside the live-byte bounds")
    return status


def _validate_raw(raw: Any, path: str, case: str) -> dict[str, Any]:
    value = _require_mapping(raw, path)
    _exact_keys(value, RAW_KEYS, path)
    source_count = _integer(_required(value, "source_member_count", path), f"{path}.source_member_count", 1)
    output_count = _integer(_required(value, "output_member_count", path), f"{path}.output_member_count", 1)
    if source_count != output_count:
        raise ValidationError(f"{path} source/output member counts differ")
    source_digest = _string(_required(value, "source_payload_sha256", path), f"{path}.source_payload_sha256", SHA256)
    output_digest = _string(_required(value, "output_payload_sha256", path), f"{path}.output_payload_sha256", SHA256)
    if (case == CASES[0]) != (source_digest == output_digest):
        raise ValidationError(f"{path} source/output payload digest relationship is wrong for {case}")
    for key in (
        "member_names_equal",
        "untouched_payloads_equal",
        "relationship_names_equal",
        "relationship_payloads_equal",
        "master_theme_payloads_equal",
        "exact_changed_member_set",
        "changed_payloads_nonempty",
    ):
        if not _boolean(_required(value, key, path), f"{path}.{key}"):
            raise ValidationError(f"{path}.{key} is false")
    relationship_count = _integer(_required(value, "relationship_member_count", path), f"{path}.relationship_member_count", 1)
    master_count = _integer(_required(value, "master_theme_member_count", path), f"{path}.master_theme_member_count", 1)
    if relationship_count < 1 or master_count < 1:
        raise ValidationError(f"{path} lacks relationship or master/theme inventory")
    changed = _required(value, "changed_members", path)
    expected = _required(value, "expected_changed_members", path)
    if not isinstance(changed, list) or not isinstance(expected, list):
        raise ValidationError(f"{path}.changed_members and expected_changed_members must be arrays")
    expected_members = {
        CASES[0]: [],
        CASES[1]: ["ppt/notesSlides/notesSlide2.xml"],
        CASES[2]: ["ppt/slides/slide1.xml"],
    }[case]
    if changed != expected_members or expected != expected_members:
        raise ValidationError(f"{path} changed-member evidence does not match {case}")
    if _integer(_required(value, "slide_root_count", path), f"{path}.slide_root_count") != 3:
        raise ValidationError(f"{path}.slide_root_count must be three")
    if _integer(_required(value, "notes_root_count", path), f"{path}.notes_root_count") != 3:
        raise ValidationError(f"{path}.notes_root_count must be three")
    changed_slide = value.get("changed_slide_index")
    changed_notes = value.get("changed_notes_slide_index")
    expected_slide = 0 if case == CASES[2] else None
    expected_notes = 1 if case == CASES[1] else None
    if changed_slide != expected_slide or changed_notes != expected_notes:
        raise ValidationError(f"{path} changed-root identity does not match {case}")
    untouched = _integer(_required(value, "untouched_slide_root_count", path), f"{path}.untouched_slide_root_count")
    if untouched != (2 if case == CASES[2] else 3):
        raise ValidationError(f"{path}.untouched_slide_root_count is not proven by the root inventory")
    return dict(value)


def _validate_rows(report: Mapping[str, Any], path: str, mode: str) -> tuple[int, dict[str, dict[str, Any]]]:
    samples = _integer(_required(report, "samples", path), f"{path}.samples", REQUIRED_SAMPLE_COUNT)
    if samples % SAMPLES_PER_ROUND != 0:
        raise ValidationError(f"{path}.samples must contain complete four-sample rounds")
    warmup = _integer(_required(report, "warmup", path), f"{path}.warmup", 1)
    if _integer(_required(report, "checked_iteration_count", path), f"{path}.checked_iteration_count") != samples + warmup:
        raise ValidationError(f"{path}.checked_iteration_count does not match samples plus warmup")
    rows_value = _required(report, "results", path)
    if not isinstance(rows_value, list) or len(rows_value) != len(CASES):
        raise ValidationError(f"{path}.results must contain exactly the three probe cases")
    rows: dict[str, dict[str, Any]] = {}
    for index, row_value in enumerate(rows_value):
        row_path = f"{path}.results[{index}]"
        row = _require_mapping(row_value, row_path)
        _exact_keys(row, ROW_KEYS, row_path)
        case = _string(_required(row, "case", row_path), f"{row_path}.case")
        if case not in CASES or case in rows:
            raise ValidationError(f"{row_path}.case is not a unique pinned probe case")
        if _required(row, "operation", row_path) != OPERATIONS[case]:
            raise ValidationError(f"{row_path}.operation does not match {case}")
        timing_scope = _string(_required(row, "timing_scope", row_path), f"{row_path}.timing_scope")
        if timing_scope != TIMING_SCOPES[case]:
            raise ValidationError(f"{row_path}.timing_scope is not the pinned operation scope")
        elapsed, elapsed_order = _validate_statistics(
            _required(row, "elapsed_ns", row_path), f"{row_path}.elapsed_ns", samples
        )
        observations_value = _required(row, "observations", row_path)
        if not isinstance(observations_value, list) or len(observations_value) != samples:
            raise ValidationError(f"{row_path}.observations must contain every retained sample")
        output_identity: tuple[int, str] | None = None
        fidelity_identity: str | None = None
        observation_elapsed: list[int] = []
        for sample_index, observation_value in enumerate(observations_value):
            observation_path = f"{row_path}.observations[{sample_index}]"
            observation = _require_mapping(observation_value, observation_path)
            _exact_keys(observation, OBSERVATION_KEYS, observation_path)
            if _integer(_required(observation, "sample_index", observation_path), f"{observation_path}.sample_index") != sample_index:
                raise ValidationError(f"{observation_path}.sample_index is not the retained order")
            observed_elapsed = _integer(
                _required(observation, "elapsed_ns", observation_path),
                f"{observation_path}.elapsed_ns",
                1,
            )
            observation_elapsed.append(observed_elapsed)
            output_bytes = _integer(_required(observation, "output_bytes", observation_path), f"{observation_path}.output_bytes", 1)
            output_sha = _string(_required(observation, "output_sha256", observation_path), f"{observation_path}.output_sha256", SHA256)
            identity = (output_bytes, output_sha)
            if output_identity is None:
                output_identity = identity
            elif output_identity != identity:
                raise ValidationError(f"{row_path} output identity is not deterministic across samples")
            fidelity = _require_mapping(_required(observation, "fidelity", observation_path), f"{observation_path}.fidelity")
            _exact_keys(
                fidelity,
                FIDELITY_KEYS if case == CASES[0] else FIDELITY_EDIT_KEYS,
                f"{observation_path}.fidelity",
            )
            semantic = _require_mapping(_required(fidelity, "semantic", observation_path), f"{observation_path}.fidelity.semantic")
            _exact_keys(semantic, SEMANTIC_KEYS, f"{observation_path}.fidelity.semantic")
            if not _boolean(_required(semantic, "verified", f"{observation_path}.fidelity.semantic"), f"{observation_path}.fidelity.semantic.verified"):
                raise ValidationError(f"{observation_path} semantic fidelity is false")
            if _integer(_required(semantic, "slide_count", f"{observation_path}.fidelity.semantic"), f"{observation_path}.fidelity.semantic.slide_count") != 3 or _integer(_required(semantic, "notes_slide_count", f"{observation_path}.fidelity.semantic"), f"{observation_path}.fidelity.semantic.notes_slide_count") != 3:
                raise ValidationError(f"{observation_path} semantic slide/notes counts are wrong")
            _string(_required(semantic, "slide_text_sha256", f"{observation_path}.fidelity.semantic"), f"{observation_path}.fidelity.semantic.slide_text_sha256", SHA256)
            _string(_required(semantic, "notes_text_sha256", f"{observation_path}.fidelity.semantic"), f"{observation_path}.fidelity.semantic.notes_text_sha256", SHA256)
            if case == CASES[0]:
                if "source_immutable" in fidelity:
                    raise ValidationError(f"{observation_path} cold fidelity must not claim an edit source check")
                changed_reported = False
            else:
                if fidelity.get("source_immutable") is not True:
                    raise ValidationError(f"{observation_path} edit source was not proven immutable")
                changed_reported = True
            if _boolean(_required(fidelity, "changed_commit_reported", f"{observation_path}.fidelity"), f"{observation_path}.fidelity.changed_commit_reported") != changed_reported:
                raise ValidationError(f"{observation_path} changed-commit diagnostic does not match the operation")
            for key in ("output_reopened",):
                if not _boolean(_required(fidelity, key, f"{observation_path}.fidelity"), f"{observation_path}.fidelity.{key}"):
                    raise ValidationError(f"{observation_path}.fidelity.{key} is false")
            expected_changed_slide = 0 if case == CASES[2] else None
            expected_untouched_slides = 2 if case == CASES[2] else 3
            if fidelity.get("changed_slide_index") != expected_changed_slide:
                raise ValidationError(f"{observation_path}.fidelity.changed_slide_index is not proven")
            if _integer(_required(fidelity, "untouched_slide_root_count", f"{observation_path}.fidelity"), f"{observation_path}.fidelity.untouched_slide_root_count") != expected_untouched_slides:
                raise ValidationError(f"{observation_path}.fidelity.untouched_slide_root_count is not proven")
            _validate_raw(_required(fidelity, "raw", f"{observation_path}.fidelity"), f"{observation_path}.fidelity.raw", case)
            current_fidelity_identity = json.dumps(fidelity, sort_keys=True, separators=(",", ":"))
            if fidelity_identity is None:
                fidelity_identity = current_fidelity_identity
            elif fidelity_identity != current_fidelity_identity:
                raise ValidationError(f"{row_path} fidelity oracle evidence is not deterministic across samples")
            _validate_allocation(_required(observation, "allocation", observation_path), f"{observation_path}.allocation", mode)
        ordered_observations = sorted(enumerate(observation_elapsed), key=lambda item: (item[1], item[0]))
        if [value for _index, value in ordered_observations] != elapsed or [index for index, _value in ordered_observations] != elapsed_order:
            raise ValidationError(f"{row_path}.observations are not aligned with elapsed_ns statistics")
        rows[case] = dict(row)
    if set(rows) != set(CASES):
        raise ValidationError(f"{path}.results does not cover every probe case")
    return samples, rows


def validate_report(report: Mapping[str, Any], label: str, mode: str = "native") -> dict[str, Any]:
    path = label
    if mode not in ("native", "allocator"):
        raise ValidationError(f"{path}: mode must be native or allocator")
    _exact_keys(report, REPORT_KEYS, path)
    if _required(report, "schema", path) != SCHEMA:
        raise ValidationError(f"{path}.schema is not {SCHEMA}")
    if _required(report, "tool", path) != "litchi-perf-baseline":
        raise ValidationError(f"{path}.tool is not litchi-perf-baseline")
    revision = _string(_required(report, "source_revision", path), f"{path}.source_revision", REVISION)
    instrumentation = _string(_required(report, "instrumentation", path), f"{path}.instrumentation")
    if mode == "native":
        if instrumentation != "none":
            raise ValidationError(f"{path}.instrumentation must be none for native ABBA")
    elif instrumentation != "system_allocator_operation_scoped":
        raise ValidationError(f"{path}.instrumentation must be system_allocator_operation_scoped for allocator ABBA")
    allocator = _string(_required(report, "allocator", path), f"{path}.allocator")
    expected_allocator = NATIVE_ALLOCATOR if mode == "native" else INSTRUMENTED_ALLOCATOR
    if allocator != expected_allocator:
        raise ValidationError(f"{path}.allocator does not match the instrumentation mode")
    counter_revision = _required(report, "allocator_counter_revision", path)
    if mode == "allocator":
        if counter_revision != ALLOCATOR_COUNTER_REVISION:
            raise ValidationError(f"{path}.allocator_counter_revision is not the pinned counter revision")
    elif counter_revision is not None:
        raise ValidationError(f"{path}.allocator_counter_revision must be null for native ABBA")
    expected_timing_claim = (
        "native_wall_clock_samples; compare only under the approved quiet ABBA protocol"
        if mode == "native"
        else "allocator-instrumented wall-clock samples; not a latency claim"
    )
    if _required(report, "timing_claim", path) != expected_timing_claim:
        raise ValidationError(f"{path}.timing_claim does not match the instrumentation mode")
    if _required(report, "allocation_scope", path) != REPORT_ALLOCATION_SCOPE:
        raise ValidationError(f"{path}.allocation_scope is not the pinned operation scope")
    if _required(report, "memo_scope", path) != MEMO_SCOPE:
        raise ValidationError(f"{path}.memo_scope is not the pinned memo scope")
    if _required(report, "fidelity_scope", path) != FIDELITY_SCOPE:
        raise ValidationError(f"{path}.fidelity_scope is not the pinned fidelity scope")
    binary = _validate_binary(report, path)
    fixture = _validate_fixture(report, path)
    samples, rows = _validate_rows(report, path, mode)
    return {
        "label": label,
        "revision": revision,
        "instrumentation": instrumentation,
        "fixture": fixture,
        "samples": samples,
        "warmup": _integer(_required(report, "warmup", path), f"{path}.warmup"),
        "binary": binary,
        "rows": rows,
    }


def validate_abba(leg_reports: Mapping[str, Mapping[str, Any]], mode: str = "native") -> dict[str, Any]:
    if mode not in ("native", "allocator"):
        raise ValidationError("mode must be native or allocator")
    if tuple(leg_reports) != LEG_ORDER:
        raise ValidationError(f"leg order must be exactly {LEG_ORDER}")
    validated = [validate_report(leg_reports[leg], leg, mode) for leg in LEG_ORDER]
    first = validated[0]
    for item in validated[1:]:
        if item["fixture"] != first["fixture"]:
            raise ValidationError("all ABBA legs must use the exact same fixture identity")
        if item["samples"] != first["samples"] or item["warmup"] != first["warmup"]:
            raise ValidationError("all ABBA legs must use the same sample and warmup counts")
        if item["instrumentation"] != first["instrumentation"]:
            raise ValidationError("all ABBA legs must use the same instrumentation mode")
    if validated[0]["revision"] != validated[3]["revision"]:
        raise ValidationError("A1 and A2 source revisions differ")
    if validated[1]["revision"] != validated[2]["revision"]:
        raise ValidationError("B1 and B2 source revisions differ")
    if validated[0]["revision"] == validated[1]["revision"]:
        raise ValidationError("A and B source revisions must differ")
    for first_index, second_index, label in ((0, 3, "A"), (1, 2, "B")):
        first_binary = validated[first_index]["binary"]
        second_binary = validated[second_index]["binary"]
        identity_keys = ("path", "binary_sha256", "binary_bytes", "mode_bits", "executable", "profile")
        if any(first_binary[key] != second_binary[key] for key in identity_keys):
            raise ValidationError(f"{label} ABBA legs do not use the same binary identity")
    if validated[0]["binary"]["binary_sha256"] == validated[1]["binary"]["binary_sha256"]:
        raise ValidationError("A and B ABBA legs must use distinct binary identities")
    if first["samples"] < REQUIRED_SAMPLE_COUNT:
        raise ValidationError("ABBA requires at least four samples per round and four rounds")
    source_inventory_baseline: dict[str, Any] | None = None
    rows = []
    for case in CASES:
        cross_report_output_identity: tuple[int, str] | None = None
        cross_report_fidelity: str | None = None
        cross_report_source_inventory: dict[str, Any] | None = None
        first_fidelity: dict[str, Any] | None = None
        for item in validated:
            observation = item["rows"][case]["observations"][0]
            output_identity = (observation["output_bytes"], observation["output_sha256"])
            fidelity = observation["fidelity"]
            raw = fidelity["raw"]
            source_inventory = {
                key: raw[key]
                for key in (
                    "source_member_count",
                    "output_member_count",
                    "source_payload_sha256",
                    "relationship_member_count",
                    "master_theme_member_count",
                    "slide_root_count",
                    "notes_root_count",
                )
            }
            if cross_report_output_identity is None:
                cross_report_output_identity = output_identity
                cross_report_fidelity = json.dumps(fidelity, sort_keys=True, separators=(",", ":"))
                cross_report_source_inventory = source_inventory
                first_fidelity = fidelity
            else:
                if output_identity != cross_report_output_identity:
                    raise ValidationError(f"{case} output bytes/digest differ across ABBA reports")
                if json.dumps(fidelity, sort_keys=True, separators=(",", ":")) != cross_report_fidelity:
                    raise ValidationError(f"{case} fidelity evidence differs across ABBA reports")
                if source_inventory != cross_report_source_inventory:
                    raise ValidationError(f"{case} source inventory differs across ABBA reports")
            if source_inventory_baseline is None:
                source_inventory_baseline = source_inventory
            elif source_inventory != source_inventory_baseline:
                raise ValidationError("source inventory differs between probe rows")
        rows.append(
            {
                "case": case,
                "operation": OPERATIONS[case],
                "elapsed_ns_samples": {
                    leg: validated[index]["rows"][case]["elapsed_ns"]["samples"]
                    for index, leg in enumerate(LEG_ORDER)
                },
                "output_identity": {
                    "bytes": cross_report_output_identity[0],
                    "sha256": cross_report_output_identity[1],
                },
                "source_inventory": cross_report_source_inventory,
                "fidelity_evidence": first_fidelity,
                "fidelity_gate": "all observations passed semantic reopen, raw payload, relationship, master/theme, and changed-root gates",
                "allocation_metric_gate": mode,
            }
        )
    return {
        "schema": SUMMARY_SCHEMA,
        "tool": "litchi-perf-baseline-pptx-notes-memo-abba",
        "leg_order": list(LEG_ORDER),
        "samples_per_round": SAMPLES_PER_ROUND,
        "rounds": first["samples"] // SAMPLES_PER_ROUND,
        "samples_per_leg": first["samples"],
        "warmup": first["warmup"],
        "instrumentation": first["instrumentation"],
        "fixture": first["fixture"],
        "source_revisions": {leg: validated[index]["revision"] for index, leg in enumerate(LEG_ORDER)},
        "binary_identities": {leg: validated[index]["binary"] for index, leg in enumerate(LEG_ORDER)},
        "rows": rows,
        "performance_claim": "descriptive four-leg ABBA evidence only; no speedup is inferred",
    }


def _load(path: Path) -> Mapping[str, Any]:
    try:
        text = path.read_text(encoding="utf-8")
        value = load_strict_json(text, str(path))
    except (OSError, UnicodeError, ValidationError) as error:
        raise ValidationError(f"cannot read {path}: {error}") from error
    return _require_mapping(value, str(path))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mode", choices=("native", "allocator"), default="native")
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument(
        "--descriptive-only",
        action="store_true",
        help="explicitly mark this direct four-report check as descriptive-only; use the process runner for acceptance",
    )
    parser.add_argument("a1", type=Path)
    parser.add_argument("b1", type=Path)
    parser.add_argument("b2", type=Path)
    parser.add_argument("a2", type=Path)
    arguments = parser.parse_args(argv)
    if not arguments.descriptive_only:
        parser.error(
            "direct four-report validation is descriptive-only; pass --descriptive-only or use run_pptx_notes_memo_abba.py"
        )
    try:
        reports = {
            "a1": _load(arguments.a1),
            "b1": _load(arguments.b1),
            "b2": _load(arguments.b2),
            "a2": _load(arguments.a2),
        }
        summary = validate_abba(reports, arguments.mode)
        with arguments.output.open("x", encoding="utf-8") as handle:
            json.dump(summary, handle, indent=2, sort_keys=True)
            handle.write("\n")
    except (OSError, ValidationError) as error:
        print(f"pptx notes memo ABBA validation failed: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
