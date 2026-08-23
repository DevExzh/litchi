#!/usr/bin/env python3
"""Fail-closed comparison of two litchi performance JSON reports.

The comparator intentionally uses only the Python standard library.  A policy
pins the accepted report shape, identity fields, sample floor, percentile
validation, metric presence, and regression thresholds.  Exit status is 0 for
a passing comparison, 1 for measured regressions, and 2 for invalid or
incomparable input. Reports carrying the benchmark-only allocator
instrumentation identity retain elapsed validation but withhold latency
comparisons; optional operation metrics remain comparable.
"""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import math
import statistics
import sys
from pathlib import Path, PureWindowsPath
from typing import Any, Iterable


COMPARATOR_NAME = "litchi-perf-compare"
COMPARATOR_VERSION = "1.3.2"
SUPPORTED_POLICY_SCHEMA = 2
SUPPORTED_REPORT_SCHEMA = 1
EVIDENCE_ONLY_LATENCY_CLAIM = "evidence_only_filesystem_selector"
COMPARABLE_LATENCY_CLAIM = "comparable_timed_operation"
OPERATION_ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"

_DEFAULT_RESULT_KEY_FIELDS = ("case", "corpus")
_CACHE_RESULT_KEY_FIELDS = ("case", "corpus", "cache_state")
_METADATA_FIELD_NAMES = frozenset(
    {
        "status",
        "scope",
        "alignment",
        "sample_count",
        "counter_scope",
        "unit",
    }
)
_ALLOCATOR_INSTRUMENTATION = "system_allocator_operation_scoped"
_ALLOCATOR_BINARY = "litchi-perf-baseline-alloc"
_ALLOCATOR_SCOPE = "operation_global_system_allocator"
_ALLOCATOR_VECTOR_FIELDS = (
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
)
_FILESYSTEM_CASE_PREFIXES = (
    "opc_file_",
    "cfb_file_",
    "pptx_file_",
    "docx_file_",
    "odt_file_",
    "ods_file_",
    "odp_file_",
)
_OPTIONAL_POLICY_KEYS = {
    "result_key_fields",
    "expected_build_identity",
    "require_binary_identity",
}
_BINARY_IDENTITY_FIELDS = frozenset(
    {"path", "binary_sha256", "binary_bytes", "mode_bits", "executable", "profile"}
)
_BINARY_IDENTITY_OPTIONAL_FIELDS = frozenset({"label"})
_UNIX_TARGET_OSES = frozenset(
    {
        "aix", "android", "dragonfly", "freebsd", "fuchsia", "haiku",
        "hurd", "illumos", "ios", "linux", "macos", "netbsd", "openbsd",
        "tvos", "visionos", "watchos",
    }
)


class ComparisonInputError(ValueError):
    """Raised when inputs cannot be compared safely."""


def _reject_nonstandard_constant(value: str) -> None:
    raise ComparisonInputError(f"non-finite JSON number {value!r}")


def load_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as handle:
            return json.load(handle, parse_constant=_reject_nonstandard_constant)
    except (OSError, json.JSONDecodeError) as error:
        raise ComparisonInputError(f"cannot read {path}: {error}") from error


def _require_object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ComparisonInputError(f"{location} must be an object")
    return value


def _validate_binary_identity(
    value: Any, location: str, *, tool: dict[str, Any]
) -> dict[str, Any]:
    """Validate a report's untimed current-executable descriptor.

    The descriptor is intentionally validated independently for baseline and
    current reports.  A comparison is allowed to use different binaries; the
    descriptor is provenance/syntax evidence, not a field that must match
    across those two reports.
    """

    identity = _require_object(value, location)
    missing = _BINARY_IDENTITY_FIELDS - set(identity)
    if missing:
        raise ComparisonInputError(
            f"{location} is missing required keys: {sorted(missing)}"
        )
    unknown = set(identity) - _BINARY_IDENTITY_FIELDS - _BINARY_IDENTITY_OPTIONAL_FIELDS
    if unknown:
        raise ComparisonInputError(f"{location} has unknown keys: {sorted(unknown)}")
    path = identity["path"]
    if not isinstance(path, str) or not path:
        raise ComparisonInputError(f"{location}.path must be a non-empty string")
    if not (Path(path).is_absolute() or PureWindowsPath(path).is_absolute()):
        raise ComparisonInputError(f"{location}.path must be absolute")
    digest = identity["binary_sha256"]
    if (
        not isinstance(digest, str)
        or len(digest) != 64
        or any(character not in "0123456789abcdefABCDEF" for character in digest)
    ):
        raise ComparisonInputError(
            f"{location}.binary_sha256 must contain exactly 64 hexadecimal characters"
        )
    binary_bytes = identity["binary_bytes"]
    if (
        isinstance(binary_bytes, bool)
        or not isinstance(binary_bytes, int)
        or binary_bytes <= 0
        or binary_bytes > (1 << 64) - 1
    ):
        raise ComparisonInputError(
            f"{location}.binary_bytes must be a positive unsigned integer"
        )
    mode_bits = identity["mode_bits"]
    if mode_bits is not None and (
        isinstance(mode_bits, bool)
        or not isinstance(mode_bits, int)
        or mode_bits < 0
        or mode_bits > 0o7777
        or mode_bits & 0o111 == 0
    ):
        raise ComparisonInputError(
            f"{location}.mode_bits must be null or executable permission bits"
        )
    if tool["target_os"] in _UNIX_TARGET_OSES and mode_bits is None:
        raise ComparisonInputError(
            f"{location}.mode_bits must be present for Unix targets"
        )
    if identity["executable"] is not True:
        raise ComparisonInputError(f"{location}.executable must be true")
    profile = identity["profile"]
    if not isinstance(profile, str) or not profile:
        raise ComparisonInputError(f"{location}.profile must be a non-empty string")
    if profile != tool["profile"]:
        raise ComparisonInputError(
            f"{location}.profile does not match report tool.profile"
        )
    if "label" in identity and (
        not isinstance(identity["label"], str) or not identity["label"]
    ):
        raise ComparisonInputError(f"{location}.label must be a non-empty string")
    return {
        **identity,
        "binary_sha256": digest.lower(),
    }


def _require_schema_version(value: Any, location: str, expected: int) -> None:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ComparisonInputError(f"{location} must be an integer")
    if value != expected:
        raise ComparisonInputError(
            f"unsupported {location} {value!r}; expected {expected}"
        )


def _finite_number(value: Any, location: str, *, positive: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ComparisonInputError(f"{location} must be a finite number")
    try:
        number = float(value)
    except OverflowError as error:
        raise ComparisonInputError(f"{location} is outside the finite range") from error
    if not math.isfinite(number):
        raise ComparisonInputError(f"{location} must be a finite number")
    if positive and number <= 0:
        raise ComparisonInputError(f"{location} must be greater than zero")
    if not positive and number < 0:
        raise ComparisonInputError(f"{location} must be non-negative")
    return number


def _reject_nonfinite_tree(value: Any, location: str) -> None:
    if isinstance(value, bool) or value is None or isinstance(value, str):
        return
    if isinstance(value, (int, float)):
        try:
            finite = math.isfinite(float(value))
        except OverflowError as error:
            raise ComparisonInputError(
                f"{location} contains a number outside the finite range"
            ) from error
        if not finite:
            raise ComparisonInputError(f"{location} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _reject_nonfinite_tree(item, f"{location}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            _reject_nonfinite_tree(item, f"{location}.{key}")


def validate_policy(raw: Any) -> dict[str, Any]:
    policy = _require_object(raw, "policy")
    required = {
        "schema_version",
        "policy_id",
        "minimum_samples",
        "expected_result_count",
        "expected_result_keys_sha256",
        "required_cases",
        "require_clean_worktree",
        "require_distinct_revisions",
        "tool_identity",
        "build_identity_fields",
        "nullable_build_identity_fields",
        "expected_configuration",
        "latency_thresholds_percent",
        "metric_classes",
    }
    missing = sorted(required - policy.keys())
    unknown = sorted(policy.keys() - required - _OPTIONAL_POLICY_KEYS)
    if missing or unknown:
        raise ComparisonInputError(
            f"policy keys mismatch: missing={missing}, unknown={unknown}"
        )
    _require_schema_version(
        policy["schema_version"], "policy.schema_version", SUPPORTED_POLICY_SCHEMA
    )
    if not isinstance(policy["policy_id"], str) or not policy["policy_id"]:
        raise ComparisonInputError("policy.policy_id must be a non-empty string")
    minimum_samples = policy["minimum_samples"]
    if isinstance(minimum_samples, bool) or not isinstance(minimum_samples, int):
        raise ComparisonInputError("policy.minimum_samples must be an integer")
    if minimum_samples < 3:
        raise ComparisonInputError("policy.minimum_samples must be at least 3")
    expected_count = policy["expected_result_count"]
    if isinstance(expected_count, bool) or not isinstance(expected_count, int):
        raise ComparisonInputError("policy.expected_result_count must be an integer")
    if expected_count < 1:
        raise ComparisonInputError("policy.expected_result_count must be positive")
    expected_keys_sha256 = policy["expected_result_keys_sha256"]
    if expected_keys_sha256 is not None and (
        not isinstance(expected_keys_sha256, str)
        or len(expected_keys_sha256) != 64
        or any(character not in "0123456789abcdef" for character in expected_keys_sha256)
    ):
        raise ComparisonInputError(
            "policy.expected_result_keys_sha256 must be null or 64 lowercase hex digits"
        )
    required_cases = policy["required_cases"]
    if (
        not isinstance(required_cases, list)
        or not required_cases
        or any(not isinstance(case, str) or not case for case in required_cases)
        or len(set(required_cases)) != len(required_cases)
    ):
        raise ComparisonInputError(
            "policy.required_cases must contain unique non-empty strings"
        )
    if not isinstance(policy["require_clean_worktree"], bool):
        raise ComparisonInputError("policy.require_clean_worktree must be boolean")
    if not isinstance(policy["require_distinct_revisions"], bool):
        raise ComparisonInputError("policy.require_distinct_revisions must be boolean")
    if "require_binary_identity" in policy and not isinstance(
        policy["require_binary_identity"], bool
    ):
        raise ComparisonInputError("policy.require_binary_identity must be boolean")
    tool = _require_object(policy["tool_identity"], "policy.tool_identity")
    for field in (
        "name",
        "version",
        "binary",
        "profile",
        "target_os",
        "target_arch",
        "instrumentation",
    ):
        if not isinstance(tool.get(field), str) or not tool[field]:
            raise ComparisonInputError(
                f"policy.tool_identity.{field} must be a non-empty string"
            )
    result_key_fields = policy.get("result_key_fields", list(_DEFAULT_RESULT_KEY_FIELDS))
    if result_key_fields not in (
        list(_DEFAULT_RESULT_KEY_FIELDS),
        list(_CACHE_RESULT_KEY_FIELDS),
    ):
        raise ComparisonInputError(
            "policy.result_key_fields must be ['case', 'corpus'] or "
            "['case', 'corpus', 'cache_state']"
        )
    if tool["instrumentation"] == _ALLOCATOR_INSTRUMENTATION:
        if result_key_fields != list(_CACHE_RESULT_KEY_FIELDS):
            raise ComparisonInputError(
                "allocator policy must include cache_state in result_key_fields"
            )
        if tool["binary"] != _ALLOCATOR_BINARY:
            raise ComparisonInputError(
                "allocator policy must pin tool_identity.binary to "
                f"{_ALLOCATOR_BINARY!r}"
            )
    elif result_key_fields != list(_DEFAULT_RESULT_KEY_FIELDS):
        raise ComparisonInputError(
            "cache_state result keys are reserved for allocator policy"
        )
    identity_fields = policy["build_identity_fields"]
    if (
        not isinstance(identity_fields, list)
        or not identity_fields
        or any(not isinstance(field, str) or not field for field in identity_fields)
        or len(set(identity_fields)) != len(identity_fields)
    ):
        raise ComparisonInputError(
            "policy.build_identity_fields must contain unique non-empty strings"
        )
    nullable_identity_fields = policy["nullable_build_identity_fields"]
    if (
        not isinstance(nullable_identity_fields, list)
        or any(
            not isinstance(field, str) or not field
            for field in nullable_identity_fields
        )
        or len(set(nullable_identity_fields)) != len(nullable_identity_fields)
        or not set(nullable_identity_fields) <= set(identity_fields)
    ):
        raise ComparisonInputError(
            "policy.nullable_build_identity_fields must be a unique subset of "
            "build_identity_fields"
        )
    expected_build_identity = policy.get("expected_build_identity")
    if expected_build_identity is not None:
        expected_identity = _require_object(
            expected_build_identity, "policy.expected_build_identity"
        )
        if not expected_identity:
            raise ComparisonInputError(
                "policy.expected_build_identity must not be empty"
            )
        unknown_identity = sorted(set(expected_identity) - set(identity_fields))
        if unknown_identity:
            raise ComparisonInputError(
                "policy.expected_build_identity contains fields outside "
                f"build_identity_fields: {unknown_identity}"
            )
        for field, value in expected_identity.items():
            if value is None or isinstance(value, bool) or not isinstance(value, (int, str)):
                raise ComparisonInputError(
                    f"policy.expected_build_identity.{field} must be a non-null scalar"
                )
            if isinstance(value, str) and not value:
                raise ComparisonInputError(
                    f"policy.expected_build_identity.{field} must be non-empty"
                )
            if isinstance(value, int) and value <= 0:
                raise ComparisonInputError(
                    f"policy.expected_build_identity.{field} must be positive"
                )
    expected_configuration = _require_object(
        policy["expected_configuration"], "policy.expected_configuration"
    )
    if not expected_configuration:
        raise ComparisonInputError("policy.expected_configuration must not be empty")
    _reject_nonfinite_tree(expected_configuration, "policy.expected_configuration")
    latency = _require_object(
        policy["latency_thresholds_percent"],
        "policy.latency_thresholds_percent",
    )
    if set(latency) != {"p50", "p95", "p99"}:
        raise ComparisonInputError(
            "policy.latency_thresholds_percent must contain exactly p50, p95, and p99"
        )
    for name, value in latency.items():
        _finite_number(value, f"policy.latency_thresholds_percent.{name}")
    classes = policy["metric_classes"]
    if not isinstance(classes, list) or not classes:
        raise ComparisonInputError("policy.metric_classes must be a non-empty list")
    class_names: set[str] = set()
    for index, item in enumerate(classes):
        metric_class = _require_object(item, f"policy.metric_classes[{index}]")
        required_class_keys = {
            "name",
            "max_regression_percent",
            "path_globs",
        }
        allowed_class_keys = required_class_keys | {"presence"}
        class_keys = set(metric_class)
        if (
            not class_keys <= allowed_class_keys
            or not required_class_keys <= class_keys
        ):
            raise ComparisonInputError(
                f"policy.metric_classes[{index}] has invalid keys"
            )
        if "presence" not in metric_class:
            raise ComparisonInputError(
                f"policy.metric_classes[{index}].presence is required by policy schema "
                f"{SUPPORTED_POLICY_SCHEMA}"
            )
        name = metric_class["name"]
        if not isinstance(name, str) or not name or name in class_names:
            raise ComparisonInputError("metric class names must be unique strings")
        class_names.add(name)
        _finite_number(
            metric_class["max_regression_percent"],
            f"policy.metric_classes[{index}].max_regression_percent",
        )
        globs = metric_class["path_globs"]
        if (
            not isinstance(globs, list)
            or not globs
            or any(not isinstance(pattern, str) or not pattern for pattern in globs)
        ):
            raise ComparisonInputError(
                f"policy.metric_classes[{index}].path_globs must be non-empty strings"
            )
        presence = metric_class["presence"]
        if not isinstance(presence, str) or presence not in {"required", "optional"}:
            raise ComparisonInputError(
                f"policy.metric_classes[{index}].presence must be 'required' or 'optional'"
            )
    return policy


def _result_key(
    result: dict[str, Any], location: str, *, include_cache_state: bool = False
) -> tuple[str, str, str]:
    case = result.get("case")
    corpus = result.get("corpus")
    if not isinstance(case, str) or not case:
        raise ComparisonInputError(f"{location}.case must be a non-empty string")
    if not isinstance(corpus, dict) or not corpus:
        raise ComparisonInputError(f"{location}.corpus must be a non-empty object")
    try:
        corpus_identity = json.dumps(
            corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
        )
    except (TypeError, ValueError) as error:
        raise ComparisonInputError(f"{location}.corpus is not canonical JSON: {error}")
    has_cache_state = "cache_state" in result
    cache_state = result.get("cache_state", "")
    if not isinstance(cache_state, str) or (has_cache_state and not cache_state):
        raise ComparisonInputError(
            f"{location}.cache_state must be a non-empty string when present"
        )
    if cache_state and cache_state not in {"warm", "cold-requested"}:
        raise ComparisonInputError(
            f"{location}.cache_state must be 'warm' or 'cold-requested'"
        )
    if include_cache_state and not cache_state:
        raise ComparisonInputError(
            f"{location}.cache_state must be a non-empty string for allocator evidence"
        )
    return case, corpus_identity, cache_state


def _index_results(
    report: dict[str, Any],
    label: str,
    expected_count: int,
    *,
    include_cache_state: bool = False,
) -> dict[tuple[str, str, str], dict[str, Any]]:
    results = report.get("results")
    if not isinstance(results, list):
        raise ComparisonInputError(f"{label}.results must be a list")
    if len(results) != expected_count:
        raise ComparisonInputError(
            f"{label}.results has {len(results)} entries; expected {expected_count}"
        )
    indexed: dict[tuple[str, str, str], dict[str, Any]] = {}
    for index, raw_result in enumerate(results):
        result = _require_object(raw_result, f"{label}.results[{index}]")
        if "elapsed_ns" not in result:
            raise ComparisonInputError(f"{label}.results[{index}] lacks elapsed_ns")
        key = _result_key(
            result,
            f"{label}.results[{index}]",
            include_cache_state=include_cache_state,
        )
        if key in indexed:
            raise ComparisonInputError(
                f"{label} contains duplicate case/corpus key for {key[0]!r}"
            )
        indexed[key] = result
    return indexed


def result_key_manifest_sha256(
    keys: Iterable[tuple[str, str] | tuple[str, str, str]]
) -> str:
    digest = hashlib.sha256()
    for key in sorted(keys):
        if len(key) not in {2, 3}:
            raise ValueError("result keys must contain case/corpus and optional cache_state")
        if len(key) == 2:
            case, corpus_identity = key
            cache_state = ""
        elif len(key) == 3:
            case, corpus_identity, cache_state = key
        else:
            raise ValueError("result keys must contain case/corpus and optional cache_state")
        digest.update(case.encode("utf-8"))
        digest.update(b"\0")
        digest.update(corpus_identity.encode("utf-8"))
        if cache_state:
            digest.update(b"\0")
            digest.update(cache_state.encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()


def report_result_key_manifest_sha256(
    report: Any,
    expected_count: int,
    *,
    result_key_fields: Iterable[str] = _DEFAULT_RESULT_KEY_FIELDS,
) -> str:
    """Returns the report key digest for the selected public key fields.

    Normal reports use ``(case, corpus)``. Filesystem allocator reports must
    pass ``("case", "corpus", "cache_state")`` so warm and cold rows remain
    distinct even when they share one corpus.
    """
    report_object = _require_object(report, "report")
    fields = tuple(result_key_fields)
    if fields not in (_DEFAULT_RESULT_KEY_FIELDS, _CACHE_RESULT_KEY_FIELDS):
        raise ComparisonInputError(
            "result_key_fields must be ('case', 'corpus') or "
            "('case', 'corpus', 'cache_state')"
        )
    indexed = _index_results(
        report_object,
        "report",
        expected_count,
        include_cache_state=fields == _CACHE_RESULT_KEY_FIELDS,
    )
    return result_key_manifest_sha256(indexed)


_PARALLEL_METRICS_SCHEMA = 1
_PARALLEL_METRICS_CLAIM = "descriptive"
_PARALLEL_METRICS_SCOPE = "explicit_local_execution_only"
_PARALLEL_METRIC_STATUSES = {"measured", "not_applicable", "unavailable"}
_PARALLEL_U64_MAX = (1 << 64) - 1
_PARALLEL_CHUNK_SCOPE = "result.execution.deterministic_chunk_count"
_PARALLEL_SIMULATION_CHUNK_SCOPE = "result.source.simulation.physical_request_count"
_PARALLEL_CFB_SELECTIVE_CHUNK_SCOPE = (
    "result.source.cfb_selective.simulation.read.physical_request_count"
)
_PARALLEL_CFB_OPEN_CHUNK_SCOPE = (
    "result.source.cfb_open_stream.simulation.samples.per_operation."
    "physical_request_count_sum"
)


def _parallel_exact_keys(
    value: Any,
    path: str,
    required: set[str],
    optional: set[str] | None = None,
) -> dict[str, Any]:
    obj = _require_object(value, path)
    actual = set(obj)
    allowed = required | (optional or set())
    missing = sorted(required - actual)
    unknown = sorted(actual - allowed)
    if missing or unknown:
        raise ComparisonInputError(
            f"{path} keys mismatch: missing={missing}, unknown={unknown}"
        )
    return obj


def _parallel_scope_matches(scope: Any, expected: str | set[str], path: str) -> None:
    if not isinstance(scope, str) or not scope:
        raise ComparisonInputError(f"{path} must be a non-empty string")
    expected_values = {expected} if isinstance(expected, str) else expected
    if scope not in expected_values:
        raise ComparisonInputError(
            f"{path}={scope!r} is outside the accepted scopes {sorted(expected_values)!r}"
        )


def _parallel_numeric_value(value: Any, path: str) -> None:
    if isinstance(value, bool):
        raise ComparisonInputError(f"{path} must be a non-negative integer")
    if isinstance(value, int):
        if value < 0 or value > _PARALLEL_U64_MAX:
            raise ComparisonInputError(
                f"{path} must be an unsigned 64-bit integer"
            )
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _parallel_numeric_value(item, f"{path}[{index}]")
        return
    raise ComparisonInputError(f"{path} must be a non-negative integer or vector")


def _parallel_scalar_value(value: Any, path: str) -> int:
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or value < 0
        or value > _PARALLEL_U64_MAX
    ):
        raise ComparisonInputError(
            f"{path} must be a non-negative unsigned 64-bit integer scalar"
        )
    return value


def _parallel_u64_vector(value: Any, path: str) -> list[int]:
    if not isinstance(value, list):
        raise ComparisonInputError(f"{path} must be an unsigned 64-bit integer vector")
    return [
        _parallel_scalar_value(item, f"{path}[{index}]")
        for index, item in enumerate(value)
    ]


def _parallel_worker_vector(value: Any, path: str) -> list[int]:
    if (
        not isinstance(value, list)
        or not value
        or any(
            isinstance(item, bool)
            or not isinstance(item, int)
            or item <= 0
            or item > _PARALLEL_U64_MAX
            for item in value
        )
        or value != sorted(set(value))
    ):
        raise ComparisonInputError(
            f"{path} must be a sorted, unique, positive worker vector"
        )
    return value


def _validate_parallel_result_worker(
    result: dict[str, Any], result_path: str, configured_workers: list[int]
) -> tuple[int | None, int | None, int | None]:
    """Validate raw worker fields and return execution/opc/team facts."""

    execution_worker: int | None = None
    logical_tasks: int | None = None
    execution = result.get("execution")
    if execution is not None:
        execution_object = _require_object(execution, f"{result_path}.execution")
        worker = execution_object.get("worker_count")
        if (
            isinstance(worker, bool)
            or not isinstance(worker, int)
            or worker <= 0
        ):
            raise ComparisonInputError(
                f"{result_path}.execution.worker_count must be a positive integer"
            )
        if worker not in configured_workers:
            raise ComparisonInputError(
                f"{result_path}.execution.worker_count={worker} is absent from "
                "configuration.execution_workers"
            )
        execution_worker = worker
        logical_tasks = execution_object.get("logical_tasks")
        if logical_tasks is None:
            raise ComparisonInputError(
                f"{result_path}.execution.logical_tasks is required when execution is present"
            )
        _parallel_scalar_value(
            logical_tasks, f"{result_path}.execution.logical_tasks"
        )
    source = result.get("source")
    if source is None:
        return execution_worker, None, None
    source_object = _require_object(source, f"{result_path}.source")
    opc_cache = source_object.get("opc_cache")
    if opc_cache is None:
        return execution_worker, None, None
    opc_cache_object = _require_object(
        opc_cache, f"{result_path}.source.opc_cache"
    )
    worker = opc_cache_object.get("worker_count")
    if isinstance(worker, bool) or not isinstance(worker, int) or worker <= 0:
        raise ComparisonInputError(
            f"{result_path}.source.opc_cache.worker_count must be a positive integer"
        )
    if worker not in configured_workers:
        raise ComparisonInputError(
            f"{result_path}.source.opc_cache.worker_count={worker} is absent from "
            "configuration.execution_workers"
        )
    teams = opc_cache_object.get("persistent_worker_teams_created")
    if teams is not None:
        _parallel_scalar_value(
            teams, f"{result_path}.source.opc_cache.persistent_worker_teams_created"
        )
    return execution_worker, worker, teams


def _validate_parallel_metric(
    value: Any,
    path: str,
    expected_scope: str | set[str],
    *,
    expected_status: str | None = None,
    vector_length: int | None = None,
    require_vector: bool = False,
    require_positive_vector: bool = False,
    scalar_only: bool = False,
) -> tuple[str, Any | None]:
    obj = _parallel_exact_keys(
        value,
        path,
        {"status", "scope"},
        {"value", "reason"},
    )
    status = obj["status"]
    if not isinstance(status, str) or status not in _PARALLEL_METRIC_STATUSES:
        raise ComparisonInputError(
            f"{path}.status must be one of {sorted(_PARALLEL_METRIC_STATUSES)}"
        )
    if expected_status is not None and status != expected_status:
        raise ComparisonInputError(
            f"{path}.status={status!r}; expected {expected_status!r}"
        )
    _parallel_scope_matches(obj["scope"], expected_scope, f"{path}.scope")
    if status == "measured":
        if "value" not in obj:
            raise ComparisonInputError(f"{path}.value is required when measured")
        if "reason" in obj:
            raise ComparisonInputError(f"{path}.reason must be omitted when measured")
        metric_value = obj["value"]
        _parallel_numeric_value(metric_value, f"{path}.value")
        if scalar_only:
            _parallel_scalar_value(metric_value, f"{path}.value")
        if require_vector and not isinstance(metric_value, list):
            raise ComparisonInputError(f"{path}.value must be a vector")
        if vector_length is not None:
            if not isinstance(metric_value, list):
                raise ComparisonInputError(f"{path}.value must be a sample vector")
            if len(metric_value) != vector_length:
                raise ComparisonInputError(
                    f"{path}.value has {len(metric_value)} samples; expected {vector_length}"
                )
        if require_positive_vector:
            if not isinstance(metric_value, list) or any(
                isinstance(item, bool)
                or not isinstance(item, int)
                or item <= 0
                or item > _PARALLEL_U64_MAX
                for item in metric_value
            ):
                raise ComparisonInputError(
                    f"{path}.value must be a non-empty positive worker vector"
                )
        return status, metric_value
    if "value" in obj:
        raise ComparisonInputError(f"{path}.value must be omitted for {status}")
    reason = obj.get("reason")
    if not isinstance(reason, str) or not reason:
        raise ComparisonInputError(
            f"{path}.reason must be a non-empty string for {status}"
        )
    return status, None


def _validate_parallel_sample_order(
    result: dict[str, Any], result_path: str
) -> tuple[int, list[int]]:
    elapsed = _require_object(result.get("elapsed_ns"), f"{result_path}.elapsed_ns")
    samples = elapsed.get("samples")
    if not isinstance(samples, list):
        raise ComparisonInputError(
            f"{result_path}.elapsed_ns.samples must be a list"
        )
    for index, sample in enumerate(samples):
        if (
            isinstance(sample, bool)
            or not isinstance(sample, int)
            or sample < 0
            or sample > _PARALLEL_U64_MAX
        ):
            raise ComparisonInputError(
                f"{result_path}.elapsed_ns.samples[{index}] must be an unsigned 64-bit integer"
            )
    if samples != sorted(samples):
        raise ComparisonInputError(
            f"{result_path}.elapsed_ns.samples must be sorted in ascending order"
        )
    order = elapsed.get("sample_order")
    if not isinstance(order, list) or len(order) != len(samples):
        raise ComparisonInputError(
            f"{result_path}.elapsed_ns.sample_order must be a permutation of samples"
        )
    if any(isinstance(item, bool) or not isinstance(item, int) for item in order):
        raise ComparisonInputError(
            f"{result_path}.elapsed_ns.sample_order must contain integer indices"
        )
    if sorted(order) != list(range(len(samples))):
        raise ComparisonInputError(
            f"{result_path}.elapsed_ns.sample_order must be a complete permutation"
        )
    if any(
        samples[index] == samples[index + 1]
        and order[index] > order[index + 1]
        for index in range(len(samples) - 1)
    ):
        raise ComparisonInputError(
            f"{result_path}.elapsed_ns.sample_order does not preserve sorted sample identity"
        )
    return len(samples), order


def _parallel_source_chunk_candidates(
    result: dict[str, Any], result_path: str
) -> list[tuple[str, list[int]]]:
    """Return exact per-sample chunk vectors exposed by the result source."""

    candidates: list[tuple[str, list[int]]] = []
    source = result.get("source")
    if source is not None:
        source_object = _require_object(source, f"{result_path}.source")
        simulation = source_object.get("simulation")
        if simulation is not None:
            simulation_object = _require_object(
                simulation, f"{result_path}.source.simulation"
            )
            if "physical_request_count" in simulation_object:
                candidates.append(
                    (
                        _PARALLEL_SIMULATION_CHUNK_SCOPE,
                        _parallel_u64_vector(
                            simulation_object["physical_request_count"],
                            f"{result_path}.source.simulation.physical_request_count",
                        ),
                    )
                )

        cfb_selective = source_object.get("cfb_selective")
        if cfb_selective is not None:
            cfb_selective_object = _require_object(
                cfb_selective, f"{result_path}.source.cfb_selective"
            )
            simulation = cfb_selective_object.get("simulation")
            if simulation is not None:
                simulation_object = _require_object(
                    simulation, f"{result_path}.source.cfb_selective.simulation"
                )
                if "read" in simulation_object:
                    read = simulation_object["read"]
                    if not isinstance(read, list):
                        raise ComparisonInputError(
                            f"{result_path}.source.cfb_selective.simulation.read "
                            "must be a list"
                        )
                    values: list[int] = []
                    for index, phase in enumerate(read):
                        phase_object = _require_object(
                            phase,
                            f"{result_path}.source.cfb_selective.simulation.read[{index}]",
                        )
                        values.append(
                            _parallel_scalar_value(
                                phase_object.get("physical_request_count"),
                                f"{result_path}.source.cfb_selective.simulation.read"
                                f"[{index}].physical_request_count",
                            )
                        )
                    candidates.append((_PARALLEL_CFB_SELECTIVE_CHUNK_SCOPE, values))

        cfb_open_stream = source_object.get("cfb_open_stream")
        if cfb_open_stream is not None:
            cfb_open_stream_object = _require_object(
                cfb_open_stream, f"{result_path}.source.cfb_open_stream"
            )
            simulation = cfb_open_stream_object.get("simulation")
            if simulation is not None:
                simulation_object = _require_object(
                    simulation, f"{result_path}.source.cfb_open_stream.simulation"
                )
                if "samples" in simulation_object:
                    samples = simulation_object["samples"]
                    if not isinstance(samples, list):
                        raise ComparisonInputError(
                            f"{result_path}.source.cfb_open_stream.simulation.samples "
                            "must be a list"
                        )
                    values = []
                    for sample_index, sample in enumerate(samples):
                        sample_object = _require_object(
                            sample,
                            f"{result_path}.source.cfb_open_stream.simulation.samples"
                            f"[{sample_index}]",
                        )
                        per_operation = sample_object.get("per_operation")
                        if not isinstance(per_operation, list):
                            raise ComparisonInputError(
                                f"{result_path}.source.cfb_open_stream.simulation.samples"
                                f"[{sample_index}].per_operation must be a list"
                            )
                        total = 0
                        for operation_index, phase in enumerate(per_operation):
                            phase_object = _require_object(
                                phase,
                                f"{result_path}.source.cfb_open_stream.simulation."
                                f"samples[{sample_index}].per_operation[{operation_index}]",
                            )
                            total += _parallel_scalar_value(
                                phase_object.get("physical_request_count"),
                                f"{result_path}.source.cfb_open_stream.simulation."
                                f"samples[{sample_index}].per_operation[{operation_index}]"
                                ".physical_request_count",
                            )
                            if total > _PARALLEL_U64_MAX:
                                raise ComparisonInputError(
                                    f"{result_path}.source.cfb_open_stream.simulation."
                                    f"samples[{sample_index}].per_operation request count "
                                    "overflows u64"
                                )
                        values.append(total)
                    candidates.append((_PARALLEL_CFB_OPEN_CHUNK_SCOPE, values))

    execution = result.get("execution")
    if execution is not None:
        execution_object = _require_object(execution, f"{result_path}.execution")
        if "deterministic_chunk_count" in execution_object:
            candidates.append(
                (
                    _PARALLEL_CHUNK_SCOPE,
                    _parallel_u64_vector(
                        execution_object["deterministic_chunk_count"],
                        f"{result_path}.execution.deterministic_chunk_count",
                    ),
                )
            )
    return candidates


def _validate_parallel_metrics(report: dict[str, Any], label: str) -> None:
    """Validate optional descriptive parallel metrics when a report emits them."""
    if "parallel_metrics" not in report:
        return
    raw = report["parallel_metrics"]
    envelope = _parallel_exact_keys(
        raw,
        f"{label}.parallel_metrics",
        {
            "schema_version",
            "scope",
            "claim",
            "configured_worker_budget",
            "observed_process_thread_count",
            "cases",
        },
    )
    _require_schema_version(
        envelope["schema_version"],
        f"{label}.parallel_metrics.schema_version",
        _PARALLEL_METRICS_SCHEMA,
    )
    if envelope["scope"] != _PARALLEL_METRICS_SCOPE:
        raise ComparisonInputError(
            f"{label}.parallel_metrics.scope must be {_PARALLEL_METRICS_SCOPE!r}"
        )
    if envelope["claim"] != _PARALLEL_METRICS_CLAIM:
        raise ComparisonInputError(
            f"{label}.parallel_metrics.claim must be {_PARALLEL_METRICS_CLAIM!r}"
        )
    configuration = _require_object(
        report.get("configuration"), f"{label}.configuration"
    )
    configured_workers = _parallel_worker_vector(
        configuration.get("execution_workers"),
        f"{label}.configuration.execution_workers",
    )
    _, configured = _validate_parallel_metric(
        envelope["configured_worker_budget"],
        f"{label}.parallel_metrics.configured_worker_budget",
        "configuration.execution_workers",
        expected_status="measured",
        require_vector=True,
        require_positive_vector=True,
    )
    if not configured:
        raise ComparisonInputError(
            f"{label}.parallel_metrics.configured_worker_budget.value must not be empty"
        )
    if configured != configured_workers:
        raise ComparisonInputError(
            f"{label}.parallel_metrics.configured_worker_budget.value must match "
            f"{label}.configuration.execution_workers"
        )
    _validate_parallel_metric(
        envelope["observed_process_thread_count"],
        f"{label}.parallel_metrics.observed_process_thread_count",
        "process_thread_count",
        expected_status="unavailable",
        scalar_only=True,
    )

    results = report.get("results")
    if not isinstance(results, list):
        raise ComparisonInputError(f"{label}.results must be a list")
    cases = envelope["cases"]
    if not isinstance(cases, list) or len(cases) != len(results):
        raise ComparisonInputError(
            f"{label}.parallel_metrics.cases must match results cardinality"
        )
    for index, (case_value, result_value) in enumerate(zip(cases, results)):
        case_path = f"{label}.parallel_metrics.cases[{index}]"
        result_path = f"{label}.results[{index}]"
        case = _parallel_exact_keys(
            case_value,
            case_path,
            {
                "case",
                "configured_worker_count",
                "observed_local_worker_count",
                "deterministic_task_count",
                "deterministic_chunk_count",
                "lock_wait_ns",
            },
            {"cache_state", "corpus_sha256"},
        )
        result = _require_object(result_value, result_path)
        if case["case"] != result.get("case"):
            raise ComparisonInputError(
                f"{case_path}.case does not match {result_path}.case"
            )
        if "cache_state" in case and (
            not isinstance(case["cache_state"], str) or not case["cache_state"]
        ):
            raise ComparisonInputError(f"{case_path}.cache_state must be non-empty")
        result_cache_state = result.get("cache_state")
        if result_cache_state is not None:
            if not isinstance(result_cache_state, str) or not result_cache_state:
                raise ComparisonInputError(
                    f"{result_path}.cache_state must be a non-empty string"
                )
            if case.get("cache_state") != result_cache_state:
                raise ComparisonInputError(
                    f"{case_path}.cache_state does not match {result_path}.cache_state"
                )
        elif "cache_state" in case:
            raise ComparisonInputError(
                f"{case_path}.cache_state is present but {result_path}.cache_state is absent"
            )
        execution_worker, opc_worker, worker_teams = _validate_parallel_result_worker(
            result, result_path, configured_workers
        )
        corpus = _require_object(result.get("corpus"), f"{result_path}.corpus")
        corpus_sha = corpus.get("archive_sha256")
        if isinstance(corpus_sha, str) and corpus_sha:
            if case.get("corpus_sha256") != corpus_sha:
                raise ComparisonInputError(
                    f"{case_path}.corpus_sha256 does not match {result_path}.corpus"
                )
        elif "corpus_sha256" in case:
            raise ComparisonInputError(
                f"{case_path}.corpus_sha256 requires a non-empty result corpus digest"
            )
        sample_count, sample_order = _validate_parallel_sample_order(result, result_path)
        if execution_worker is None and opc_worker is None:
            expected_configured_status = "not_applicable"
            expected_configured_value = None
            expected_configured_scope = "result.execution.worker_count"
        elif (
            execution_worker is not None
            and opc_worker is not None
            and execution_worker != opc_worker
        ):
            expected_configured_status = "unavailable"
            expected_configured_value = None
            expected_configured_scope = "result.execution.worker_count"
        else:
            expected_configured_status = "measured"
            expected_configured_value = (
                execution_worker if execution_worker is not None else opc_worker
            )
            expected_configured_scope = (
                "result.execution.worker_count"
                if execution_worker is not None
                else "result.source.opc_cache.worker_count"
            )
        configured_status, configured_value = _validate_parallel_metric(
            case["configured_worker_count"],
            f"{case_path}.configured_worker_count",
            {"result.execution.worker_count", "result.source.opc_cache.worker_count"},
            expected_status=expected_configured_status,
            scalar_only=True,
        )
        _parallel_scope_matches(
            case["configured_worker_count"]["scope"],
            expected_configured_scope,
            f"{case_path}.configured_worker_count.scope",
        )
        if configured_status == "measured" and configured_value != expected_configured_value:
            raise ComparisonInputError(
                f"{case_path}.configured_worker_count.value does not match the result worker width"
            )
        if configured_status == "measured" and configured_value not in configured_workers:
            raise ComparisonInputError(
                f"{case_path}.configured_worker_count.value is absent from "
                f"{label}.configuration.execution_workers"
            )

        if opc_worker is None or worker_teams is None or worker_teams == 0:
            expected_observed_status = "not_applicable"
            expected_observed_value = None
        elif worker_teams == 1:
            expected_observed_status = "measured"
            expected_observed_value = opc_worker
        else:
            expected_observed_status = "unavailable"
            expected_observed_value = None
        observed_status, observed_value = _validate_parallel_metric(
            case["observed_local_worker_count"],
            f"{case_path}.observed_local_worker_count",
            "result.source.opc_cache.worker_count_with_one_created_local_worker_team",
            expected_status=expected_observed_status,
            scalar_only=True,
        )
        if observed_status == "measured" and observed_value != expected_observed_value:
            raise ComparisonInputError(
                f"{case_path}.observed_local_worker_count.value does not match the "
                f"{result_path}.source.opc_cache.worker_count"
            )
        if observed_status == "measured" and observed_value not in configured_workers:
            raise ComparisonInputError(
                f"{case_path}.observed_local_worker_count.value is absent from "
                f"{label}.configuration.execution_workers"
            )

        expected_task_status = "measured" if execution_worker is not None else "not_applicable"
        expected_task_value = (
            result["execution"]["logical_tasks"] if execution_worker is not None else None
        )
        task_status, task_value = _validate_parallel_metric(
            case["deterministic_task_count"],
            f"{case_path}.deterministic_task_count",
            "result.execution.logical_tasks",
            expected_status=expected_task_status,
            scalar_only=True,
        )
        if task_status == "measured" and task_value != expected_task_value:
            raise ComparisonInputError(
                f"{case_path}.deterministic_task_count.value does not match "
                f"{result_path}.execution.logical_tasks"
            )

        chunk_candidates = _parallel_source_chunk_candidates(result, result_path)
        if len(chunk_candidates) > 1:
            expected_chunk_status = "unavailable"
            expected_chunk_scope = _PARALLEL_CHUNK_SCOPE
            expected_chunk_values = None
        elif not chunk_candidates:
            expected_chunk_status = (
                "unavailable" if execution_worker is not None else "not_applicable"
            )
            expected_chunk_scope = _PARALLEL_CHUNK_SCOPE
            expected_chunk_values = None
        else:
            expected_chunk_scope, source_values = chunk_candidates[0]
            if not source_values:
                expected_chunk_status = "unavailable"
                expected_chunk_values = None
            else:
                if len(source_values) != sample_count:
                    raise ComparisonInputError(
                        f"{result_path}.{expected_chunk_scope} has {len(source_values)} "
                        f"values; elapsed_ns.samples has {sample_count}"
                    )
                expected_chunk_status = "measured"
                expected_chunk_values = [source_values[index] for index in sample_order]
        chunk_status, chunk_value = _validate_parallel_metric(
            case["deterministic_chunk_count"],
            f"{case_path}.deterministic_chunk_count",
            expected_chunk_scope,
            expected_status=expected_chunk_status,
            vector_length=sample_count,
        )
        if chunk_status == "measured" and chunk_value != expected_chunk_values:
            raise ComparisonInputError(
                f"{case_path}.deterministic_chunk_count.value does not match "
                f"{result_path}.{expected_chunk_scope} after sample_order alignment"
            )
        _validate_parallel_metric(
            case["lock_wait_ns"],
            f"{case_path}.lock_wait_ns",
            "lock_wait_ns",
            expected_status="unavailable",
            scalar_only=True,
        )


def validate_parallel_metrics(report: Any, label: str = "report") -> None:
    """Validate an emitted descriptive parallel-metrics envelope.

    This public wrapper lets summary and packaging tools apply the same
    fail-closed schema and result cross-checks without comparing the envelope
    as a performance metric.
    """

    _validate_parallel_metrics(_require_object(report, label), label)


def _validate_report_identity(
    baseline: dict[str, Any], current: dict[str, Any], policy: dict[str, Any]
) -> None:
    baseline_has_parallel_metrics = "parallel_metrics" in baseline
    current_has_parallel_metrics = "parallel_metrics" in current
    if baseline_has_parallel_metrics != current_has_parallel_metrics:
        raise ComparisonInputError(
            "baseline and current must either both emit parallel_metrics or both omit it"
        )
    for label, report in (("baseline", baseline), ("current", current)):
        _reject_nonfinite_tree(report, label)
        _require_schema_version(
            report.get("schema_version"),
            f"{label}.schema_version",
            SUPPORTED_REPORT_SCHEMA,
        )
        _validate_parallel_metrics(report, label)
        if report.get("tool") != policy["tool_identity"]:
            raise ComparisonInputError(
                f"{label}.tool does not match the policy tool identity"
            )
        if policy.get("require_binary_identity", False) or "binary_identity" in report:
            _validate_binary_identity(
                report.get("binary_identity"),
                f"{label}.binary_identity",
                tool=report["tool"],
            )
        _require_object(report.get("environment"), f"{label}.environment")
        configuration = _require_object(
            report.get("configuration"), f"{label}.configuration"
        )
        for field, expected in policy["expected_configuration"].items():
            if configuration.get(field) != expected:
                raise ComparisonInputError(
                    f"{label}.configuration.{field} does not match policy: "
                    f"{configuration.get(field)!r} != {expected!r}"
                )
        if configuration.get("cases") != policy["required_cases"]:
            raise ComparisonInputError(
                f"{label}.configuration.cases does not match the exact policy case list"
            )
        if policy["require_clean_worktree"]:
            dirty = report["environment"].get("git_worktree_dirty")
            if dirty is not False:
                raise ComparisonInputError(
                    f"{label}.environment.git_worktree_dirty must be false"
                )
    if baseline["tool"] != current["tool"]:
        raise ComparisonInputError("tool identity mismatch between reports")
    if ("binary_identity" in baseline) != ("binary_identity" in current):
        raise ComparisonInputError(
            "baseline and current must either both emit binary_identity or both omit it"
        )
    if baseline["configuration"] != current["configuration"]:
        raise ComparisonInputError("benchmark configuration mismatch between reports")
    revisions = []
    for label, report in (("baseline", baseline), ("current", current)):
        revision = report["environment"].get("git_revision")
        if not isinstance(revision, str) or not revision:
            raise ComparisonInputError(
                f"{label}.environment.git_revision must be a non-empty string"
            )
        revisions.append(revision)
    if policy["require_distinct_revisions"] and revisions[0] == revisions[1]:
        raise ComparisonInputError("reference and current git revisions must differ")
    for field in policy["build_identity_fields"]:
        if field not in baseline["environment"] or field not in current["environment"]:
            raise ComparisonInputError(f"missing build identity field {field!r}")
        before = baseline["environment"][field]
        after = current["environment"][field]
        nullable = field in policy["nullable_build_identity_fields"]
        integer_fields = {
            "logical_cpus_available",
            "total_memory_bytes",
            "page_size_bytes",
        }
        for label, value in (("baseline", before), ("current", after)):
            if value is None:
                if not nullable:
                    raise ComparisonInputError(
                        f"{label} build identity {field!r} must not be null"
                    )
            elif field in integer_fields and (
                isinstance(value, bool) or not isinstance(value, int) or value <= 0
            ):
                raise ComparisonInputError(
                    f"{label} build identity {field!r} must be a positive integer"
                )
            elif field not in integer_fields and (
                not isinstance(value, str) or not value
            ):
                raise ComparisonInputError(
                    f"{label} build identity {field!r} must be a non-empty string"
                )
        if before != after:
            raise ComparisonInputError(
                f"build identity mismatch for {field!r}: {before!r} != {after!r}"
            )
    for field, expected in policy.get("expected_build_identity", {}).items():
        for label, report in (("baseline", baseline), ("current", current)):
            actual = report["environment"].get(field)
            if actual != expected:
                raise ComparisonInputError(
                    f"{label} build identity {field!r} does not match the expected "
                    f"allocator identity: {actual!r} != {expected!r}"
                )
    logical_cpus = baseline["environment"]["logical_cpus_available"]
    expected_workers = sorted(
        {min(requested, logical_cpus) for requested in (1, 2, 4, 8, logical_cpus)}
    )
    for label, report in (("baseline", baseline), ("current", current)):
        workers = report["configuration"].get("execution_workers")
        if workers != expected_workers:
            raise ComparisonInputError(
                f"{label}.configuration.execution_workers does not match the "
                f"derived default {expected_workers!r}"
            )


def _percentile(samples: list[float], percentile: int) -> float:
    ordered = sorted(samples)
    if percentile == 50:
        return float(statistics.median(ordered))
    rank = max(1, math.ceil((percentile / 100.0) * len(ordered)))
    return ordered[rank - 1]


def _latencies(
    result: dict[str, Any], location: str, minimum_samples: int
) -> dict[str, float]:
    elapsed = _require_object(result.get("elapsed_ns"), f"{location}.elapsed_ns")
    if elapsed.get("unit") != "ns":
        raise ComparisonInputError(f"{location}.elapsed_ns.unit must be 'ns'")
    samples_raw = elapsed.get("samples")
    if not isinstance(samples_raw, list):
        raise ComparisonInputError(f"{location}.elapsed_ns.samples must be a list")
    if len(samples_raw) < minimum_samples:
        raise ComparisonInputError(
            f"{location} has {len(samples_raw)} latency samples; "
            f"minimum is {minimum_samples}"
        )
    samples = [
        _finite_number(value, f"{location}.elapsed_ns.samples[{index}]", positive=True)
        for index, value in enumerate(samples_raw)
    ]
    values = {
        "p50": _percentile(samples, 50),
        "p95": _percentile(samples, 95),
        "p99": _percentile(samples, 99),
    }
    reported_values: dict[str, float] = {}
    for name, computed in values.items():
        reported = _finite_number(elapsed.get(name), f"{location}.elapsed_ns.{name}")
        if abs(reported - computed) > 0.5:
            raise ComparisonInputError(
                f"{location}.elapsed_ns.{name}={reported} disagrees with "
                f"samples ({computed})"
            )
        reported_values[name] = reported
    if not values["p50"] <= values["p95"] <= values["p99"]:
        raise ComparisonInputError(
            f"{location}.elapsed_ns percentiles must be non-decreasing"
        )
    if not (
        reported_values["p50"]
        <= reported_values["p95"]
        <= reported_values["p99"]
    ):
        raise ComparisonInputError(
            f"{location}.elapsed_ns reported percentiles must be non-decreasing"
        )
    return values


def _metric_class_for_path(
    path: str, policy: dict[str, Any]
) -> dict[str, Any] | None:
    component_path = path.replace(".", "/")
    matching = [
        metric_class
        for metric_class in policy["metric_classes"]
        if any(
            fnmatch.fnmatchcase(component_path, pattern)
            for pattern in metric_class["path_globs"]
        )
    ]
    if len(matching) > 1:
        names = [item["name"] for item in matching]
        raise ComparisonInputError(
            f"metric {path!r} matches multiple classes: {names}"
        )
    return matching[0] if matching else None


_METRIC_VECTOR_MISSING = object()
_METRIC_VECTOR_KEYS = {"values", "status", "scope"}
_METRIC_VECTOR_STATUSES = {"measured", "not_applicable", "unavailable", "overflow"}

_OPERATION_METRICS_KEYS = {
    "sample_count",
    "sample_indices",
    "alignment",
    "latency_claim",
    "source",
    "process",
    "sink",
    "publication",
    "materialization",
    "cfb_phases",
}
_ALLOCATION_VECTOR_KEYS = (
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
)
_ALLOCATION_METRICS_KEYS = {"status", "scope", *_ALLOCATION_VECTOR_KEYS}
_SOURCE_METRICS_KEYS = {
    "status",
    "counter_scope",
    "logical_read_calls",
    "logical_read_requested_bytes",
    "logical_read_returned_bytes",
    "logical_read_largest_requested_bytes",
    "logical_read_largest_returned_bytes",
    "logical_read_pattern",
    "compressed_bytes",
    "decompressed_bytes",
    "recompressed_bytes",
    "max_concurrent_reads",
}
_SOURCE_NUMERIC_VECTOR_KEYS = (
    "logical_read_calls",
    "logical_read_requested_bytes",
    "logical_read_returned_bytes",
    "logical_read_largest_requested_bytes",
    "logical_read_largest_returned_bytes",
    "max_concurrent_reads",
)
_SOURCE_BOUNDARY_VECTOR_KEYS = (
    "compressed_bytes",
    "decompressed_bytes",
    "recompressed_bytes",
)
_SOURCE_COUNTER_SCOPES = {
    "timed_read_at",
    "untimed_source_replay_only",
    "not_applicable_eager_opc",
    "not_applicable_eager_pptx",
    "not_applicable_eager_docx",
    "not_applicable_filesystem_xlsx",
    "not_applicable_immutable_owned_slice",
    "not_applicable_in_process_sink",
}
_PATTERN_VALUES = {"sequential", "random", "unknown"}
_PROCESS_METRICS_KEYS = {
    "status",
    "user_cpu_ticks",
    "system_cpu_ticks",
    "clock_ticks_per_second",
    "minor_faults",
    "major_faults",
    "voluntary_context_switches",
    "nonvoluntary_context_switches",
    "rss_delta_bytes",
    "peak_rss_bytes",
    "rchar",
    "wchar",
    "read_bytes",
    "write_bytes",
    "cancelled_write_bytes",
    "syscr",
    "syscw",
}
_SINK_METRICS_KEYS = {
    "status",
    "output_bytes",
    "write_status",
    "accepted_bytes",
    "write_calls",
    "largest_write",
    "write_size_buckets",
}
_WRITE_SIZE_BUCKET_KEYS = {
    "status",
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
}
_PUBLICATION_METRICS_KEYS = {"status", "changed_spans", "published_bytes"}
_MATERIALIZATION_METRICS_KEYS = {"status", "opc_parts"}
_CFB_PHASE_METRICS_KEYS = {"status", "open", "plan", "atomic_publication"}
_CFB_PHASE_SET_KEYS = {
    "elapsed_ns",
    "logical_read_calls",
    "logical_read_requested_bytes",
    "logical_read_returned_bytes",
}


def _require_exact_keys(
    value: Any, path: str, expected: set[str]
) -> dict[str, Any]:
    obj = _require_object(value, path)
    actual = set(obj)
    if actual != expected:
        missing = sorted(expected - actual)
        unknown = sorted(actual - expected)
        raise ComparisonInputError(
            f"{path} keys mismatch: missing={missing}, unknown={unknown}"
        )
    return obj


def _validate_metric_status(value: Any, path: str) -> str:
    if not isinstance(value, str) or value not in _METRIC_VECTOR_STATUSES:
        raise ComparisonInputError(
            f"{path} must be one of {sorted(_METRIC_VECTOR_STATUSES)}"
        )
    return value


def _validate_metric_vector(
    value: Any, path: str, sample_count: int
) -> str:
    obj = _require_object(value, path)
    actual = set(obj)
    allowed = {"status", "scope", "values"}
    if not {"status", "scope"} <= actual:
        missing = sorted({"status", "scope"} - actual)
        raise ComparisonInputError(
            f"{path} has a partial MetricVector wrapper; "
            f"missing required keys: {missing}"
        )
    unknown = sorted(actual - allowed)
    if unknown:
        raise ComparisonInputError(
            f"{path} MetricVector wrapper has unknown keys: {unknown}"
        )
    status = _validate_metric_status(obj["status"], f"{path}.status")
    scope = obj["scope"]
    if not isinstance(scope, str) or not scope:
        raise ComparisonInputError(f"{path}.scope must be a non-empty string")
    has_values = "values" in obj
    if status == "measured":
        if not has_values:
            raise ComparisonInputError(
                f"{path}.values is required for a measured MetricVector"
            )
        values = obj["values"]
        if not isinstance(values, list):
            raise ComparisonInputError(f"{path}.values must be a sample vector")
        if len(values) != sample_count:
            raise ComparisonInputError(
                f"{path}.values has {len(values)} samples; expected {sample_count}"
            )
        for index, item in enumerate(values):
            if isinstance(item, bool) or not isinstance(item, int) or item < 0:
                raise ComparisonInputError(
                    f"{path}.values[{index}] must be a non-negative integer"
                )
    elif has_values:
        raise ComparisonInputError(
            f"{path}.values must be omitted for a {status} MetricVector"
        )
    return status


def _validate_pattern_vector(
    value: Any, path: str, sample_count: int
) -> str:
    obj = _require_object(value, path)
    actual = set(obj)
    allowed = {"status", "scope", "values"}
    if not {"status", "scope"} <= actual:
        missing = sorted({"status", "scope"} - actual)
        raise ComparisonInputError(
            f"{path} has a partial PatternVector wrapper; "
            f"missing required keys: {missing}"
        )
    unknown = sorted(actual - allowed)
    if unknown:
        raise ComparisonInputError(
            f"{path} PatternVector wrapper has unknown keys: {unknown}"
        )
    status = _validate_metric_status(obj["status"], f"{path}.status")
    scope = obj["scope"]
    if not isinstance(scope, str) or not scope:
        raise ComparisonInputError(f"{path}.scope must be a non-empty string")
    has_values = "values" in obj
    if status == "measured":
        if not has_values:
            raise ComparisonInputError(
                f"{path}.values is required for a measured PatternVector"
            )
        values = obj["values"]
        if not isinstance(values, list):
            raise ComparisonInputError(f"{path}.values must be a sample vector")
        if len(values) != sample_count:
            raise ComparisonInputError(
                f"{path}.values has {len(values)} samples; expected {sample_count}"
            )
        for index, item in enumerate(values):
            if not isinstance(item, str) or item not in _PATTERN_VALUES:
                raise ComparisonInputError(
                    f"{path}.values[{index}] must be one of "
                    f"{sorted(_PATTERN_VALUES)}"
                )
    elif has_values:
        raise ComparisonInputError(
            f"{path}.values must be omitted for a {status} PatternVector"
        )
    return status


def _validate_status_group(
    value: Any,
    path: str,
    expected_keys: set[str],
    vector_keys: tuple[str, ...],
    sample_count: int,
) -> str:
    obj = _require_exact_keys(value, path, expected_keys)
    status = _validate_metric_status(obj["status"], f"{path}.status")
    for key in vector_keys:
        vector_status = _validate_metric_vector(
            obj[key], f"{path}.{key}", sample_count
        )
        if vector_status != status:
            raise ComparisonInputError(
                f"{path}.status does not match {path}.{key}.status"
            )
    return status


def _validate_phase_set(value: Any, path: str, status: str, sample_count: int) -> None:
    obj = _require_exact_keys(value, path, _CFB_PHASE_SET_KEYS)
    for key in sorted(_CFB_PHASE_SET_KEYS):
        vector_status = _validate_metric_vector(
            obj[key], f"{path}.{key}", sample_count
        )
        if vector_status != status:
            raise ComparisonInputError(
                f"{path}.{key}.status does not match cfb_phases.status"
            )


def _validate_operation_metrics(
    value: Any, path: str, elapsed_samples: list[Any], report_schema: int
) -> None:
    """Validate the exact operation-metrics envelope for report schema 1."""
    # `_validate_report_identity` rejects future report schemas before this
    # validator runs; keep the guard here for direct callers as well.
    if report_schema != SUPPORTED_REPORT_SCHEMA:
        raise ComparisonInputError(
            f"{path} strict validation requires supported report schema "
            f"{SUPPORTED_REPORT_SCHEMA}, got {report_schema}"
        )
    operation_keys = _OPERATION_METRICS_KEYS | (
        {"allocation"} if isinstance(value, dict) and "allocation" in value else set()
    )
    obj = _require_exact_keys(value, path, operation_keys)
    sample_count = len(elapsed_samples)
    declared_sample_count = obj["sample_count"]
    if (
        isinstance(declared_sample_count, bool)
        or not isinstance(declared_sample_count, int)
        or declared_sample_count <= 0
    ):
        raise ComparisonInputError(
            f"{path}.sample_count must be a positive integer"
        )
    if declared_sample_count != sample_count:
        raise ComparisonInputError(
            f"{path}.sample_count={declared_sample_count} does not match "
            f"elapsed_ns.samples length {sample_count}"
        )
    sample_indices = obj["sample_indices"]
    if not isinstance(sample_indices, list):
        raise ComparisonInputError(f"{path}.sample_indices must be a list")
    if len(sample_indices) != sample_count:
        raise ComparisonInputError(
            f"{path}.sample_indices has {len(sample_indices)} samples; "
            f"expected {sample_count}"
        )
    if any(
        isinstance(index, bool) or not isinstance(index, int) or index < 0
        for index in sample_indices
    ):
        raise ComparisonInputError(
            f"{path}.sample_indices must contain non-negative integers"
        )
    if len(set(sample_indices)) != len(sample_indices):
        raise ComparisonInputError(f"{path}.sample_indices must be unique")
    alignment = obj["alignment"]
    if alignment != OPERATION_ALIGNMENT:
        raise ComparisonInputError(
            f"{path}.alignment must be {OPERATION_ALIGNMENT!r}"
        )
    latency_claim = obj["latency_claim"]
    if not isinstance(latency_claim, str) or latency_claim not in {
        EVIDENCE_ONLY_LATENCY_CLAIM,
        COMPARABLE_LATENCY_CLAIM,
    }:
        raise ComparisonInputError(
            f"{path}.latency_claim must be one of "
            f"{[COMPARABLE_LATENCY_CLAIM, EVIDENCE_ONLY_LATENCY_CLAIM]}"
        )
    elapsed_values = [
        _finite_number(
            value,
            f"elapsed_ns.samples[{index}]",
            positive=True,
        )
        for index, value in enumerate(elapsed_samples)
    ]
    if elapsed_values != sorted(elapsed_values):
        raise ComparisonInputError(
            f"{path}.alignment requires elapsed_ns.samples sorted by elapsed time"
        )
    for index in range(1, sample_count):
        if (
            elapsed_values[index] == elapsed_values[index - 1]
            and sample_indices[index] <= sample_indices[index - 1]
        ):
            raise ComparisonInputError(
                f"{path}.sample_indices must increase across tied elapsed samples"
            )

    source = _require_exact_keys(obj["source"], f"{path}.source", _SOURCE_METRICS_KEYS)
    source_status = _validate_metric_status(
        source["status"], f"{path}.source.status"
    )
    if source_status == "measured" and latency_claim != EVIDENCE_ONLY_LATENCY_CLAIM:
        raise ComparisonInputError(
            f"{path}.measured source metrics require "
            f"latency_claim={EVIDENCE_ONLY_LATENCY_CLAIM!r}"
        )
    for key in _SOURCE_NUMERIC_VECTOR_KEYS:
        vector_status = _validate_metric_vector(
            source[key], f"{path}.source.{key}", sample_count
        )
        if vector_status != source_status:
            raise ComparisonInputError(
                f"{path}.source.status does not match {path}.source.{key}.status"
            )
    pattern_status = _validate_pattern_vector(
        source["logical_read_pattern"],
        f"{path}.source.logical_read_pattern",
        sample_count,
    )
    if pattern_status != source_status:
        raise ComparisonInputError(
            f"{path}.source.status does not match "
            f"{path}.source.logical_read_pattern.status"
        )
    boundary_status = "unavailable" if source_status == "measured" else source_status
    for key in _SOURCE_BOUNDARY_VECTOR_KEYS:
        vector_status = _validate_metric_vector(
            source[key], f"{path}.source.{key}", sample_count
        )
        if vector_status != boundary_status:
            raise ComparisonInputError(
                f"{path}.source.{key}.status must be {boundary_status!r} "
                f"for source status {source_status!r}"
            )
    counter_scope = source["counter_scope"]
    if not isinstance(counter_scope, str) or counter_scope not in _SOURCE_COUNTER_SCOPES:
        raise ComparisonInputError(
            f"{path}.source.counter_scope must be one of "
            f"{sorted(_SOURCE_COUNTER_SCOPES)}"
        )
    _validate_status_group(
        obj["process"],
        f"{path}.process",
        _PROCESS_METRICS_KEYS,
        (
            "user_cpu_ticks",
            "system_cpu_ticks",
            "clock_ticks_per_second",
            "minor_faults",
            "major_faults",
            "voluntary_context_switches",
            "nonvoluntary_context_switches",
            "rss_delta_bytes",
            "peak_rss_bytes",
            "rchar",
            "wchar",
            "read_bytes",
            "write_bytes",
            "cancelled_write_bytes",
            "syscr",
            "syscw",
        ),
        sample_count,
    )

    sink = _require_exact_keys(obj["sink"], f"{path}.sink", _SINK_METRICS_KEYS)
    sink_status = _validate_metric_status(sink["status"], f"{path}.sink.status")
    output_status = _validate_metric_vector(
        sink["output_bytes"], f"{path}.sink.output_bytes", sample_count
    )
    if output_status != sink_status:
        raise ComparisonInputError(
            f"{path}.sink.status does not match {path}.sink.output_bytes.status"
        )
    write_status = _validate_metric_status(
        sink["write_status"], f"{path}.sink.write_status"
    )
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        vector_status = _validate_metric_vector(
            sink[key], f"{path}.sink.{key}", sample_count
        )
        if vector_status != write_status:
            raise ComparisonInputError(
                f"{path}.sink.write_status does not match {path}.sink.{key}.status"
            )
    bucket_status = _validate_status_group(
        sink["write_size_buckets"],
        f"{path}.sink.write_size_buckets",
        _WRITE_SIZE_BUCKET_KEYS,
        (
            "bytes_0",
            "bytes_1_to_512",
            "bytes_513_to_4096",
            "bytes_4097_to_16384",
            "bytes_16385_to_65536",
            "bytes_over_65536",
        ),
        sample_count,
    )
    if bucket_status != write_status:
        raise ComparisonInputError(
            f"{path}.sink.write_status does not match "
            f"{path}.sink.write_size_buckets.status"
        )

    _validate_status_group(
        obj["publication"],
        f"{path}.publication",
        _PUBLICATION_METRICS_KEYS,
        ("changed_spans", "published_bytes"),
        sample_count,
    )
    _validate_status_group(
        obj["materialization"],
        f"{path}.materialization",
        _MATERIALIZATION_METRICS_KEYS,
        ("opc_parts",),
        sample_count,
    )
    phases = _require_exact_keys(
        obj["cfb_phases"], f"{path}.cfb_phases", _CFB_PHASE_METRICS_KEYS
    )
    phase_status = _validate_metric_status(
        phases["status"], f"{path}.cfb_phases.status"
    )
    for key in ("open", "plan", "atomic_publication"):
        _validate_phase_set(
            phases[key], f"{path}.cfb_phases.{key}", phase_status, sample_count
        )
    allocation = obj.get("allocation")
    if allocation is not None:
        allocation = _require_exact_keys(
            allocation, f"{path}.allocation", _ALLOCATION_METRICS_KEYS
        )
        allocation_status = _validate_metric_status(
            allocation["status"], f"{path}.allocation.status"
        )
        if allocation["scope"] != "operation_global_system_allocator":
            raise ComparisonInputError(
                f"{path}.allocation.scope must be 'operation_global_system_allocator'"
            )
        for key in _ALLOCATION_VECTOR_KEYS:
            vector_status = _validate_metric_vector(
                allocation[key], f"{path}.allocation.{key}", sample_count
            )
            if vector_status != allocation_status:
                raise ComparisonInputError(
                    f"{path}.allocation.status does not match "
                    f"{path}.allocation.{key}.status"
                )


def _unwrap_metric_vector(value: Any, path: str) -> Any:
    """Return a MetricVector's values and metadata without exposing wrappers.

    Operation metrics serialize each vector as an object containing a status and
    scope, with ``values`` omitted when it is not applicable or unavailable.
    Policy globs intentionally name the logical metric (for example
    ``*write_calls``), not its serialization detail (``write_calls.values``).
    Recognizing the wrapper before policy matching keeps status/scope metadata
    out of numeric traversal while retaining strict validation of malformed
    wrappers.
    """
    if not isinstance(value, dict):
        return _METRIC_VECTOR_MISSING
    keys = set(value)
    if keys == _ALLOCATION_METRICS_KEYS:
        return _METRIC_VECTOR_MISSING
    has_values_or_scope = bool(keys & {"values", "scope"})
    if not has_values_or_scope:
        # Aggregate metric groups also carry a status field, so status alone
        # is not enough to identify a wrapper. A valid MetricStatus with no
        # other field is nevertheless a malformed wrapper-shaped object.
        if (
            keys == {"status"}
            and isinstance(value["status"], str)
            and value["status"] in _METRIC_VECTOR_STATUSES
        ):
            raise ComparisonInputError(
                f"{path} has a partial MetricVector wrapper; "
                "status and scope are required"
            )
        return _METRIC_VECTOR_MISSING
    if not {"status", "scope"} <= keys:
        missing = ", ".join(sorted({"status", "scope"} - keys))
        raise ComparisonInputError(
            f"{path} has a partial MetricVector wrapper; missing {missing}"
        )
    unknown = keys - _METRIC_VECTOR_KEYS
    if unknown:
        names = ", ".join(sorted(unknown))
        raise ComparisonInputError(
            f"{path} MetricVector wrapper has unknown keys: {names}"
        )
    status = value["status"]
    if not isinstance(status, str) or status not in _METRIC_VECTOR_STATUSES:
        raise ComparisonInputError(
            f"{path}.status must be one of {sorted(_METRIC_VECTOR_STATUSES)}"
        )
    scope = value["scope"]
    if not isinstance(scope, str) or not scope:
        raise ComparisonInputError(f"{path}.scope must be a non-empty string")
    has_values = "values" in value
    values = value.get("values")
    if has_values and values is not None and not isinstance(values, list):
        raise ComparisonInputError(
            f"{path}.values must be a sample vector or omitted when unavailable"
        )
    if status == "measured" and (not has_values or values is None):
        raise ComparisonInputError(
            f"{path}.values is required for a measured MetricVector"
        )
    if status != "measured" and has_values and values is not None:
        raise ComparisonInputError(
            f"{path}.values must be omitted for a {status} MetricVector"
        )
    return values, status, scope


def _walk_metrics(
    value: Any,
    path: str,
    policy: dict[str, Any],
    selected: dict[str, tuple[str, float, float, str]],
    vector_metadata: dict[str, tuple[str, str]],
    metric_vector_context: bool,
) -> None:
    if path.rsplit(".", 1)[-1] in _METADATA_FIELD_NAMES:
        return
    vector = (
        _unwrap_metric_vector(value, path)
        if metric_vector_context
        else _METRIC_VECTOR_MISSING
    )
    metric_path = path
    if vector is not _METRIC_VECTOR_MISSING:
        vector_values, status, scope = vector
        vector_metadata[path] = (status, scope)
        if vector_values is None:
            return
        value = vector_values
        values_path = f"{path}.values"
        values_metric_class = _metric_class_for_path(values_path, policy)
        if values_metric_class is not None:
            metric_path = values_path
    metric_class = _metric_class_for_path(metric_path, policy)
    if metric_class is not None:
        presence = metric_class["presence"]
        metric_label = f"{presence} metric"
        if isinstance(value, list):
            if len(value) < policy["minimum_samples"]:
                raise ComparisonInputError(
                    f"{metric_label} {metric_path} has {len(value)} samples; "
                    f"minimum is {policy['minimum_samples']}"
                )
            samples = [
                _finite_number(item, f"{metric_path}[{index}]")
                for index, item in enumerate(value)
            ]
            metric_value = _percentile(samples, 50)
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            metric_value = _finite_number(value, path)
        else:
            raise ComparisonInputError(
                f"{metric_label} {metric_path} must be a numeric scalar or sample vector"
            )
        selected[metric_path] = (
            metric_class["name"],
            metric_value,
            float(metric_class["max_regression_percent"]),
            presence,
        )
        return
    if isinstance(value, dict):
        for key, item in value.items():
            child = f"{path}.{key}" if path else key
            _walk_metrics(
                item,
                child,
                policy,
                selected,
                vector_metadata,
                metric_vector_context,
            )
    elif isinstance(value, list):
        for index, item in enumerate(value):
            _walk_metrics(
                item,
                f"{path}[{index}]",
                policy,
                selected,
                vector_metadata,
                metric_vector_context,
            )


def _collect_metrics(
    result: dict[str, Any],
    policy: dict[str, Any],
    report_schema: int = SUPPORTED_REPORT_SCHEMA,
) -> tuple[
    dict[str, tuple[str, float, float, str]],
    dict[str, tuple[str, str]],
]:
    selected: dict[str, tuple[str, float, float, str]] = {}
    vector_metadata: dict[str, tuple[str, str]] = {}
    for root, value in result.items():
        if root in {"case", "corpus", "elapsed_ns", "output_sha256"}:
            continue
        metric_vector_context = root == "operation_metrics"
        if metric_vector_context:
            elapsed = _require_object(result.get("elapsed_ns"), "elapsed_ns")
            elapsed_samples = elapsed.get("samples")
            if not isinstance(elapsed_samples, list):
                raise ComparisonInputError("elapsed_ns.samples must be a list")
            _validate_operation_metrics(
                value,
                root,
                elapsed_samples,
                report_schema,
            )
        _walk_metrics(
            value,
            root,
            policy,
            selected,
            vector_metadata,
            metric_vector_context,
        )
    return selected, vector_metadata


def _allocator_policy(policy: dict[str, Any]) -> bool:
    return policy["tool_identity"]["instrumentation"] == _ALLOCATOR_INSTRUMENTATION


def _validate_allocator_evidence(
    result: dict[str, Any],
    location: str,
    minimum_samples: int,
    raw_samples: dict[tuple[str, str, int], dict[str, int]],
) -> None:
    case = result.get("case")
    if not isinstance(case, str) or not case.startswith(_FILESYSTEM_CASE_PREFIXES):
        raise ComparisonInputError(
            f"{location}.case must select a filesystem allocator case"
        )
    operation_metrics = _require_object(
        result.get("operation_metrics"), f"{location}.operation_metrics"
    )
    sample_count = operation_metrics.get("sample_count")
    if (
        isinstance(sample_count, bool)
        or not isinstance(sample_count, int)
        or sample_count < minimum_samples
    ):
        raise ComparisonInputError(
            f"{location}.operation_metrics.sample_count must be at least "
            f"{minimum_samples}"
        )
    elapsed = _require_object(result.get("elapsed_ns"), f"{location}.elapsed_ns")
    elapsed_samples = elapsed.get("samples")
    if not isinstance(elapsed_samples, list) or len(elapsed_samples) != sample_count:
        raise ComparisonInputError(
            f"{location}.operation_metrics.sample_count must match elapsed samples"
        )
    if operation_metrics.get("alignment") != OPERATION_ALIGNMENT:
        raise ComparisonInputError(
            f"{location}.operation_metrics.alignment must be {OPERATION_ALIGNMENT!r}"
        )
    allocation = _require_object(
        operation_metrics.get("allocation"), f"{location}.operation_metrics.allocation"
    )
    expected_keys = {"status", "scope", *_ALLOCATOR_VECTOR_FIELDS}
    if set(allocation) != expected_keys:
        raise ComparisonInputError(
            f"{location}.operation_metrics.allocation has an invalid schema"
        )
    if allocation.get("status") != "measured":
        raise ComparisonInputError(
            f"{location}.operation_metrics.allocation.status must be 'measured'"
        )
    if allocation.get("scope") != _ALLOCATOR_SCOPE:
        raise ComparisonInputError(
            f"{location}.operation_metrics.allocation.scope must be "
            f"'{_ALLOCATOR_SCOPE}'"
        )
    vector_lengths: set[int] = set()
    for field in _ALLOCATOR_VECTOR_FIELDS:
        vector = _require_object(
            allocation.get(field), f"{location}.operation_metrics.allocation.{field}"
        )
        if set(vector) != {"values", "status", "scope"}:
            raise ComparisonInputError(
                f"{location}.operation_metrics.allocation.{field} has an invalid schema"
            )
        if vector.get("status") != "measured" or vector.get("scope") != _ALLOCATOR_SCOPE:
            raise ComparisonInputError(
                f"{location}.operation_metrics.allocation.{field} status/scope mismatch"
            )
        values = vector.get("values")
        if not isinstance(values, list) or not values:
            raise ComparisonInputError(
                f"{location}.operation_metrics.allocation.{field}.values must be non-empty"
            )
        if len(values) != sample_count or len(values) < minimum_samples:
            raise ComparisonInputError(
                f"{location}.operation_metrics.allocation.{field}.values cardinality "
                f"{len(values)} does not match sample_count {sample_count}"
            )
        for index, item in enumerate(values):
            if isinstance(item, bool) or not isinstance(item, int) or item < 0:
                raise ComparisonInputError(
                    f"{location}.operation_metrics.allocation.{field}.values[{index}] "
                    "must be a non-negative integer"
                )
        vector_lengths.add(len(values))
    if vector_lengths != {sample_count}:
        raise ComparisonInputError(
            f"{location}.operation_metrics.allocation vectors have mismatched cardinality"
        )
    sample_indices = operation_metrics.get("sample_indices")
    if (
        not isinstance(sample_indices, list)
        or len(sample_indices) != sample_count
        or any(
            isinstance(sample_index, bool) or not isinstance(sample_index, int)
            for sample_index in sample_indices
        )
        or set(sample_indices) != set(range(sample_count))
    ):
        raise ComparisonInputError(
            f"{location}.operation_metrics.sample_indices must be a permutation "
            "of the raw allocator sample indices"
        )
    cache_state = result.get("cache_state")
    if not isinstance(cache_state, str) or not cache_state:
        raise ComparisonInputError(
            f"{location}.cache_state is required for raw allocator alignment"
        )
    for position, sample_index in enumerate(sample_indices):
        raw = raw_samples.get((case, cache_state, sample_index))
        if raw is None:
            raise ComparisonInputError(
                f"{location}.filesystem_evidence has no raw allocator sample "
                f"for sample_index {sample_index}"
            )
        for field in _ALLOCATOR_VECTOR_FIELDS:
            expected = allocation[field]["values"][position]
            observed = raw[field]
            if observed != expected:
                raise ComparisonInputError(
                    f"{location}.filesystem_evidence raw allocator {field} at "
                    f"sample_index {sample_index} does not match "
                    f"operation_metrics.allocation.{field}.values[{position}]"
                )


def _validate_allocator_filesystem_evidence(
    report: dict[str, Any], policy: dict[str, Any], label: str
) -> dict[tuple[str, str, int], dict[str, int]]:
    evidence = report.get("filesystem_evidence")
    if not isinstance(evidence, list) or not evidence:
        raise ComparisonInputError(
            f"{label}.filesystem_evidence is required for allocator policy"
        )
    expected_cases = set(policy["required_cases"])
    expected_configuration = policy["expected_configuration"]
    expected_sample_count = expected_configuration.get("samples_per_case")
    expected_cache_states = expected_configuration.get("filesystem_cache_states")
    if (
        isinstance(expected_sample_count, bool)
        or not isinstance(expected_sample_count, int)
        or expected_sample_count < policy["minimum_samples"]
    ):
        raise ComparisonInputError(
            "allocator policy.samples_per_case must meet minimum_samples"
        )
    if (
        not isinstance(expected_cache_states, list)
        or not expected_cache_states
        or any(not isinstance(state, str) or not state for state in expected_cache_states)
        or len(set(expected_cache_states)) != len(expected_cache_states)
    ):
        raise ComparisonInputError(
            "allocator policy.filesystem_cache_states must be unique non-empty strings"
        )
    expected_sample_total = expected_sample_count * len(expected_cache_states)
    observed_cases: set[str] = set()
    raw_samples: dict[tuple[str, str, int], dict[str, int]] = {}
    for index, raw_item in enumerate(evidence):
        item = _require_object(raw_item, f"{label}.filesystem_evidence[{index}]")
        case = item.get("case")
        if not isinstance(case, str) or not case or not case.startswith(
            _FILESYSTEM_CASE_PREFIXES
        ):
            raise ComparisonInputError(
                f"{label}.filesystem_evidence[{index}].case must select a filesystem case"
            )
        if case in observed_cases:
            raise ComparisonInputError(
                f"{label}.filesystem_evidence contains duplicate case {case!r}"
            )
        observed_cases.add(case)
        item_sample_count = item.get("sample_count")
        if (
            isinstance(item_sample_count, bool)
            or not isinstance(item_sample_count, int)
            or item_sample_count != expected_sample_count
        ):
            raise ComparisonInputError(
                f"{label}.filesystem_evidence[{index}].sample_count does not match "
                f"policy: {item_sample_count!r} != {expected_sample_count!r}"
            )
        item_cache_states = item.get("cache_states")
        if item_cache_states != expected_cache_states:
            raise ComparisonInputError(
                f"{label}.filesystem_evidence[{index}].cache_states does not match policy"
            )
        samples = item.get("samples")
        if not isinstance(samples, list) or len(samples) != expected_sample_total:
            raise ComparisonInputError(
                f"{label}.filesystem_evidence[{index}].samples cardinality does not "
                f"match sample_count/cache_states ({expected_sample_total})"
            )
        sample_indices = {state: [] for state in expected_cache_states}
        for sample_index, raw_sample in enumerate(samples):
            sample = _require_object(
                raw_sample,
                f"{label}.filesystem_evidence[{index}].samples[{sample_index}]",
            )
            allocation = _require_object(
                sample.get("allocation_metrics"),
                f"{label}.filesystem_evidence[{index}].samples[{sample_index}]"
                ".allocation_metrics",
            )
            if allocation.get("status") != "measured" or allocation.get("scope") != _ALLOCATOR_SCOPE:
                raise ComparisonInputError(
                    f"{label}.filesystem_evidence[{index}].samples[{sample_index}] "
                    "allocation status/scope must be measured/system allocator"
                )
            expected_allocation_keys = {"status", "scope", *_ALLOCATOR_VECTOR_FIELDS}
            if set(allocation) != expected_allocation_keys:
                raise ComparisonInputError(
                    f"{label}.filesystem_evidence[{index}].samples[{sample_index}] "
                    "raw allocation_metrics must contain status, scope, and every "
                    "numeric allocator field"
                )
            raw_values: dict[str, int] = {}
            for field in _ALLOCATOR_VECTOR_FIELDS:
                value = allocation[field]
                if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                    raise ComparisonInputError(
                        f"{label}.filesystem_evidence[{index}].samples[{sample_index}] "
                        f"allocation_metrics.{field} must be a non-negative integer"
                    )
                raw_values[field] = value
            cache_state = sample.get("cache_state")
            if cache_state not in sample_indices:
                raise ComparisonInputError(
                    f"{label}.filesystem_evidence[{index}].samples[{sample_index}] "
                    "has an unexpected cache_state"
                )
            sample_number = sample.get("sample_index")
            if (
                isinstance(sample_number, bool)
                or not isinstance(sample_number, int)
                or sample_number < 0
            ):
                raise ComparisonInputError(
                    f"{label}.filesystem_evidence[{index}].samples[{sample_index}] "
                    "sample_index must be a non-negative integer"
                )
            sample_indices[cache_state].append(sample_number)
            raw_samples[(case, cache_state, sample_number)] = raw_values
        expected_indices = set(range(expected_sample_count))
        for cache_state, indices in sample_indices.items():
            if len(indices) != expected_sample_count or set(indices) != expected_indices:
                raise ComparisonInputError(
                    f"{label}.filesystem_evidence[{index}] {cache_state!r} "
                    "sample_index cardinality does not match sample_count"
                )
    if observed_cases != expected_cases:
        raise ComparisonInputError(
            f"{label}.filesystem_evidence cases do not match policy"
        )
    return raw_samples


def _validate_allocator_result_key_cardinality(
    keys: set[tuple[str, str, str]], policy: dict[str, Any], label: str
) -> None:
    expected_states = policy["expected_configuration"].get("filesystem_cache_states")
    if not isinstance(expected_states, list) or not expected_states:
        raise ComparisonInputError(
            "allocator policy.filesystem_cache_states must be non-empty"
        )
    expected_count = policy["expected_result_count"]
    if expected_count % len(expected_states) != 0:
        raise ComparisonInputError(
            "allocator policy.expected_result_count must divide cache-state count"
        )
    state_counts = {state: 0 for state in expected_states}
    for key in keys:
        if len(key) != 3 or key[2] not in state_counts:
            raise ComparisonInputError(
                f"{label} allocator result keys must use the configured cache states"
            )
        state_counts[key[2]] += 1
    expected_per_state = expected_count // len(expected_states)
    if any(count != expected_per_state for count in state_counts.values()):
        raise ComparisonInputError(
            f"{label} allocator result cache-state cardinality does not match policy: "
            f"{state_counts!r}"
        )


def _latency_claim(result: dict[str, Any]) -> str:
    operation_metrics = result.get("operation_metrics")
    if operation_metrics is None:
        return COMPARABLE_LATENCY_CLAIM
    return _require_object(operation_metrics, "operation_metrics")["latency_claim"]


def _source_counter_scope(result: dict[str, Any], location: str) -> str | None:
    operation_metrics = result.get("operation_metrics")
    if operation_metrics is None:
        return None
    operation_metrics = _require_object(operation_metrics, location)
    source = _require_object(operation_metrics.get("source"), f"{location}.source")
    counter_scope = source.get("counter_scope")
    if not isinstance(counter_scope, str):
        raise ComparisonInputError(
            f"{location}.source.counter_scope must be a string"
        )
    return counter_scope


def _optional_metrics_from_selected(
    selected: dict[str, tuple[str, float, float, str]]
) -> dict[str, tuple[str, float, float, str]]:
    return {
        path: metric
        for path, metric in selected.items()
        if metric[3] == "optional"
    }


def _required_metrics_from_selected(
    selected: dict[str, tuple[str, float, float, str]], policy: dict[str, Any]
) -> dict[str, tuple[str, float, float, str]]:
    required = {
        path: metric
        for path, metric in selected.items()
        if metric[3] == "required"
    }
    required_classes = {
        metric_class["name"]
        for metric_class in policy["metric_classes"]
        if metric_class["presence"] == "required"
    }
    present_classes = {metric[0] for metric in required.values()}
    missing_classes = sorted(required_classes - present_classes)
    if missing_classes:
        raise ComparisonInputError(
            "missing required metrics: " + ", ".join(missing_classes)
        )
    return required


def _optional_metrics(
    result: dict[str, Any], policy: dict[str, Any]
) -> dict[str, tuple[str, float, float, str]]:
    selected, _ = _collect_metrics(result, policy)
    return _optional_metrics_from_selected(selected)


def _required_metrics(
    result: dict[str, Any], policy: dict[str, Any]
) -> dict[str, tuple[str, float, float, str]]:
    selected, _ = _collect_metrics(result, policy)
    return _required_metrics_from_selected(selected, policy)


def _compare_metric_vector_metadata(
    case: str,
    baseline: dict[str, tuple[str, str]],
    current: dict[str, tuple[str, str]],
) -> None:
    baseline_paths = set(baseline)
    current_paths = set(current)
    if baseline_paths != current_paths:
        missing = sorted(baseline_paths - current_paths)
        extra = sorted(current_paths - baseline_paths)
        raise ComparisonInputError(
            f"MetricVector path mismatch for {case!r}: "
            f"missing_current={missing}, extra_current={extra}"
        )
    for path in sorted(baseline_paths):
        baseline_status, baseline_scope = baseline[path]
        current_status, current_scope = current[path]
        if baseline_scope != current_scope:
            raise ComparisonInputError(
                f"MetricVector scope mismatch for {case}.{path}: "
                f"baseline={baseline_scope!r}, current={current_scope!r}"
            )
        if baseline_status != current_status:
            raise ComparisonInputError(
                f"MetricVector status mismatch for {case}.{path}: "
                f"baseline={baseline_status!r}, current={current_status!r}"
            )


def _delta_percent(baseline: float, current: float) -> float:
    if baseline == 0:
        return 0.0 if current == 0 else math.inf
    return ((current / baseline) - 1.0) * 100.0


def _comparison_record(
    *,
    case: str,
    corpus: dict[str, Any],
    metric: str,
    metric_class: str,
    baseline: float,
    current: float,
    threshold: float,
    cache_state: str = "",
) -> dict[str, Any]:
    delta = _delta_percent(baseline, current)
    regression = delta > threshold
    record = {
        "case": case,
        "corpus": {
            key: corpus[key]
            for key in (
                "name",
                "generator",
                "shape",
                "payload_kind",
                "archive_sha256",
                "target_entry",
            )
            if key in corpus
        },
        "metric": metric,
        "metric_class": metric_class,
        "baseline": baseline,
        "current": current,
        "delta_percent": delta if math.isfinite(delta) else None,
        "delta_is_infinite": math.isinf(delta),
        "max_regression_percent": threshold,
        "regression": regression,
    }
    if cache_state:
        record["cache_state"] = cache_state
    return record


def compare_reports(
    baseline_raw: Any, current_raw: Any, policy_raw: Any
) -> dict[str, Any]:
    policy = validate_policy(policy_raw)
    baseline = _require_object(baseline_raw, "baseline")
    current = _require_object(current_raw, "current")
    _validate_report_identity(baseline, current, policy)
    baseline_binary_identity = (
        _validate_binary_identity(
            baseline["binary_identity"],
            "baseline.binary_identity",
            tool=baseline["tool"],
        )
        if "binary_identity" in baseline
        else None
    )
    current_binary_identity = (
        _validate_binary_identity(
            current["binary_identity"],
            "current.binary_identity",
            tool=current["tool"],
        )
        if "binary_identity" in current
        else None
    )
    allocator_mode = _allocator_policy(policy)
    baseline_raw_samples: dict[tuple[str, str, int], dict[str, int]] = {}
    current_raw_samples: dict[tuple[str, str, int], dict[str, int]] = {}
    if allocator_mode:
        baseline_raw_samples = _validate_allocator_filesystem_evidence(
            baseline, policy, "baseline"
        )
        current_raw_samples = _validate_allocator_filesystem_evidence(
            current, policy, "current"
        )
    expected_count = policy["expected_result_count"]
    baseline_results = _index_results(
        baseline,
        "baseline",
        expected_count,
        include_cache_state=allocator_mode,
    )
    current_results = _index_results(
        current,
        "current",
        expected_count,
        include_cache_state=allocator_mode,
    )
    baseline_keys = set(baseline_results)
    current_keys = set(current_results)
    if allocator_mode:
        _validate_allocator_result_key_cardinality(baseline_keys, policy, "baseline")
        _validate_allocator_result_key_cardinality(current_keys, policy, "current")
    if baseline_keys != current_keys:
        missing = sorted(key[0] for key in baseline_keys - current_keys)
        extra = sorted(key[0] for key in current_keys - baseline_keys)
        raise ComparisonInputError(
            f"case/corpus key mismatch: missing_current={missing}, extra_current={extra}"
        )
    observed_cases = {key[0] for key in baseline_keys}
    expected_cases = set(policy["required_cases"])
    if observed_cases != expected_cases:
        missing = sorted(expected_cases - observed_cases)
        extra = sorted(observed_cases - expected_cases)
        raise ComparisonInputError(
            f"case set does not match policy: missing={missing}, extra={extra}"
        )
    expected_keys_sha256 = policy["expected_result_keys_sha256"]
    if expected_keys_sha256 is None:
        raise ComparisonInputError(
            "policy has no approved case/corpus manifest digest"
        )
    actual_keys_sha256 = result_key_manifest_sha256(baseline_keys)
    if actual_keys_sha256 != expected_keys_sha256:
        raise ComparisonInputError(
            "case/corpus manifest digest does not match policy: "
            f"{actual_keys_sha256} != {expected_keys_sha256}"
        )

    comparisons: list[dict[str, Any]] = []
    minimum_samples = policy["minimum_samples"]
    latency_thresholds = policy["latency_thresholds_percent"]
    latency_compared_results = 0
    latency_excluded_results = 0
    for key in sorted(baseline_keys):
        case = key[0]
        cache_state = key[2]
        before_result = baseline_results[key]
        after_result = current_results[key]
        if allocator_mode:
            _validate_allocator_evidence(
                before_result,
                f"baseline.{case}[{cache_state}]",
                minimum_samples,
                baseline_raw_samples,
            )
            _validate_allocator_evidence(
                after_result,
                f"current.{case}[{cache_state}]",
                minimum_samples,
                current_raw_samples,
            )
        before_latency = _latencies(before_result, f"baseline.{case}", minimum_samples)
        after_latency = _latencies(after_result, f"current.{case}", minimum_samples)
        before_selected, before_vector_metadata = _collect_metrics(
            before_result, policy
        )
        after_selected, after_vector_metadata = _collect_metrics(after_result, policy)
        before_latency_claim = _latency_claim(before_result)
        after_latency_claim = _latency_claim(after_result)
        before_source_scope = _source_counter_scope(
            before_result, f"baseline.{case}.operation_metrics"
        )
        after_source_scope = _source_counter_scope(
            after_result, f"current.{case}.operation_metrics"
        )
        if before_source_scope != after_source_scope:
            raise ComparisonInputError(
                f"source counter scope mismatch for {case!r}: "
                f"baseline={before_source_scope!r}, current={after_source_scope!r}"
            )
        if before_latency_claim != after_latency_claim:
            raise ComparisonInputError(
                f"latency claim mismatch for {case!r}: "
                f"baseline={before_latency_claim!r}, current={after_latency_claim!r}"
            )
        corpus = before_result["corpus"]
        if before_latency_claim == EVIDENCE_ONLY_LATENCY_CLAIM:
            latency_excluded_results += 1
        else:
            latency_compared_results += 1
            for percentile in ("p50", "p95", "p99"):
                comparisons.append(
                    _comparison_record(
                        case=case,
                        corpus=corpus,
                        metric=f"elapsed_ns.{percentile}",
                        metric_class="latency",
                        baseline=before_latency[percentile],
                        current=after_latency[percentile],
                        threshold=float(latency_thresholds[percentile]),
                        cache_state=cache_state,
                    )
                )
        _compare_metric_vector_metadata(
            case, before_vector_metadata, after_vector_metadata
        )
        before_required = _required_metrics_from_selected(before_selected, policy)
        after_required = _required_metrics_from_selected(after_selected, policy)
        if set(before_required) != set(after_required):
            missing = sorted(set(before_required) - set(after_required))
            extra = sorted(set(after_required) - set(before_required))
            raise ComparisonInputError(
                f"required metric mismatch for {case!r}: "
                f"missing_current={missing}, extra_current={extra}"
            )
        before_optional = _optional_metrics_from_selected(before_selected)
        after_optional = _optional_metrics_from_selected(after_selected)
        if set(before_optional) != set(after_optional):
            missing = sorted(set(before_optional) - set(after_optional))
            extra = sorted(set(after_optional) - set(before_optional))
            raise ComparisonInputError(
                f"optional metric mismatch for {case!r}: "
                f"missing_current={missing}, extra_current={extra}"
            )
        for path in sorted(before_required):
            before_class, before_value, before_threshold, before_presence = (
                before_required[path]
            )
            after_class, after_value, after_threshold, after_presence = after_required[
                path
            ]
            if (before_class, before_threshold, before_presence) != (
                after_class,
                after_threshold,
                after_presence,
            ):
                raise ComparisonInputError(f"metric policy mismatch for {case}.{path}")
            comparisons.append(
                _comparison_record(
                    case=case,
                    corpus=corpus,
                    metric=path,
                    metric_class=before_class,
                    baseline=before_value,
                    current=after_value,
                    threshold=before_threshold,
                    cache_state=cache_state,
                )
            )
        for path in sorted(before_optional):
            before_class, before_value, before_threshold, before_presence = (
                before_optional[path]
            )
            after_class, after_value, after_threshold, after_presence = after_optional[
                path
            ]
            if (before_class, before_threshold, before_presence) != (
                after_class,
                after_threshold,
                after_presence,
            ):
                raise ComparisonInputError(f"metric policy mismatch for {case}.{path}")
            comparisons.append(
                _comparison_record(
                    case=case,
                    corpus=corpus,
                    metric=path,
                    metric_class=before_class,
                    baseline=before_value,
                    current=after_value,
                    threshold=before_threshold,
                    cache_state=cache_state,
                )
            )

    regressions = [item for item in comparisons if item["regression"]]
    return {
        "schema_version": 1,
        "tool": {"name": COMPARATOR_NAME, "version": COMPARATOR_VERSION},
        "policy": {
            "schema_version": policy["schema_version"],
            "policy_id": policy["policy_id"],
            "minimum_samples": minimum_samples,
        },
        "status": "regression" if regressions else "pass",
        "summary": {
            "matched_results": len(baseline_keys),
            "compared_metrics": len(comparisons),
            "regressions": len(regressions),
            "latency_claims": (
                "withheld_instrumentation" if allocator_mode else "compared"
            ),
            "latency_compared_results": latency_compared_results,
            "latency_excluded_results": latency_excluded_results,
        },
        "baseline_revision": baseline["environment"].get("git_revision"),
        "current_revision": current["environment"].get("git_revision"),
        "baseline_binary_sha256": (
            baseline_binary_identity["binary_sha256"]
            if baseline_binary_identity is not None
            else None
        ),
        "current_binary_sha256": (
            current_binary_identity["binary_sha256"]
            if current_binary_identity is not None
            else None
        ),
        "comparisons": comparisons,
        "regressions": regressions,
        "errors": [],
    }


def invalid_report(error: Exception) -> dict[str, Any]:
    return {
        "schema_version": 1,
        "tool": {"name": COMPARATOR_NAME, "version": COMPARATOR_VERSION},
        "status": "invalid",
        "summary": {
            "matched_results": 0,
            "compared_metrics": 0,
            "regressions": 0,
            "latency_compared_results": 0,
            "latency_excluded_results": 0,
        },
        "comparisons": [],
        "regressions": [],
        "errors": [str(error)],
    }


def human_summary(report: dict[str, Any]) -> str:
    status = report["status"].upper()
    summary = report["summary"]
    lines = [
        f"{status}: {summary['matched_results']} matched results, "
        f"{summary['compared_metrics']} metrics, {summary['regressions']} regressions"
    ]
    latency_excluded = summary.get("latency_excluded_results", 0)
    if latency_excluded:
        lines.append(
            "Latency comparison excluded for "
            f"{latency_excluded} evidence-only result(s)"
        )
    if report["status"] == "invalid":
        lines.extend(f"ERROR: {error}" for error in report["errors"])
        return "\n".join(lines) + "\n"
    if report["status"] == "regression":
        for item in report["regressions"]:
            delta = "infinite" if item["delta_is_infinite"] else f"{item['delta_percent']:+.2f}%"
            shape = item["corpus"].get("shape", item["corpus"].get("name", "unknown"))
            cache_state = item.get("cache_state")
            state_label = f"/{cache_state}" if cache_state else ""
            lines.append(
                f"REGRESSION {item['case']}[{shape}{state_label}] {item['metric']}: "
                f"{item['baseline']:.6g} -> {item['current']:.6g} ({delta}; "
                f"limit +{item['max_regression_percent']:.2f}%)"
            )
    else:
        latency = [
            item
            for item in report["comparisons"]
            if item["metric_class"] == "latency"
        ]
        if latency:
            worst = max(
                latency,
                key=lambda item: math.inf
                if item["delta_is_infinite"]
                else item["delta_percent"],
            )
            delta = "infinite" if worst["delta_is_infinite"] else f"{worst['delta_percent']:+.2f}%"
            state_label = f"/{worst['cache_state']}" if worst.get("cache_state") else ""
            lines.append(
                f"Worst latency movement: {worst['case']}{state_label} {worst['metric']} {delta} "
                f"(limit +{worst['max_regression_percent']:.2f}%)"
            )
    return "\n".join(lines) + "\n"


def _write_text(path: Path, contents: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(contents, encoding="utf-8")


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--policy", type=Path, required=True)
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--current", type=Path, required=True)
    parser.add_argument("--json-out", type=Path, required=True)
    parser.add_argument("--summary-out", type=Path)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        policy = load_json(args.policy)
        baseline = load_json(args.baseline)
        current = load_json(args.current)
        report = compare_reports(baseline, current, policy)
    except (ComparisonInputError, OverflowError, TypeError, ValueError) as error:
        report = invalid_report(error)
    machine = json.dumps(report, indent=2, sort_keys=True, allow_nan=False) + "\n"
    human = human_summary(report)
    try:
        _write_text(args.json_out, machine)
        if args.summary_out is not None:
            _write_text(args.summary_out, human)
    except OSError as error:
        print(f"INVALID: cannot write comparator output: {error}", file=sys.stderr)
        return 2
    print(human, end="")
    if report["status"] == "pass":
        return 0
    if report["status"] == "regression":
        return 1
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
