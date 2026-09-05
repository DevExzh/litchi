#!/usr/bin/env python3
"""Validate and summarize one revision of the CRUD baseline capture.

The capture is a small manifest which points at the original Rust reports and
their corpus catalogs.  This script keeps those reports as the raw evidence:
it validates their sample vectors and exact statistics, checks the catalog
binding for every report, and emits repeat summaries.  Normal and allocator
reports are separate evidence streams; allocator-instrumented elapsed values
are never compared with normal elapsed values.
"""

from __future__ import annotations

import argparse
import copy
import json
import math
import re
import sys
from pathlib import Path, PureWindowsPath
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tools import perf_compare  # noqa: E402
from tools import validate_perf_corpus_binding  # noqa: E402


SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")
REPORT_SCHEMA = 1
ALLOCATOR_VECTOR_FIELDS = (
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
CI_METHOD = "two-sided Student's t interval for the mean"


class SummaryError(ValueError):
    """Raised when the capture cannot be summarized safely."""


def _reject_constant(value: str) -> None:
    raise SummaryError(f"non-finite JSON number {value!r}")


def _no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise SummaryError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _finite_tree(value: Any, path: str) -> None:
    if value is None or isinstance(value, (bool, str)):
        return
    if isinstance(value, (int, float)):
        try:
            finite = math.isfinite(float(value))
        except OverflowError as error:
            raise SummaryError(f"{path} contains a number outside the finite range") from error
        if not finite:
            raise SummaryError(f"{path} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _finite_tree(item, f"{path}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            _finite_tree(item, f"{path}.{key}")
        return
    raise SummaryError(f"{path} contains unsupported JSON value type {type(value).__name__}")


def _load_json(path: Path) -> Any:
    try:
        value = json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_pairs,
            parse_constant=_reject_constant,
        )
    except SummaryError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise SummaryError(f"cannot read {path}: {error}") from error
    _finite_tree(value, str(path))
    return value


def _object(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise SummaryError(f"{path} must be an object")
    return value


def _string(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value:
        raise SummaryError(f"{path} must be a non-empty string")
    return value


def _integer(value: Any, path: str, *, minimum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise SummaryError(f"{path} must be an integer")
    if minimum is not None and value < minimum:
        raise SummaryError(f"{path} must be at least {minimum}")
    return value


def _number(value: Any, path: str, *, minimum: float | None = None) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SummaryError(f"{path} must be a finite number")
    if not math.isfinite(float(value)):
        raise SummaryError(f"{path} must be a finite number")
    if minimum is not None and value < minimum:
        raise SummaryError(f"{path} must be at least {minimum}")
    return value


def _sha256(value: Any, path: str) -> str:
    if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
        raise SummaryError(f"{path} must be a 64-character hexadecimal SHA-256")
    return value.lower()


def _canonical(value: Any, path: str) -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        raise SummaryError(f"{path} is not canonical JSON: {error}") from error


def _same(left: Any, right: Any, path: str) -> None:
    if _canonical(left, path) != _canonical(right, path):
        raise SummaryError(f"{path} differs between required matching repeats")


def _resolve_relative(root: Path, value: Any, path: str) -> tuple[str, Path]:
    relative = _string(value, path)
    if Path(relative).is_absolute() or PureWindowsPath(relative).is_absolute():
        raise SummaryError(f"{path} must be relative to the capture root")
    parts = PureWindowsPath(relative).parts
    if ".." in parts:
        raise SummaryError(f"{path} must not escape the capture root")
    root = root.resolve()
    resolved = (root / relative).resolve()
    try:
        resolved.relative_to(root)
    except ValueError as error:
        raise SummaryError(f"{path} escapes the capture root") from error
    if not resolved.is_file():
        raise SummaryError(f"{path} does not name a regular file: {relative}")
    return relative, resolved


def _argv_value(run: dict[str, Any], flag: str) -> str:
    argv = run["argv"]
    positions = [index for index, value in enumerate(argv) if value == flag]
    if len(positions) != 1 or positions[0] + 1 >= len(argv):
        raise SummaryError(f"{run['report']}.argv must contain one {flag} value")
    value = argv[positions[0] + 1]
    if not isinstance(value, str) or not value:
        raise SummaryError(f"{run['report']}.argv {flag} value must be a non-empty string")
    return value


def _historical_root(value: str, relative: str, path: str) -> Path | None:
    """Infer the old capture root from an absolute argv output path.

    Reports are replayed from copied bundles.  Their argv intentionally keeps
    the original absolute paths, so comparing those strings with the replay
    directory would make an otherwise identical bundle fail validation.
    """
    relative_parts = PureWindowsPath(relative).parts
    if not relative_parts or any(part in {"", ".", ".."} for part in relative_parts):
        raise SummaryError(f"{path} has an invalid relative report path")
    candidate = Path(value)
    if not candidate.is_absolute():
        normalized = PureWindowsPath(value).parts
        if tuple(normalized) != tuple(relative_parts):
            raise SummaryError(f"{path} relative path does not match the manifest path")
        return None
    candidate_parts = candidate.parts
    if len(candidate_parts) <= len(relative_parts) or tuple(candidate_parts[-len(relative_parts):]) != tuple(relative_parts):
        raise SummaryError(f"{path} absolute path does not end in its manifest path")
    root_parts = candidate_parts[:-len(relative_parts)]
    if not root_parts:
        raise SummaryError(f"{path} absolute path has no capture root")
    return Path(*root_parts)


def _student_t_critical_95(degrees_of_freedom: int) -> float:
    values = (
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
    if degrees_of_freedom == 0:
        return 0.0
    if degrees_of_freedom <= len(values):
        return values[degrees_of_freedom - 1]
    z = 1.959963984540054
    degrees = float(degrees_of_freedom)
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


def _midpoint(left: int, right: int) -> int:
    # This is the overflow-safe integer midpoint used by the Rust harness.
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def _nearest_rank(values: list[int], percentile: int) -> int:
    index = min((percentile * len(values) + 99) // 100 - 1, len(values) - 1)
    return values[index]


def _rust_statistics(values: list[int], sample_order: list[int]) -> dict[str, Any]:
    if not values:
        raise SummaryError("elapsed_ns.samples must not be empty")
    ordered = sorted(values)
    count = len(ordered)
    mean = 0.0
    squared_deviation_sum = 0.0
    for index, value in enumerate(ordered):
        number = float(value)
        next_count = float(index + 1)
        delta = number - mean
        next_mean = mean + delta / next_count
        squared_deviation_sum += delta * (number - next_mean)
        mean = next_mean
    standard_deviation = (
        math.sqrt(squared_deviation_sum / float(count - 1)) if count > 1 else 0.0
    )
    margin = (
        _student_t_critical_95(count - 1) * standard_deviation / math.sqrt(float(count))
        if count > 1
        else 0.0
    )
    return {
        "unit": "ns",
        "sample_order": sample_order,
        "min": ordered[0],
        "p50": _midpoint(ordered[(count - 1) // 2], ordered[count // 2]),
        "p95": _nearest_rank(ordered, 95),
        "p99": _nearest_rank(ordered, 99),
        "max": ordered[-1],
        "mean": mean,
        "standard_deviation": standard_deviation,
        "confidence_interval_95": {
            "method": CI_METHOD,
            "lower": max(0.0, mean - margin),
            "upper": mean + margin,
        },
    }


def _median_order_statistic_interval(values: list[int]) -> dict[str, Any] | None:
    """Return the narrowest exact distribution-free IID 95% median interval.

    For a sample of size ``n``, the interval [X_(k), X_(n-k+1)] has coverage
    ``1 - 2 * sum(C(n,j) / 2**n for j < k)`` for a population median.  Integer
    arithmetic keeps the selected rank exact instead of relying on a rounded
    binomial probability.
    """
    count = len(values)
    total = 1 << count
    tail = 0
    selected: int | None = None
    for k in range(1, count // 2 + 1):
        tail += math.comb(count, k - 1)
        if (total - 2 * tail) * 20 >= total * 19:
            selected = k
    if selected is None:
        return None
    ordered = sorted(values)
    lower_rank = selected
    upper_rank = count - selected + 1
    return {
        "method": "exact_distribution_free_iid_95_percent_median_order_statistic",
        "assumption": "IID samples",
        "lower_rank": lower_rank,
        "upper_rank": upper_rank,
        "lower_ns": ordered[lower_rank - 1],
        "upper_ns": ordered[upper_rank - 1],
    }


def _validate_elapsed(result: dict[str, Any], expected_samples: int, path: str) -> dict[str, Any]:
    elapsed = _object(result.get("elapsed_ns"), f"{path}.elapsed_ns")
    required = {
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
    missing = sorted(required - set(elapsed))
    if missing:
        raise SummaryError(f"{path}.elapsed_ns is missing {missing[0]!r}")
    if elapsed["unit"] != "ns":
        raise SummaryError(f"{path}.elapsed_ns.unit must be 'ns'")
    if not isinstance(elapsed["samples"], list) or len(elapsed["samples"]) != expected_samples:
        raise SummaryError(
            f"{path}.elapsed_ns.samples must contain exactly {expected_samples} values"
        )
    samples = []
    for index, value in enumerate(elapsed["samples"]):
        samples.append(_integer(value, f"{path}.elapsed_ns.samples[{index}]", minimum=1))
        if samples[-1] > (1 << 64) - 1:
            raise SummaryError(f"{path}.elapsed_ns.samples[{index}] exceeds u64")
    if samples != sorted(samples):
        raise SummaryError(f"{path}.elapsed_ns.samples must be sorted")
    order = elapsed["sample_order"]
    if not isinstance(order, list) or len(order) != expected_samples:
        raise SummaryError(
            f"{path}.elapsed_ns.sample_order must contain exactly {expected_samples} values"
        )
    for index, value in enumerate(order):
        _integer(value, f"{path}.elapsed_ns.sample_order[{index}]")
        if value < 0 or value >= expected_samples:
            raise SummaryError(f"{path}.elapsed_ns.sample_order[{index}] is out of range")
    if len(set(order)) != expected_samples:
        raise SummaryError(f"{path}.elapsed_ns.sample_order must be a permutation")
    for index in range(1, expected_samples):
        if samples[index] == samples[index - 1] and order[index] <= order[index - 1]:
            raise SummaryError(
                f"{path}.elapsed_ns.sample_order must increase across tied samples"
            )
    expected = _rust_statistics(samples, order)
    for field in ("min", "p50", "p95", "p99", "max"):
        _integer(elapsed[field], f"{path}.elapsed_ns.{field}", minimum=0)
        if elapsed[field] != expected[field]:
            raise SummaryError(
                f"{path}.elapsed_ns.{field} does not match its raw samples"
            )
    for field in ("mean", "standard_deviation"):
        _number(elapsed[field], f"{path}.elapsed_ns.{field}", minimum=0.0)
        if elapsed[field] != expected[field]:
            raise SummaryError(
                f"{path}.elapsed_ns.{field} does not match its raw samples"
            )
    interval = _object(elapsed["confidence_interval_95"], f"{path}.elapsed_ns.confidence_interval_95")
    if interval.get("method") != CI_METHOD:
        raise SummaryError(f"{path}.elapsed_ns.confidence_interval_95.method is not the producer method")
    for field in ("lower", "upper"):
        _number(interval.get(field), f"{path}.elapsed_ns.confidence_interval_95.{field}", minimum=0.0)
        if interval[field] != expected["confidence_interval_95"][field]:
            raise SummaryError(
                f"{path}.elapsed_ns.confidence_interval_95.{field} does not match raw samples"
            )
    return {
        "samples": samples,
        "sample_order": order,
        "statistics": expected,
    }


def _validate_capture(capture: Any, root: Path) -> tuple[str, dict[str, dict[str, Any]], list[dict[str, Any]]]:
    capture = _object(capture, "capture.json")
    revision = _string(capture.get("revision"), "capture.revision")
    binaries = _object(capture.get("binaries"), "capture.binaries")
    identities: dict[str, dict[str, Any]] = {}
    for phase in ("normal", "allocator"):
        identity = _object(binaries.get(phase), f"capture.binaries.{phase}")
        identities[phase] = identity
        _sha256(identity.get("sha256"), f"capture.binaries.{phase}.sha256")
        binary_path = _string(identity.get("path"), f"capture.binaries.{phase}.path")
        if not (Path(binary_path).is_absolute() or PureWindowsPath(binary_path).is_absolute()):
            raise SummaryError(f"capture.binaries.{phase}.path must be absolute")
    runs_value = capture.get("runs")
    if not isinstance(runs_value, list) or not runs_value:
        raise SummaryError("capture.runs must be a non-empty list")
    runs: list[dict[str, Any]] = []
    seen_keys: set[tuple[str, int, str]] = set()
    seen_paths: set[str] = set()
    historical_capture_root: Path | None = None
    for index, value in enumerate(runs_value):
        path = f"capture.runs[{index}]"
        run = _object(value, path)
        phase = run.get("phase")
        if phase not in {"normal", "allocator"}:
            raise SummaryError(f"{path}.phase must be 'normal' or 'allocator'")
        repeat = _integer(run.get("repeat"), f"{path}.repeat")
        if repeat not in {1, 2}:
            raise SummaryError(f"{path}.repeat must be 1 or 2")
        selector = _string(run.get("selector"), f"{path}.selector")
        report_rel, report_path = _resolve_relative(root, run.get("report"), f"{path}.report")
        catalog_rel, catalog_path = _resolve_relative(root, run.get("catalog"), f"{path}.catalog")
        argv = run.get("argv")
        if not isinstance(argv, list) or not argv or any(not isinstance(item, str) for item in argv):
            raise SummaryError(f"{path}.argv must be a non-empty string list")
        report_root = _historical_root(
            _argv_value(run, "--json"), report_rel, f"{path}.argv --json"
        )
        catalog_root = _historical_root(
            _argv_value(run, "--corpus-manifest"), catalog_rel, f"{path}.argv --corpus-manifest"
        )
        if report_root != catalog_root:
            raise SummaryError(f"{path}.argv report and catalog paths have different historical roots")
        if historical_capture_root is None:
            historical_capture_root = report_root
        elif report_root != historical_capture_root:
            raise SummaryError(f"{path}.argv does not use the capture's historical root")
        exit_code = _integer(run.get("exit_code"), f"{path}.exit_code")
        if exit_code != 0:
            raise SummaryError(f"{path}.exit_code must be zero")
        key = (phase, repeat, selector)
        if key in seen_keys:
            raise SummaryError(f"capture contains duplicate run {key!r}")
        seen_keys.add(key)
        for relative in (report_rel, catalog_rel):
            if relative in seen_paths:
                raise SummaryError(f"capture references a report or catalog more than once: {relative}")
            seen_paths.add(relative)
        runs.append(
            {
                "phase": phase,
                "repeat": repeat,
                "selector": selector,
                "report": report_rel,
                "report_path": report_path,
                "catalog": catalog_rel,
                "catalog_path": catalog_path,
                "argv": list(argv),
                "exit_code": exit_code,
                "historical_root": report_root,
            }
        )
    grouped: dict[tuple[str, str], set[int]] = {}
    for run in runs:
        grouped.setdefault((run["phase"], run["selector"]), set()).add(run["repeat"])
    for key, repeats in grouped.items():
        if repeats != {1, 2}:
            raise SummaryError(f"missing repeat for phase/selector {key!r}; expected repeats 1 and 2")
    if not any(run["phase"] == "normal" for run in runs):
        raise SummaryError("capture must contain normal runs")
    return revision, identities, runs


def _validate_report_metadata(
    report: dict[str, Any],
    run: dict[str, Any],
    revision: str,
    identity: dict[str, Any],
    expected_samples: int,
    expected_warmups: int,
) -> None:
    path = f"{run['report']}"
    _validate_argv(run, expected_samples, expected_warmups, identity)
    if _integer(report.get("schema_version"), f"{path}.schema_version") != REPORT_SCHEMA:
        raise SummaryError(f"{path}.schema_version must be {REPORT_SCHEMA}")
    tool = _object(report.get("tool"), f"{path}.tool")
    binary = report.get("binary_identity")
    try:
        checked_binary = perf_compare._validate_binary_identity(
            binary, f"{path}.binary_identity", tool=tool
        )
    except Exception as error:
        raise SummaryError(f"{path}.binary_identity failed perf_compare validation: {error}") from error
    if checked_binary["binary_sha256"].lower() != identity["sha256"].lower():
        raise SummaryError(f"{path}.binary_identity.binary_sha256 does not match capture.binaries.{run['phase']}.sha256")
    if checked_binary["path"] != identity["path"]:
        raise SummaryError(f"{path}.binary_identity.path does not match capture.binaries.{run['phase']}.path")
    if "bytes" in identity:
        if _integer(identity["bytes"], f"capture.binaries.{run['phase']}.bytes", minimum=1) != checked_binary["binary_bytes"]:
            raise SummaryError(f"{path}.binary_identity.binary_bytes does not match capture.binaries.{run['phase']}.bytes")
    environment = _object(report.get("environment"), f"{path}.environment")
    if environment.get("git_revision") != revision:
        raise SummaryError(f"{path}.environment.git_revision does not match capture.revision")
    if environment.get("git_worktree_dirty") is not False:
        raise SummaryError(f"{path}.environment.git_worktree_dirty must be false")
    configuration = _object(report.get("configuration"), f"{path}.configuration")
    if _integer(configuration.get("samples_per_case"), f"{path}.configuration.samples_per_case") != expected_samples:
        raise SummaryError(f"{path}.configuration.samples_per_case must be {expected_samples}")
    if _integer(configuration.get("warmup_iterations_per_case"), f"{path}.configuration.warmup_iterations_per_case") != expected_warmups:
        raise SummaryError(
            f"{path}.configuration.warmup_iterations_per_case must be {expected_warmups}"
        )
    cases = configuration.get("cases")
    if cases is not None and (not isinstance(cases, list) or run["selector"] not in cases):
        raise SummaryError(f"{path}.configuration.cases does not contain {run['selector']!r}")
    try:
        perf_compare.validate_parallel_metrics(report, path)
    except Exception as error:
        raise SummaryError(f"{path}.parallel_metrics failed perf_compare validation: {error}") from error


def _validate_argv(
    run: dict[str, Any],
    expected_samples: int,
    expected_warmups: int,
    identity: dict[str, Any],
) -> None:
    argv = run["argv"]
    case_position = [index for index, value in enumerate(argv) if value == "--case"]
    if len(case_position) != 1 or case_position[0] + 1 >= len(argv):
        raise SummaryError(f"{run['report']}.argv must contain one --case value")
    if argv[case_position[0] + 1] != run["selector"]:
        raise SummaryError(f"{run['report']}.argv --case does not match selector")
    for flag, expected in (("--samples", expected_samples), ("--warmup", expected_warmups)):
        if _argv_value(run, flag) != str(expected):
            raise SummaryError(f"{run['report']}.argv {flag} does not match capture phase configuration")
    binary_path = _string(identity.get("path"), f"capture.binaries.{run['phase']}.path")
    binary_positions = [index for index, value in enumerate(argv) if value == binary_path]
    if len(binary_positions) != 1:
        raise SummaryError(f"{run['report']}.argv must contain the phase binary exactly once")
    historical_root = run["historical_root"]
    report_arg = Path(_argv_value(run, "--json"))
    catalog_arg = Path(_argv_value(run, "--corpus-manifest"))
    expected_report = Path(run["report"]) if historical_root is None else historical_root / run["report"]
    expected_catalog = Path(run["catalog"]) if historical_root is None else historical_root / run["catalog"]
    if report_arg != expected_report:
        raise SummaryError(f"{run['report']}.argv --json does not name its report")
    if catalog_arg != expected_catalog:
        raise SummaryError(f"{run['report']}.argv --corpus-manifest does not name its catalog")


def _validate_report_and_select(
    run: dict[str, Any],
    revision: str,
    identity: dict[str, Any],
    expected_samples: int,
    expected_warmups: int,
) -> dict[str, Any]:
    report = _load_json(run["report_path"])
    report = _object(report, run["report"])
    try:
        validate_perf_corpus_binding.validate_paths(run["report_path"], run["catalog_path"])
    except Exception as error:
        raise SummaryError(f"{run['report']} failed corpus binding validation: {error}") from error
    _validate_report_metadata(
        report, run, revision, identity, expected_samples, expected_warmups
    )
    results = report.get("results")
    if not isinstance(results, list) or not results:
        raise SummaryError(f"{run['report']}.results must be a non-empty list")
    result_cases: set[str] = set()
    selected: list[dict[str, Any]] = []
    selected_index: int | None = None
    for index, value in enumerate(results):
        result = _object(value, f"{run['report']}.results[{index}]")
        case = _string(result.get("case"), f"{run['report']}.results[{index}].case")
        if case in result_cases:
            raise SummaryError(f"{run['report']} contains duplicate result case {case!r}")
        result_cases.add(case)
        if case == run["selector"]:
            selected.append(result)
            selected_index = index
    if len(selected) != 1:
        raise SummaryError(
            f"{run['report']} must contain exactly one result for selector {run['selector']!r}"
        )
    result = selected[0]
    corpus = _object(
        result.get("corpus"),
        f"{run['report']}.results[{run['selector']}].corpus",
    )
    selected_elapsed: dict[str, Any] | None = None
    selected_operation_metrics: dict[str, Any] | None = None
    for index, value in enumerate(results):
        result_value = _object(value, f"{run['report']}.results[{index}]")
        result_path = f"{run['report']}.results[{index}]"
        result_elapsed = _validate_elapsed(result_value, expected_samples, result_path)
        operation_metrics = result_value.get("operation_metrics")
        if operation_metrics is not None:
            if not isinstance(operation_metrics, dict):
                raise SummaryError(f"{result_path}.operation_metrics must be an object")
            try:
                perf_compare._validate_operation_metrics(
                    operation_metrics,
                    f"{result_path}.operation_metrics",
                    result_elapsed["samples"],
                    report["schema_version"],
                    elapsed_sample_order=result_elapsed["sample_order"],
                )
            except Exception as error:
                raise SummaryError(
                    f"{result_path}.operation_metrics failed perf_compare validation: {error}"
                ) from error
        output_sha256 = result_value.get("output_sha256")
        if output_sha256 is not None:
            _sha256(output_sha256, f"{result_path}.output_sha256")
        if index == selected_index:
            selected_elapsed = result_elapsed
            selected_operation_metrics = operation_metrics
    if selected_elapsed is None:
        raise SummaryError(f"{run['report']} selected result was not validated")
    oracle = {}
    for field in ("output_sha256", "sink", "oracle", "oracle_sha256", "observation"):
        if field in result:
            oracle[field] = copy.deepcopy(result[field])
    return {
        "run": run,
        "report": report,
        "result": result,
        "corpus": copy.deepcopy(corpus),
        "elapsed": selected_elapsed,
        "oracle": oracle,
        "operation_metrics": selected_operation_metrics,
    }


def _repeat_summary(record: dict[str, Any]) -> dict[str, Any]:
    run = record["run"]
    result = record["result"]
    allocation = _allocation_status(record)
    return {
        "repeat": run["repeat"],
        "report": run["report"],
        "catalog": run["catalog"],
        "p50_ns": record["elapsed"]["statistics"]["p50"],
        "p95_ns": record["elapsed"]["statistics"]["p95"],
        "p99_ns": record["elapsed"]["statistics"]["p99"],
        "median_interval_95": _median_order_statistic_interval(
            record["elapsed"]["samples"]
        ),
        "sample_count": len(record["elapsed"]["samples"]),
        "output_sha256": result.get("output_sha256"),
        "allocation_status": allocation["status"],
    }


def _allocation_status(record: dict[str, Any]) -> dict[str, Any]:
    operation_metrics = record["operation_metrics"]
    if operation_metrics is None:
        return {"status": "unavailable", "reason": "operation_metrics_not_emitted"}
    allocation = operation_metrics.get("allocation")
    if allocation is None:
        return {"status": "unavailable", "reason": "allocation_metrics_not_emitted"}
    if not isinstance(allocation, dict):
        raise SummaryError("operation_metrics.allocation must be an object or null")
    status = allocation.get("status")
    scope = allocation.get("scope")
    if status != "measured":
        return {
            "status": "unavailable",
            "reported_status": status,
            "scope": scope,
            "reason": "allocation_vector_unavailable",
        }
    vectors: dict[str, list[int]] = {}
    for field in ALLOCATOR_VECTOR_FIELDS:
        wrapper = allocation.get(field)
        if not isinstance(wrapper, dict) or wrapper.get("status") != "measured":
            raise SummaryError(f"measured allocation is missing a measured {field} vector")
        values = wrapper.get("values")
        if not isinstance(values, list) or not values:
            raise SummaryError(f"measured allocation {field} vector is empty")
        vectors[field] = []
        for index, value in enumerate(values):
            vectors[field].append(_integer(value, f"allocation.{field}.values[{index}]", minimum=0))
    return {"status": "measured", "scope": scope, "vectors": vectors}


def _allocation_statistics(vectors: dict[str, list[int]]) -> dict[str, dict[str, int]]:
    result: dict[str, dict[str, int]] = {}
    for field, values in vectors.items():
        ordered = sorted(values)
        result[field] = {
            "p50": _midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
            "p95": _nearest_rank(ordered, 95),
            "p99": _nearest_rank(ordered, 99),
        }
    return result


def _allocation_delta(first: dict[str, Any], second: dict[str, Any]) -> dict[str, Any]:
    statuses = {"repeat_1": first["status"], "repeat_2": second["status"]}
    if first["status"] != "measured" or second["status"] != "measured":
        return {
            "status": "unavailable",
            "repeat_status": statuses,
            "reason": "one_or_more_repeats_have_no_allocation_attribution",
        }
    first_stats = _allocation_statistics(first["vectors"])
    second_stats = _allocation_statistics(second["vectors"])
    metrics: dict[str, Any] = {}
    for field in ALLOCATOR_VECTOR_FIELDS:
        first_metric = first_stats[field]
        second_metric = second_stats[field]
        delta = {
            percentile: second_metric[percentile] - first_metric[percentile]
            for percentile in ("p50", "p95", "p99")
        }
        percent: dict[str, float | None] = {}
        for percentile in ("p50", "p95", "p99"):
            baseline = first_metric[percentile]
            percent[percentile] = (
                None if baseline == 0 else delta[percentile] * 100.0 / baseline
            )
        metrics[field] = {
            "repeat_1": first_metric,
            "repeat_2": second_metric,
            "delta": delta,
            "delta_percent": percent,
        }
    return {"status": "measured", "repeat_status": statuses, "metrics": metrics}


def _phase_rows(
    records: list[dict[str, Any]],
    phase: str,
    *,
    cross_phase_records: dict[str, dict[str, Any]],
    drift_threshold: float,
) -> list[dict[str, Any]]:
    grouped: dict[str, dict[int, dict[str, Any]]] = {}
    for record in records:
        selector = record["run"]["selector"]
        grouped.setdefault(selector, {})[record["run"]["repeat"]] = record
    rows = []
    for selector in sorted(grouped):
        repeats = grouped[selector]
        first = repeats[1]
        second = repeats[2]
        _same(first["corpus"], second["corpus"], f"{phase}/{selector}.corpus")
        _same(first["oracle"], second["oracle"], f"{phase}/{selector}.oracle")
        cross_phase_records[selector] = {
            "corpus": copy.deepcopy(first["corpus"]),
            "oracle": copy.deepcopy(first["oracle"]),
        }
        summaries = [_repeat_summary(first), _repeat_summary(second)]
        if phase == "normal":
            first_stats = first["elapsed"]["statistics"]
            second_stats = second["elapsed"]["statistics"]
            drift = {
                field: (second_stats[field] - first_stats[field]) * 100.0 / first_stats[field]
                for field in ("p50", "p95", "p99")
            }
            flags = [
                field for field, value in drift.items() if abs(value) > drift_threshold
            ]
            row = {
                "selector": selector,
                "corpus": copy.deepcopy(first["corpus"]),
                "repeats": summaries,
                "repeat_drift_percent": drift,
                "repeat_drift_threshold_percent": drift_threshold,
                "repeat_drift_flags": flags,
            }
        else:
            first_allocation = _allocation_status(first)
            second_allocation = _allocation_status(second)
            row = {
                "selector": selector,
                "corpus": copy.deepcopy(first["corpus"]),
                "repeats": summaries,
                "latency_comparison": "excluded_for_allocator_instrumentation",
                "allocation_delta": _allocation_delta(first_allocation, second_allocation),
            }
        rows.append(row)
    return rows


def summarize(
    root: Path,
    *,
    samples: int = 500,
    warmups: int = 20,
    allocation_samples: int = 30,
    allocation_warmups: int = 3,
    drift_threshold_percent: float = 5.0,
) -> dict[str, Any]:
    for value, name in (
        (samples, "samples"),
        (warmups, "warmups"),
        (allocation_samples, "allocation_samples"),
        (allocation_warmups, "allocation_warmups"),
    ):
        if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
            raise SummaryError(f"{name} must be a positive integer")
    if isinstance(drift_threshold_percent, bool) or not isinstance(drift_threshold_percent, (int, float)):
        raise SummaryError("drift_threshold_percent must be a finite positive number")
    if not math.isfinite(float(drift_threshold_percent)) or drift_threshold_percent <= 0:
        raise SummaryError("drift_threshold_percent must be a finite positive number")
    root = root.resolve()
    if not root.is_dir():
        raise SummaryError(f"capture root is not a directory: {root}")
    capture = _load_json(root / "capture.json")
    revision, identities, runs = _validate_capture(capture, root)
    records: list[dict[str, Any]] = []
    for run in runs:
        if run["phase"] == "normal":
            expected_samples, expected_warmups = samples, warmups
        else:
            expected_samples, expected_warmups = allocation_samples, allocation_warmups
        records.append(
            _validate_report_and_select(
                run,
                revision,
                identities[run["phase"]],
                expected_samples,
                expected_warmups,
            )
        )
    normal = [record for record in records if record["run"]["phase"] == "normal"]
    allocator = [record for record in records if record["run"]["phase"] == "allocator"]
    normal_records: dict[str, dict[str, Any]] = {}
    allocator_records: dict[str, dict[str, Any]] = {}
    drift_threshold = float(drift_threshold_percent)
    normal_rows = _phase_rows(
        normal,
        "normal",
        cross_phase_records=normal_records,
        drift_threshold=drift_threshold,
    )
    allocator_rows = _phase_rows(
        allocator,
        "allocator",
        cross_phase_records=allocator_records,
        drift_threshold=drift_threshold,
    )
    for selector in sorted(set(normal_records) & set(allocator_records)):
        _same(
            normal_records[selector]["corpus"],
            allocator_records[selector]["corpus"],
            f"{selector}.corpus across normal/allocator",
        )
        _same(
            normal_records[selector]["oracle"],
            allocator_records[selector]["oracle"],
            f"{selector}.oracle across normal/allocator",
        )
    drift_flags = [
        {
            "selector": row["selector"],
            "flags": row["repeat_drift_flags"],
            "drift_percent": {
                field: row["repeat_drift_percent"][field]
                for field in row["repeat_drift_flags"]
            },
        }
        for row in normal_rows
        if row["repeat_drift_flags"]
    ]
    return {
        "schema_version": 1,
        "tool": "litchi-crud-baseline-summarizer",
        "performance_claim": "none",
        "revision": revision,
        "configuration": {
            "samples": samples,
            "warmups": warmups,
            "allocation_samples": allocation_samples,
            "allocation_warmups": allocation_warmups,
            "normal_repeat_drift_threshold_percent": drift_threshold,
        },
        "verification": {
            "raw_vectors_validated": True,
            "exact_elapsed_statistics_validated": True,
            "p50_convention": "producer_integer_midpoint_of_middle_sorted_samples",
            "oracle_cross_repeat_validated": True,
            "corpus_binding_validated_for_every_run": True,
            "normal_repeat_pairs": len(normal_rows),
            "allocator_repeat_pairs": len(allocator_rows),
            "allocator_latency_comparison": "excluded",
            "median_uncertainty": (
                "exact_distribution_free_iid_95_percent_order_statistic_interval"
            ),
        },
        "binaries": {
            phase: copy.deepcopy(identity) for phase, identity in identities.items()
        },
        "normal": normal_rows,
        "normal_repeat_drift_flags": drift_flags,
        "allocator": allocator_rows,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True, help="capture bundle root containing capture.json")
    parser.add_argument("--output", type=Path, help="summary JSON path (default: ROOT/summary.json)")
    parser.add_argument("--samples", type=int, default=500)
    parser.add_argument("--warmups", type=int, default=20)
    parser.add_argument("--allocation-samples", type=int, default=30)
    parser.add_argument("--allocation-warmups", type=int, default=3)
    parser.add_argument("--drift-threshold-percent", type=float, default=5.0)
    args = parser.parse_args(argv)
    try:
        summary = summarize(
            args.root,
            samples=args.samples,
            warmups=args.warmups,
            allocation_samples=args.allocation_samples,
            allocation_warmups=args.allocation_warmups,
            drift_threshold_percent=args.drift_threshold_percent,
        )
        output = args.output or (args.root / "summary.json")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(
            json.dumps(summary, sort_keys=True, indent=2, allow_nan=False) + "\n",
            encoding="utf-8",
        )
    except (SummaryError, OSError) as error:
        parser.exit(1, f"CRUD baseline summary refused: {error}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
