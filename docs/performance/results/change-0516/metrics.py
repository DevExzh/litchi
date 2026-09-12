"""Extract descriptive XLSX/OOXML metrics from the retained 0516 captures.

This helper is intentionally smaller than the evidence verifier.  It reads
any capture lanes that already exist, checks the raw timing vectors against
the producer statistics, and emits matched before/after observations.  A
missing candidate lane is represented in the output so that a partial run is
useful while it is being collected.  A report that is present but has missing
or malformed rows is reported as incomplete and is never silently omitted.

The output is descriptive only.  It does not authorize a performance claim
or turn a lower elapsed time into a speedup claim.  GNU ``time`` RSS is read
only from normal native capture logs; allocator capture logs are excluded.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from tools.summarize_crud_baseline import _validate_elapsed  # noqa: E402


MAIN_CASES = (
    "xlsx_one_cell_commit",
    "xlsx_one_percent_commit",
    "xlsx_one_cell_commit_save",
    "xlsx_one_percent_commit_save",
)
MAIN_SHAPES = ("tiny", "medium", "dense-wide")
GUARD_SHAPES = MAIN_SHAPES
GUARD_SCENARIOS = (
    "cold-first-cell-read",
    "cold-same-one-cell",
    "cold-same-one-percent",
    "warm-same-one-cell",
    "warm-same-one-percent",
    "warm-changed-one-cell",
    "warm-changed-one-percent",
)
FALLBACK_SCENARIOS = (
    "warm-changed-one-cell",
    "warm-changed-one-percent",
)
LANES = (
    "preflight",
    "pilot",
    "r1",
    "r2",
    "allocator-r1",
    "allocator-r2",
)
FAMILIES = ("main", "guard", "fallback")
STATISTICS = ("p50", "mean", "p95", "p99")
ALLOC_FIELDS = (
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
RSS_RE = re.compile(
    r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$",
    re.MULTILINE,
)
U64_MAX = (1 << 64) - 1
ADVERSE_THRESHOLD_PERCENT = 5.0


class MetricsError(ValueError):
    """A present evidence item cannot be interpreted safely."""


def _reject_constant(value: str) -> None:
    raise MetricsError(f"non-finite JSON number {value!r}")


def _no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise MetricsError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _load_json(path: Path) -> Any:
    if path.is_symlink() or not path.is_file():
        raise MetricsError(f"{path.name} is not a regular file")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_pairs,
            parse_constant=_reject_constant,
        )
    except MetricsError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise MetricsError(f"cannot load {path.name}: {error}") from error


def _integer(value: Any, context: str, minimum: int = 0,
             maximum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise MetricsError(f"{context} must be an integer")
    if value < minimum:
        raise MetricsError(f"{context} is below {minimum}")
    if maximum is not None and value > maximum:
        raise MetricsError(f"{context} exceeds {maximum}")
    return value


def _finite_number(value: Any, context: str, minimum: float = 0.0) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise MetricsError(f"{context} must be a number")
    result = float(value)
    if not math.isfinite(result) or result < minimum:
        raise MetricsError(f"{context} is not a finite number at least {minimum}")
    return result


def _problem(issues: list[dict[str, Any]], context: dict[str, Any],
             message: str, kind: str = "invalid") -> None:
    issues.append({**context, "kind": kind, "message": message})


def _relative(path: Path, root: Path) -> str:
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return str(path)


def _p50_midpoint(ordered: list[int]) -> float:
    left = ordered[(len(ordered) - 1) // 2]
    right = ordered[len(ordered) // 2]
    return (float(left) + float(right)) / 2.0


def _p50_lower(ordered: list[int]) -> int:
    return ordered[(len(ordered) - 1) // 2]


def _nearest_rank(ordered: list[int], percentile: int) -> int:
    index = min((percentile * len(ordered) + 99) // 100 - 1, len(ordered) - 1)
    return ordered[index]


def _vector_summary(values: list[int], lower_p50: bool = False) -> dict[str, Any]:
    if not values:
        raise MetricsError("cannot summarize an empty metric vector")
    ordered = sorted(values)
    return {
        "count": len(values),
        "min": ordered[0],
        "p50": _p50_lower(ordered) if lower_p50 else _p50_midpoint(ordered),
        "mean": sum(values) / float(len(values)),
        "p95": _nearest_rank(ordered, 95),
        "p99": _nearest_rank(ordered, 99),
        "max": ordered[-1],
    }


def _guard_percentile(ordered: list[int], percentile: int) -> int:
    # This is the public guard's producer convention.  Main harness timing
    # statistics are validated through _validate_elapsed instead.
    return ordered[(len(ordered) - 1) * percentile // 100]


def _guard_latency(values: list[int], context: str) -> dict[str, Any]:
    if not values:
        raise MetricsError(f"{context} has no elapsed samples")
    ordered = sorted(values)
    expected = {
        "min_ns": ordered[0],
        "p50_ns": _guard_percentile(ordered, 50),
        "p95_ns": _guard_percentile(ordered, 95),
        "p99_ns": _guard_percentile(ordered, 99),
        "max_ns": ordered[-1],
        "mean_ns": sum(values) / float(len(values)),
    }
    return expected


def _change(before: Any, after: Any) -> dict[str, Any]:
    left = _finite_number(before, "comparison before", minimum=0.0)
    right = _finite_number(after, "comparison after", minimum=0.0)
    if left <= 0.0:
        return {
            "before": before,
            "after": after,
            "candidate_minus_control_percent": None,
            "adverse_over_5_percent": None,
            "comparison_status": "non_positive_control",
        }
    change = (right / left - 1.0) * 100.0
    return {
        "before": before,
        "after": after,
        "candidate_minus_control_percent": change,
        "adverse_over_5_percent": change > ADVERSE_THRESHOLD_PERCENT,
        "comparison_status": "descriptive",
    }


def _capture_context(stage: str, family: str, lane: str,
                     path: Path, root: Path) -> dict[str, Any]:
    return {
        "stage": stage,
        "family": family,
        "lane": lane,
        "path": _relative(path, root),
    }


def _paths(root: Path, stage: str, family: str, lane: str) -> dict[str, Path]:
    prefix = {"main": "", "guard": "guard-", "fallback": "fallback-"}[family]
    stem = f"{prefix}{lane}"
    directory = root / stage
    return {
        "report": directory / f"{stem}-report.json",
        "receipt": directory / f"{stem}-receipt.json",
        "log": directory / f"{stem}.log",
    }


def _expected_keys(family: str) -> list[tuple[str, str]]:
    if family == "main":
        return [(shape, case) for shape in MAIN_SHAPES for case in MAIN_CASES]
    scenarios = GUARD_SCENARIOS if family == "guard" else FALLBACK_SCENARIOS
    return [(shape, scenario) for shape in GUARD_SHAPES for scenario in scenarios]


def _row_key(family: str, key: tuple[str, str]) -> str:
    first, second = key
    return f"{first}/{second}"


def _empty_capture(stage: str, family: str, lane: str, paths: dict[str, Path],
                   root: Path) -> dict[str, Any]:
    return {
        "stage": stage,
        "family": family,
        "lane": lane,
        "allocator": lane.startswith("allocator-"),
        "path": _relative(paths["report"], root),
        "report_present": paths["report"].is_file() and not paths["report"].is_symlink(),
        "complete": False,
        "sample_count": None,
        "warmups": None,
        "rows": [],
        "row_map": {},
        "missing_rows": [_row_key(family, key) for key in _expected_keys(family)],
        "issues": [],
    }


def _validate_main_allocation(row: dict[str, Any], samples: int,
                              allocator: bool, context: str) -> dict[str, Any]:
    operation = row.get("operation_metrics")
    if not isinstance(operation, dict):
        raise MetricsError(f"{context}.operation_metrics is missing")
    allocation = operation.get("allocation")
    if not isinstance(allocation, dict):
        raise MetricsError(f"{context}.operation_metrics.allocation is missing")
    status = allocation.get("status")
    if status != "measured":
        if allocator:
            raise MetricsError(f"{context} allocator allocation is not measured")
        if status not in {"unavailable", "not_applicable", "overflow"}:
            raise MetricsError(f"{context} allocation status is invalid")
        return {"status": status, "scope": allocation.get("scope"), "fields": {}}

    fields: dict[str, dict[str, Any]] = {}
    for field in ALLOC_FIELDS:
        vector = allocation.get(field)
        if not isinstance(vector, dict) or vector.get("status") != "measured":
            raise MetricsError(f"{context}.allocation.{field} is not measured")
        values = vector.get("values")
        if not isinstance(values, list) or len(values) != samples:
            raise MetricsError(
                f"{context}.allocation.{field}.values must contain {samples} values"
            )
        parsed = [
            _integer(value, f"{context}.allocation.{field}.values[{index}]")
            for index, value in enumerate(values)
        ]
        fields[field] = _vector_summary(parsed)

    before = fields["live_bytes_before"]
    after = fields["live_bytes_after"]
    region = fields["region_peak_live_bytes"]
    peak_after = fields["peak_live_bytes_after"]
    # Compare the same p50 values used by the retained allocator evidence,
    # while retaining every counter and peak summary for review.
    if region["p50"] < before["p50"] or region["p50"] < after["p50"]:
        raise MetricsError(f"{context} region peak precedes a live-byte baseline")
    if region["p50"] > peak_after["p50"]:
        raise MetricsError(f"{context} region peak exceeds peak_live_bytes_after")
    # Derive the incremental peak from the raw vectors.  Keep the calculation
    # close to their validation so the published field cannot be confused
    # with an allocator counter supplied by the producer.
    before_values = allocation["live_bytes_before"]["values"]
    region_values = allocation["region_peak_live_bytes"]["values"]
    derived = []
    for index, (region_value, before_value) in enumerate(zip(region_values, before_values)):
        region_number = _integer(
            region_value,
            f"{context}.allocation.region_peak_live_bytes.values[{index}]",
        )
        before_number = _integer(
            before_value,
            f"{context}.allocation.live_bytes_before.values[{index}]",
        )
        if region_number < before_number:
            raise MetricsError(f"{context} incremental region peak is negative")
        derived.append(region_number - before_number)
    fields["incremental_region_peak_bytes"] = _vector_summary(derived)
    return {
        "status": "measured",
        "scope": allocation.get("scope"),
        "fields": fields,
    }


def _main_row_summary(row: Any, key: tuple[str, str], samples: int,
                      allocator: bool, context: str) -> dict[str, Any]:
    if not isinstance(row, dict):
        raise MetricsError(f"{context} is not an object")
    try:
        elapsed = _validate_elapsed(row, samples, context)
    except Exception as error:
        raise MetricsError(f"elapsed statistics are invalid: {error}") from error
    statistics = elapsed["statistics"]
    latency = None if allocator else {
        "unit": "ns",
        "sample_count": samples,
        "min": statistics["min"],
        "p50": statistics["p50"],
        "mean": statistics["mean"],
        "p95": statistics["p95"],
        "p99": statistics["p99"],
        "max": statistics["max"],
    }
    allocation = _validate_main_allocation(row, samples, allocator, context)
    return {
        "row_key": _row_key("main", key),
        "shape": key[0],
        "case": key[1],
        "valid": True,
        "sample_count": samples,
        "latency": latency,
        "elapsed_excluded": allocator,
        "allocation": allocation,
    }


def _main_capture(stage: str, lane: str, path: Path, root: Path,
                  issues: list[dict[str, Any]]) -> dict[str, Any]:
    context_base = _capture_context(stage, "main", lane, path, root)
    capture = _empty_capture(stage, "main", lane, _paths(root, stage, "main", lane), root)
    local_issues: list[dict[str, Any]] = []
    try:
        report = _load_json(path)
    except MetricsError as error:
        _problem(local_issues, context_base, str(error))
        capture["issues"] = local_issues
        issues.extend(local_issues)
        return capture
    if not isinstance(report, dict):
        _problem(local_issues, context_base, "report is not an object")
        capture["issues"] = local_issues
        issues.extend(local_issues)
        return capture

    configuration = report.get("configuration")
    samples: int | None = None
    warmups: int | None = None
    if not isinstance(configuration, dict):
        _problem(local_issues, context_base, "configuration is missing")
    else:
        try:
            samples = _integer(
                configuration.get("samples_per_case"),
                f"{context_base['path']}.configuration.samples_per_case",
                minimum=1,
            )
            warmups = _integer(
                configuration.get("warmup_iterations_per_case"),
                f"{context_base['path']}.configuration.warmup_iterations_per_case",
            )
        except MetricsError as error:
            _problem(local_issues, context_base, str(error))
    capture["sample_count"] = samples
    capture["warmups"] = warmups

    expected = _expected_keys("main")
    expected_set = set(expected)
    rows = report.get("results")
    seen: set[tuple[str, str]] = set()
    if not isinstance(rows, list):
        _problem(local_issues, context_base, "results is missing or is not a list")
        rows = []
    for index, row in enumerate(rows):
        row_context = f"{context_base['path']}.results[{index}]"
        entry: dict[str, Any] = {
            "row_index": index,
            "valid": False,
        }
        key: tuple[str, str] | None = None
        if isinstance(row, dict):
            case = row.get("case")
            corpus = row.get("corpus")
            shape = corpus.get("shape") if isinstance(corpus, dict) else None
            if isinstance(shape, str) and isinstance(case, str):
                key = (shape, case)
                entry.update({"row_key": _row_key("main", key), "shape": shape, "case": case})
                if key in seen:
                    _problem(local_issues, {**context_base, "row": entry["row_key"]},
                             "duplicate row key")
                seen.add(key)
                if key not in expected_set:
                    _problem(local_issues, {**context_base, "row": entry["row_key"]},
                             "row key is outside the expected main corpus")
            else:
                _problem(local_issues, {**context_base, "row": index},
                         "row lacks string case and corpus.shape")
        else:
            _problem(local_issues, {**context_base, "row": index}, "row is not an object")
        if key is not None and samples is not None and key in expected_set:
            try:
                summary = _main_row_summary(
                    row, key, samples, lane.startswith("allocator-"), row_context,
                )
            except MetricsError as error:
                entry["error"] = str(error)
                _problem(local_issues, {**context_base, "row": entry["row_key"]}, str(error))
            else:
                entry.update(summary)
                capture["row_map"].setdefault(key, summary)
        elif key is not None and samples is None:
            entry["error"] = "sample count is unavailable"
        capture["rows"].append(entry)

    missing = [key for key in expected if key not in seen]
    capture["missing_rows"] = [_row_key("main", key) for key in missing]
    for key in missing:
        _problem(
            local_issues,
            {**context_base, "row": _row_key("main", key)},
            "expected row is missing from a present report",
            kind="missing-row",
        )
    capture["issues"] = local_issues
    capture["complete"] = bool(
        not local_issues
        and samples is not None
        and len(capture["row_map"]) == len(expected)
    )
    issues.extend(local_issues)
    return capture


def _guard_allocation_samples(value: Any, samples: int, allocator: bool,
                              context: str) -> dict[str, Any]:
    if not isinstance(value, list) or len(value) != samples:
        raise MetricsError(f"{context} must contain {samples} allocation samples")
    field_values = {field: [] for field in ALLOC_FIELDS}
    for index, sample in enumerate(value):
        sample_context = f"{context}[{index}]"
        if not isinstance(sample, dict):
            raise MetricsError(f"{sample_context} is not an object")
        status = sample.get("status")
        if not allocator:
            if status != "unavailable" or sample.get("scope") != "operation_global_system_allocator":
                raise MetricsError(f"{sample_context} is not an unavailable native sample")
            if set(sample) != {"status", "scope"}:
                raise MetricsError(f"{sample_context} publishes unavailable values")
            continue
        if status != "measured" or sample.get("scope") != "operation_global_system_allocator":
            raise MetricsError(f"{sample_context} is not a measured allocator sample")
        for field in ALLOC_FIELDS:
            field_values[field].append(
                _integer(sample.get(field), f"{sample_context}.{field}", maximum=U64_MAX)
            )
        if field_values["region_peak_live_bytes"][-1] < field_values["live_bytes_before"][-1]:
            raise MetricsError(f"{sample_context} region peak precedes live_bytes_before")
        if field_values["region_peak_live_bytes"][-1] < field_values["live_bytes_after"][-1]:
            raise MetricsError(f"{sample_context} region peak precedes live_bytes_after")
        if field_values["region_peak_live_bytes"][-1] > field_values["peak_live_bytes_after"][-1]:
            raise MetricsError(f"{sample_context} region peak exceeds peak_live_bytes_after")
        if field_values["peak_live_bytes_after"][-1] < field_values["peak_live_bytes_before"][-1]:
            raise MetricsError(f"{sample_context} peak live bytes decrease")
    if not allocator:
        return {"status": "unavailable", "fields": {}}
    fields = {
        field: _vector_summary(values, lower_p50=True)
        for field, values in field_values.items()
    }
    fields["incremental_region_peak_bytes"] = _vector_summary(
        [
            region - before
            for region, before in zip(
                field_values["region_peak_live_bytes"],
                field_values["live_bytes_before"],
            )
        ],
        lower_p50=True,
    )
    return {"status": "measured", "fields": fields}


def _guard_row_summary(row: Any, shape: str, scenario: str, samples: int,
                       warmups: int, allocator: bool, context: str,
                       family: str) -> dict[str, Any]:
    if not isinstance(row, dict):
        raise MetricsError(f"{context} is not an object")
    if row.get("scenario") != scenario:
        raise MetricsError(f"{context}.scenario differs from its map key")
    if row.get("sample_count") != samples:
        raise MetricsError(f"{context}.sample_count differs from report")
    if row.get("warmup_iterations") != warmups:
        raise MetricsError(f"{context}.warmup_iterations differs from report")
    elapsed = row.get("elapsed_ns")
    if not isinstance(elapsed, list) or len(elapsed) != samples:
        raise MetricsError(f"{context}.elapsed_ns must contain {samples} values")
    values = [
        _integer(value, f"{context}.elapsed_ns[{index}]", minimum=1, maximum=U64_MAX)
        for index, value in enumerate(elapsed)
    ]
    if row.get("sample_indices") != list(range(samples)):
        raise MetricsError(f"{context}.sample_indices are not chronological")
    expected_stats = _guard_latency(values, context)
    stats = row.get("stats")
    if not isinstance(stats, dict):
        raise MetricsError(f"{context}.stats is missing")
    if set(stats) != set(expected_stats):
        raise MetricsError(f"{context}.stats schema differs")
    for field, expected in expected_stats.items():
        actual = stats[field]
        if field == "mean_ns":
            if _finite_number(actual, f"{context}.stats.{field}") != expected:
                raise MetricsError(f"{context}.stats.{field} differs from raw samples")
        elif _integer(actual, f"{context}.stats.{field}") != expected:
            raise MetricsError(f"{context}.stats.{field} differs from raw samples")
    allocation = _guard_allocation_samples(
        row.get("allocation_samples"), samples, allocator,
        f"{context}.allocation_samples",
    )
    stats_out = None if allocator else {
        "unit": "ns",
        "sample_count": samples,
        "min": expected_stats["min_ns"],
        "p50": expected_stats["p50_ns"],
        "mean": expected_stats["mean_ns"],
        "p95": expected_stats["p95_ns"],
        "p99": expected_stats["p99_ns"],
        "max": expected_stats["max_ns"],
    }
    return {
        "row_key": _row_key(family, (shape, scenario)),
        "shape": shape,
        "scenario": scenario,
        "valid": True,
        "sample_count": samples,
        "latency": stats_out,
        "elapsed_excluded": allocator,
        "allocation": allocation,
    }


def _guard_capture(stage: str, family: str, lane: str, path: Path, root: Path,
                   issues: list[dict[str, Any]]) -> dict[str, Any]:
    context_base = _capture_context(stage, family, lane, path, root)
    capture = _empty_capture(stage, family, lane, _paths(root, stage, family, lane), root)
    local_issues: list[dict[str, Any]] = []
    try:
        report = _load_json(path)
    except MetricsError as error:
        _problem(local_issues, context_base, str(error))
        capture["issues"] = local_issues
        issues.extend(local_issues)
        return capture
    if not isinstance(report, dict):
        _problem(local_issues, context_base, "report is not an object")
        capture["issues"] = local_issues
        issues.extend(local_issues)
        return capture
    try:
        samples = _integer(report.get("samples"), f"{context_base['path']}.samples", minimum=1)
        warmups = _integer(report.get("warmups"), f"{context_base['path']}.warmups")
    except MetricsError as error:
        samples = None
        warmups = None
        _problem(local_issues, context_base, str(error))
    capture["sample_count"] = samples
    capture["warmups"] = warmups

    expected = _expected_keys(family)
    expected_set = set(expected)
    shapes = report.get("shapes")
    if not isinstance(shapes, list):
        _problem(local_issues, context_base, "shapes is missing or is not a list")
        shapes = []
    seen_shapes: set[str] = set()
    seen: set[tuple[str, str]] = set()
    scenarios_expected = set(GUARD_SCENARIOS if family == "guard" else FALLBACK_SCENARIOS)
    for shape_index, shape_report in enumerate(shapes):
        shape_context = f"{context_base['path']}.shapes[{shape_index}]"
        if not isinstance(shape_report, dict) or not isinstance(shape_report.get("shape"), str):
            _problem(local_issues, {**context_base, "shape": shape_index},
                     "shape entry lacks a string shape")
            continue
        shape = shape_report["shape"]
        if shape in seen_shapes:
            _problem(local_issues, {**context_base, "shape": shape}, "duplicate shape entry")
        seen_shapes.add(shape)
        if shape not in GUARD_SHAPES:
            _problem(local_issues, {**context_base, "shape": shape},
                     "shape is outside the expected guard corpus")
        scenarios = shape_report.get("scenarios")
        if not isinstance(scenarios, list):
            _problem(local_issues, {**context_base, "shape": shape},
                     "scenarios is missing or is not a list")
            continue
        seen_scenarios: set[str] = set()
        for scenario_index, row in enumerate(scenarios):
            row_context = f"{shape_context}.scenarios[{scenario_index}]"
            scenario = row.get("scenario") if isinstance(row, dict) else None
            if not isinstance(scenario, str):
                _problem(local_issues, {**context_base, "shape": shape,
                                         "row": scenario_index},
                         "scenario entry lacks a string scenario")
                continue
            key = (shape, scenario)
            row_key = _row_key(family, key)
            if scenario in seen_scenarios:
                _problem(local_issues, {**context_base, "row": row_key},
                         "duplicate scenario row")
            seen_scenarios.add(scenario)
            seen.add(key)
            if scenario not in scenarios_expected:
                _problem(local_issues, {**context_base, "row": row_key},
                         "scenario is outside the expected guard corpus")
            entry: dict[str, Any] = {
                "row_index": scenario_index,
                "row_key": row_key,
                "shape": shape,
                "scenario": scenario,
                "valid": False,
            }
            if samples is not None and warmups is not None and key in expected_set:
                try:
                    summary = _guard_row_summary(
                        row, shape, scenario, samples, warmups,
                        lane.startswith("allocator-"), row_context, family,
                    )
                except MetricsError as error:
                    entry["error"] = str(error)
                    _problem(local_issues, {**context_base, "row": row_key}, str(error))
                else:
                    entry.update(summary)
                    capture["row_map"].setdefault(key, summary)
            capture["rows"].append(entry)
        missing_scenarios = scenarios_expected - seen_scenarios
        for scenario in sorted(missing_scenarios):
            _problem(
                local_issues,
                {**context_base, "row": _row_key(family, (shape, scenario))},
                "expected scenario row is missing from a present report",
                kind="missing-row",
            )

    for shape in GUARD_SHAPES:
        if shape not in seen_shapes:
            _problem(local_issues, {**context_base, "shape": shape},
                     "expected shape is missing from a present report", kind="missing-row")
    missing = [key for key in expected if key not in seen]
    capture["missing_rows"] = [_row_key(family, key) for key in missing]
    capture["issues"] = local_issues
    capture["complete"] = bool(
        not local_issues
        and samples is not None
        and warmups is not None
        and len(capture["row_map"]) == len(expected)
    )
    issues.extend(local_issues)
    return capture


def _artifact_issues(stage: str, family: str, lane: str, paths: dict[str, Path],
                     root: Path, report_present: bool,
                     issues: list[dict[str, Any]]) -> None:
    context = _capture_context(stage, family, lane, paths["report"], root)
    for name, path in paths.items():
        if name == "report":
            continue
        if report_present and (path.is_symlink() or not path.is_file()):
            _problem(issues, {**context, "artifact": name},
                     f"present report has no regular {name} artifact", kind="missing-artifact")
    if not report_present:
        present = [name for name, path in paths.items()
                   if path.is_symlink() or path.exists()]
        if present:
            _problem(issues, {**context, "artifact": "report"},
                     f"capture has {', '.join(present)} but its report is missing",
                     kind="missing-artifact")


def _parse_rss(path: Path) -> int:
    if path.is_symlink() or not path.is_file():
        raise MetricsError(f"{path.name} is not a regular file")
    try:
        text = path.read_text(encoding="utf-8", errors="strict")
    except (OSError, UnicodeError) as error:
        raise MetricsError(f"cannot read {path.name}: {error}") from error
    values = RSS_RE.findall(text)
    if len(values) != 1:
        raise MetricsError(f"{path.name} must contain exactly one GNU time RSS line")
    result = int(values[0])
    if result <= 0:
        raise MetricsError(f"{path.name} reports non-positive RSS")
    return result


def _load_captures(root: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]],
                                          dict[tuple[str, str, str], dict[str, Any]]]:
    issues: list[dict[str, Any]] = []
    captures: list[dict[str, Any]] = []
    by_key: dict[tuple[str, str, str], dict[str, Any]] = {}
    for stage in ("before", "after"):
        for family in FAMILIES:
            for lane in LANES:
                paths = _paths(root, stage, family, lane)
                report_present = paths["report"].is_file() and not paths["report"].is_symlink()
                artifact_present = report_present or any(
                    path.is_symlink() or path.exists() for path in paths.values()
                )
                if not artifact_present:
                    continue
                local_start = len(issues)
                _artifact_issues(stage, family, lane, paths, root, report_present, issues)
                if report_present:
                    if family == "main":
                        capture = _main_capture(stage, lane, paths["report"], root, issues)
                    else:
                        capture = _guard_capture(stage, family, lane, paths["report"], root, issues)
                else:
                    capture = _empty_capture(stage, family, lane, paths, root)
                capture["artifact_paths"] = {
                    name: _relative(path, root) for name, path in paths.items()
                    if path.is_symlink() or path.exists()
                }
                capture["issue_count"] = len(issues) - local_start
                if any(item.get("kind") == "missing-artifact"
                       for item in issues[local_start:]):
                    capture["complete"] = False
                captures.append(capture)
                by_key[stage, family, lane] = capture
    return captures, issues, by_key


def _individual_rows(captures: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for capture in captures:
        for row in capture["rows"]:
            result.append({
                "stage": capture["stage"],
                "family": capture["family"],
                "lane": capture["lane"],
                "allocator": capture["allocator"],
                "report": capture["path"],
                **row,
            })
    return result


def _public_capture(capture: dict[str, Any]) -> dict[str, Any]:
    # row_map uses tuple keys for safe internal matching and cannot be
    # serialized as a JSON object.  The public rows list is the complete
    # row-level view; the map remains private to comparison construction.
    return {key: value for key, value in capture.items() if key != "row_map"}


def _allocation_comparison(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    if before.get("status") != "measured" or after.get("status") != "measured":
        return {
            "status": "not_compared",
            "before_status": before.get("status"),
            "after_status": after.get("status"),
            "fields": {},
        }
    before_fields = before.get("fields", {})
    after_fields = after.get("fields", {})
    names = sorted(set(before_fields) | set(after_fields))
    fields: dict[str, Any] = {}
    for name in names:
        if name not in before_fields or name not in after_fields:
            fields[name] = {
                "status": "missing_side",
                "before": before_fields.get(name),
                "after": after_fields.get(name),
            }
            continue
        left = before_fields[name]
        right = after_fields[name]
        fields[name] = {
            "status": "compared",
            "statistics": {
                metric: _change(left[metric], right[metric])
                for metric in ("p50", "mean")
            },
            "before_summary": left,
            "after_summary": right,
        }
    return {"status": "compared", "fields": fields}


def _row_comparison(family: str, lane: str, key: tuple[str, str],
                    before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    before_latency = before["latency"]
    after_latency = after["latency"]
    result = {
        "family": family,
        "lane": lane,
        "row_key": _row_key(family, key),
        "status": "compared",
        "latency": None if before.get("elapsed_excluded") or after.get("elapsed_excluded") else {
            metric: _change(before_latency[metric], after_latency[metric])
            for metric in STATISTICS
        },
        "elapsed_excluded": bool(before.get("elapsed_excluded") or after.get("elapsed_excluded")),
        "allocation": _allocation_comparison(
            before.get("allocation", {"status": "missing"}),
            after.get("allocation", {"status": "missing"}),
        ),
    }
    return result


def _comparisons(by_key: dict[tuple[str, str, str], dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for family in FAMILIES:
        for lane in LANES:
            before = by_key.get(("before", family, lane))
            after = by_key.get(("after", family, lane))
            if before is None and after is None:
                continue
            pair: dict[str, Any] = {
                "family": family,
                "lane": lane,
                "status": "compared" if before is not None and after is not None else (
                    "missing_before" if before is None else "missing_after"
                ),
                "before_report": before.get("path") if before else None,
                "after_report": after.get("path") if after else None,
                "before_complete": before.get("complete") if before else False,
                "after_complete": after.get("complete") if after else False,
                "rows": [],
                "missing_before_rows": [],
                "missing_after_rows": [],
            }
            if before is None or after is None:
                pair["status"] = "missing_before" if before is None else "missing_after"
                result.append(pair)
                continue
            keys = _expected_keys(family)
            for key in keys:
                left = before.get("row_map", {}).get(key)
                right = after.get("row_map", {}).get(key)
                if left is None:
                    pair["missing_before_rows"].append(_row_key(family, key))
                if right is None:
                    pair["missing_after_rows"].append(_row_key(family, key))
                if left is not None and right is not None:
                    pair["rows"].append(_row_comparison(family, lane, key, left, right))
            if pair["missing_before_rows"] or pair["missing_after_rows"]:
                pair["status"] = "partial"
            result.append(pair)
    return result


def _rss(root: Path, captures: list[dict[str, Any]],
         issues: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    values: list[dict[str, Any]] = []
    by_key: dict[tuple[str, str, str], dict[str, Any]] = {}
    for capture in captures:
        if capture["allocator"]:
            continue
        paths = _paths(root, capture["stage"], capture["family"], capture["lane"])
        log = paths["log"]
        if not log.is_file() or log.is_symlink():
            continue
        try:
            value = _parse_rss(log)
        except MetricsError as error:
            context = _capture_context(capture["stage"], capture["family"],
                                        capture["lane"], log, root)
            _problem(issues, context, str(error), kind="invalid-rss")
            continue
        item = {
            "stage": capture["stage"],
            "family": capture["family"],
            "lane": capture["lane"],
            "log": _relative(log, root),
            "rss_kib": value,
            "scope": "GNU time whole-child RSS from a normal native capture log",
        }
        values.append(item)
        by_key[capture["stage"], capture["family"], capture["lane"]] = item
    comparisons: list[dict[str, Any]] = []
    for family in FAMILIES:
        for lane in LANES:
            if lane.startswith("allocator-"):
                continue
            before = by_key.get(("before", family, lane))
            after = by_key.get(("after", family, lane))
            if before is None and after is None:
                continue
            if before is None or after is None:
                comparisons.append({
                    "family": family,
                    "lane": lane,
                    "status": "missing_before" if before is None else "missing_after",
                    "before_rss_kib": before["rss_kib"] if before else None,
                    "after_rss_kib": after["rss_kib"] if after else None,
                })
                continue
            comparison = _change(before["rss_kib"], after["rss_kib"])
            comparisons.append({
                "family": family,
                "lane": lane,
                "status": "compared",
                "rss": comparison,
                "scope": "GNU time whole-child RSS from normal native logs only; allocator logs excluded",
            })
    return values, comparisons


def collect(root: Path = HERE) -> dict[str, Any]:
    """Return descriptive metrics for all currently retained capture lanes."""
    root = root.resolve()
    captures, issues, by_key = _load_captures(root)
    rss_values, rss_comparisons = _rss(root, captures, issues)
    missing_lanes = []
    for stage in ("before", "after"):
        for family in FAMILIES:
            for lane in LANES:
                if (stage, family, lane) in by_key:
                    continue
                paths = _paths(root, stage, family, lane)
                missing_lanes.append({
                    "stage": stage,
                    "family": family,
                    "lane": lane,
                    "report": _relative(paths["report"], root),
                })
    report_count = sum(1 for capture in captures if capture["report_present"])
    status = "empty" if report_count == 0 else (
        "partial" if issues or len(missing_lanes) else "complete"
    )
    return {
        "schema": "litchi-0516-descriptive-metrics-v1",
        "performance_claim": "none",
        "comparison_scope": (
            "Matched before/after descriptive observations; latency changes and adverse-over-5-percent flags only"
        ),
        "allocation_scope": (
            "Allocator counters and peaks are compared only for allocator lanes; normal lanes remain uninstrumented"
        ),
        "rss_scope": (
            "GNU time whole-child RSS from normal native logs only; allocator elapsed/RSS is excluded"
        ),
        "adverse_threshold_percent": ADVERSE_THRESHOLD_PERCENT,
        "statistics": list(STATISTICS),
        "captures": [_public_capture(capture) for capture in captures],
        "individual_rows": _individual_rows(captures),
        "comparisons": _comparisons(by_key),
        "rss": {
            "individual": rss_values,
            "comparisons": rss_comparisons,
        },
        "missing_lanes": missing_lanes,
        "issues": issues,
        "status": status,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract descriptive metrics from retained 0516 capture reports."
    )
    parser.add_argument(
        "--root",
        type=Path,
        default=HERE,
        help="0516 evidence directory (default: this script's directory)",
    )
    args = parser.parse_args(argv)
    try:
        result = collect(args.root)
    except (OSError, MetricsError) as error:
        print(json.dumps({
            "schema": "litchi-0516-descriptive-metrics-v1",
            "performance_claim": "none",
            "status": "error",
            "issues": [{"kind": "fatal", "message": str(error)}],
        }, indent=2, sort_keys=True))
        return 1
    print(json.dumps(result, indent=2, sort_keys=True))
    return 1 if result["issues"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
