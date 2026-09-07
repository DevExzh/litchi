#!/usr/bin/env python3
"""Compare authenticated 0457 control and candidate summaries.

The control and candidate derive scripts authenticate their own receipts,
reports, source custody, and report oracles before writing summaries.  This
script consumes those summaries and their retained per-sample vectors.  It
keeps the two retained-result contracts explicit and reports descriptive
numeric endpoint deltas only; it does not make an ordinary Commit/Patch or
speedup claim.

The elapsed-time median interval is an independent nonparametric bootstrap of
the two thirty-sample vectors.  A small fixed LCG, rather than Python's
version-sensitive global PRNG, makes the 10,000-resample result reproducible.
Run this script only after the candidate summary exists.
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 457
SAMPLES = 30
REPEATS = ("R1", "R2")
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
QUANTILES = ("p50", "p95", "p99")
THRESHOLD_PERCENT = 5.0
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_SEED = 4_570_457
UINT64_MASK = (1 << 64) - 1
LCG_MULTIPLIER = 6_364_136_223_846_793_005
LCG_INCREMENT = 1_442_695_040_888_963_407

PROCESS_FIELDS = ("peak_rss_bytes", "rss_delta_bytes")
ALLOCATION_FIELDS = (
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
    "heap_growth_bytes",
    "region_peak_above_entry",
)
VECTOR_FIELDS = ("elapsed_ns", *PROCESS_FIELDS, *ALLOCATION_FIELDS)


class ComparisonError(ValueError):
    pass


def fail(label: str, message: str) -> None:
    raise ComparisonError(f"{label}: {message}")


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected an object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(label, "expected a non-empty string")
    return value


def integer(value: Any, label: str, *, minimum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        fail(label, "expected an integer")
    if minimum is not None and value < minimum:
        fail(label, f"expected an integer >= {minimum}")
    return value


def number(value: Any, label: str) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        fail(label, "expected a finite number")
    return value


def clean(value: float | int) -> int | float:
    numeric = float(value)
    return int(numeric) if numeric.is_integer() else numeric


def close(left: Any, right: Any) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return False
    if not isinstance(left, (int, float)) or not isinstance(right, (int, float)):
        return False
    return math.isclose(float(left), float(right), rel_tol=1e-12, abs_tol=1e-9)


def permutation(value: Any, label: str) -> list[int]:
    if not isinstance(value, list) or len(value) != SAMPLES:
        fail(label, f"expected {SAMPLES} sample indices")
    result = [integer(item, f"{label}[{index}]", minimum=0) for index, item in enumerate(value)]
    if sorted(result) != list(range(SAMPLES)):
        fail(label, "must be a permutation of retained sample indices")
    return result


def relative_path(value: Any, label: str) -> str:
    path = Path(text(value, label))
    if path.is_absolute() or ".." in path.parts:
        fail(label, "must be a bundle-relative path")
    return path.as_posix()


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    rank = max(1, math.ceil(fraction * len(ordered)))
    return ordered[rank - 1]


def statistics(values: list[int | float], label: str) -> dict[str, int | float]:
    if len(values) != SAMPLES:
        fail(label, f"expected {SAMPLES} values")
    ordered = sorted(float(number(value, f"{label}[{index}]")) for index, value in enumerate(values))
    middle = (ordered[(SAMPLES - 1) // 2] + ordered[SAMPLES // 2]) / 2.0
    return {
        "count": SAMPLES,
        "min": clean(ordered[0]),
        "p50": clean(middle),
        "p95": clean(percentile(ordered, 0.95)),
        "p99": clean(percentile(ordered, 0.99)),
        "max": clean(ordered[-1]),
        "mean": clean(sum(ordered) / SAMPLES),
    }


def validate_reported_statistics(row: dict[str, Any], field: str, values: list[int | float], label: str) -> None:
    statistics_row = obj(obj(row.get("statistics"), f"{label}.statistics").get(field), f"{label}.statistics.{field}")
    expected = statistics(values, f"{label}.vectors.{field}")
    for key, expected_value in expected.items():
        if key not in statistics_row or not close(statistics_row[key], expected_value):
            fail(f"{label}.statistics.{field}.{key}", "does not match its retained vector")


def vector(row: dict[str, Any], field: str, label: str) -> list[int | float] | None:
    raw = obj(row.get("vectors"), f"{label}.vectors").get(field)
    if raw is None:
        statistics_row = obj(row.get("statistics"), f"{label}.statistics").get(field)
        if statistics_row is not None:
            fail(f"{label}.{field}", "statistics must remain unavailable when its vector is unavailable")
        return None
    vector_row = obj(raw, f"{label}.vectors.{field}")
    order = permutation(vector_row.get("sample_order"), f"{label}.vectors.{field}.sample_order")
    elapsed_order = vector_row.get("elapsed_order")
    sample_index_order = vector_row.get("sample_index_order")
    if not isinstance(elapsed_order, list) or len(elapsed_order) != SAMPLES:
        fail(f"{label}.vectors.{field}.elapsed_order", f"expected {SAMPLES} values")
    if not isinstance(sample_index_order, list) or len(sample_index_order) != SAMPLES:
        fail(f"{label}.vectors.{field}.sample_index_order", f"expected {SAMPLES} values")
    values = [number(item, f"{label}.vectors.{field}.elapsed_order[{index}]") for index, item in enumerate(elapsed_order)]
    indexed = [number(item, f"{label}.vectors.{field}.sample_index_order[{index}]") for index, item in enumerate(sample_index_order)]
    for position, sample_index in enumerate(order):
        if not close(indexed[sample_index], values[position]):
            fail(f"{label}.vectors.{field}", "sample-index order does not bind to sample order")
    validate_reported_statistics(row, field, values, label)
    return indexed


def validate_summary(path: Path, role: str) -> tuple[dict[str, Any], dict[tuple[str, str, str], dict[str, Any]]]:
    summary = obj(load(path, str(path)), str(path))
    expected_schema = "litchi-0457-odp-control-summary-v1" if role == "control" else "litchi-0457-odp-source-tail-candidate-summary-v1"
    if summary.get("schema") != expected_schema or summary.get("change") != CHANGE:
        fail(str(path), f"unexpected {role} summary schema or change")
    matrix = obj(summary.get("matrix"), f"{path}.matrix")
    expected_matrix = {
        "reports": 12,
        "retained_samples": 360,
        "samples_per_report": SAMPLES,
        "warmups_per_report": 3,
        "cpu": 2,
        "workers": 1,
        "phases": ["R1", "R2"],
    }
    for key, expected in expected_matrix.items():
        if matrix.get(key) != expected:
            fail(f"{path}.matrix.{key}", f"expected {expected!r}")
    expected_selector = "odp_existing_append_lifecycle" if role == "control" else "odp_source_tail_append_lifecycle"
    if matrix.get("selector") != expected_selector:
        fail(f"{path}.matrix.selector", f"expected {expected_selector}")
    selector = text(summary.get("scope"), f"{path}.scope")
    if role == "control":
        if "existing-append" not in selector or "no" not in selector.lower() or "speedup" not in selector.lower():
            fail(f"{path}.scope", "does not withhold control speedup claims")
    else:
        if "source-backed" not in selector or "specialized" not in selector:
            fail(f"{path}.scope", "does not state the specialized source-backed candidate contract")
        contract = obj(summary.get("comparison_contract"), f"{path}.comparison_contract")
        if contract.get("status") != "withheld":
            fail(f"{path}.comparison_contract.status", "candidate summary must withhold comparison claims")
    protocol_sha = text(summary.get("protocol_sha256"), f"{path}.protocol_sha256")
    if len(protocol_sha) != 64 or any(character not in "0123456789abcdef" for character in protocol_sha):
        fail(f"{path}.protocol_sha256", "must be a lowercase SHA-256 digest")
    rows = summary.get("rows")
    if not isinstance(rows, list) or len(rows) != 12:
        fail(f"{path}.rows", "expected twelve authenticated rows")
    row_map: dict[tuple[str, str, str], dict[str, Any]] = {}
    for index, raw_row in enumerate(rows):
        row = obj(raw_row, f"{path}.rows[{index}]")
        repeat = text(row.get("repeat"), f"{path}.rows[{index}].repeat")
        mode = text(row.get("mode"), f"{path}.rows[{index}].mode")
        shape = text(row.get("shape"), f"{path}.rows[{index}].shape")
        if repeat not in REPEATS or mode not in MODES or shape not in SHAPES:
            fail(f"{path}.rows[{index}]", "row dimensions differ from the frozen matrix")
        key = (repeat, mode, shape)
        if key in row_map:
            fail(f"{path}.rows[{index}]", "duplicate repeat/mode/shape row")
        text(row.get("name"), f"{path}.rows[{index}].name")
        relative_path(row.get("receipt"), f"{path}.rows[{index}].receipt")
        relative_path(row.get("report"), f"{path}.rows[{index}].report")
        gnu_rss = integer(row.get("gnu_time_max_rss_bytes"), f"{path}.rows[{index}].gnu_time_max_rss_bytes", minimum=0)
        row["gnu_time_max_rss_bytes"] = gnu_rss
        label = f"{path}.rows[{index}]"
        for field in VECTOR_FIELDS:
            values = vector(row, field, label)
            if field in ("elapsed_ns", *PROCESS_FIELDS) and values is None:
                fail(f"{label}.{field}", "timing/process vector is required")
            if mode == "allocator" and field in ALLOCATION_FIELDS and values is None:
                fail(f"{label}.{field}", "allocator vector is required in allocator mode")
            if mode == "normal" and field in ALLOCATION_FIELDS and values is not None:
                fail(f"{label}.{field}", "normal rows must keep allocator metrics unavailable")
        row_map[key] = row
    expected_keys = {(repeat, mode, shape) for repeat in REPEATS for mode in MODES for shape in SHAPES}
    if set(row_map) != expected_keys:
        fail(str(path), "row dimensions do not cover the complete matrix")
    return summary, row_map


def endpoint(control: int | float, candidate: int | float) -> dict[str, Any]:
    control_value = float(control)
    candidate_value = float(candidate)
    delta = candidate_value - control_value
    if control_value == 0:
        relative: float | None = None
        flag = "baseline_zero"
    else:
        relative = delta / abs(control_value) * 100.0
        flag = "over_5_percent" if abs(relative) > THRESHOLD_PERCENT else "within_5_percent"
    return {
        "control": clean(control_value),
        "candidate": clean(candidate_value),
        "delta": clean(delta),
        "relative_percent": None if relative is None else clean(relative),
        "flag": flag,
    }


def metric_comparison(control_row: dict[str, Any], candidate_row: dict[str, Any], field: str, label: str) -> dict[str, Any] | None:
    control_stats = obj(control_row.get("statistics"), f"{label}.control.statistics").get(field)
    candidate_stats = obj(candidate_row.get("statistics"), f"{label}.candidate.statistics").get(field)
    if control_stats is None or candidate_stats is None:
        if control_stats is not None or candidate_stats is not None:
            fail(label, "control and candidate availability differs")
        return None
    control_stats = obj(control_stats, f"{label}.control.statistics.{field}")
    candidate_stats = obj(candidate_stats, f"{label}.candidate.statistics.{field}")
    return {
        quantile: endpoint(number(control_stats.get(quantile), f"{label}.control.{field}.{quantile}"), number(candidate_stats.get(quantile), f"{label}.candidate.{field}.{quantile}"))
        for quantile in QUANTILES
    }


def fnv_seed(label: str) -> int:
    state = 14_695_981_039_346_656_037
    for byte in label.encode("utf-8"):
        state ^= byte
        state = (state * 1_099_511_628_211) & UINT64_MASK
    return (state ^ BOOTSTRAP_SEED) & UINT64_MASK


def next_index(state: int, size: int) -> tuple[int, int]:
    state = (state * LCG_MULTIPLIER + LCG_INCREMENT) & UINT64_MASK
    return state, state % size


def resampled_median(values: list[int | float], state: int) -> tuple[int, float]:
    sample: list[float] = []
    for _ in range(SAMPLES):
        state, index = next_index(state, len(values))
        sample.append(float(values[index]))
    sample.sort()
    return state, (sample[(SAMPLES - 1) // 2] + sample[SAMPLES // 2]) / 2.0


def bootstrap_median_delta(control_values: list[int | float], candidate_values: list[int | float], label: str) -> dict[str, Any]:
    if len(control_values) != SAMPLES or len(candidate_values) != SAMPLES:
        fail(label, "bootstrap requires one thirty-sample vector on each side")
    point_control = statistics(control_values, f"{label}.control")['p50']
    point_candidate = statistics(candidate_values, f"{label}.candidate")['p50']
    lane_seed = fnv_seed(label)
    state = lane_seed
    deltas: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        state, control_median = resampled_median(control_values, state)
        state, candidate_median = resampled_median(candidate_values, state)
        deltas.append(candidate_median - control_median)
    ordered = sorted(deltas)
    lower = ordered[max(0, math.ceil(0.025 * BOOTSTRAP_RESAMPLES) - 1)]
    upper = ordered[max(0, math.ceil(0.975 * BOOTSTRAP_RESAMPLES) - 1)]
    return {
        "estimate_delta_ns": clean(float(point_candidate) - float(point_control)),
        "lower_95_ns": clean(lower),
        "upper_95_ns": clean(upper),
        "confidence": 0.95,
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
        "lane_seed": lane_seed,
        "method": "independent nonparametric bootstrap of candidate and control median vectors; nearest-rank 2.5th/97.5th endpoints",
    }


def compare(control_path: Path, candidate_path: Path) -> dict[str, Any]:
    control_summary, control_rows = validate_summary(control_path, "control")
    candidate_summary, candidate_rows = validate_summary(candidate_path, "candidate")
    rows: list[dict[str, Any]] = []
    for repeat in REPEATS:
        for mode in MODES:
            for shape in SHAPES:
                key = (repeat, mode, shape)
                control_row = control_rows[key]
                candidate_row = candidate_rows[key]
                label = f"{repeat}/{mode}/{shape}"
                metrics = {
                    "elapsed_ns": metric_comparison(control_row, candidate_row, "elapsed_ns", label),
                    "peak_rss_bytes": metric_comparison(control_row, candidate_row, "peak_rss_bytes", label),
                    "rss_delta_bytes": metric_comparison(control_row, candidate_row, "rss_delta_bytes", label),
                }
                allocation = {field: metric_comparison(control_row, candidate_row, field, label) for field in ALLOCATION_FIELDS}
                control_elapsed = vector(control_row, "elapsed_ns", label + ".control")
                candidate_elapsed = vector(candidate_row, "elapsed_ns", label + ".candidate")
                if control_elapsed is None or candidate_elapsed is None:
                    fail(label, "elapsed vectors disappeared after validation")
                rows.append({
                    "repeat": repeat,
                    "mode": mode,
                    "shape": shape,
                    "contracts": {
                        "control_selector": "odp_existing_append_lifecycle",
                        "candidate_selector": "odp_source_tail_append_lifecycle",
                        "relationship": "distinct retained-result contracts; numeric endpoint comparison only",
                    },
                    "provenance": {
                        "control_receipt": control_row["receipt"],
                        "control_report": control_row["report"],
                        "candidate_receipt": candidate_row["receipt"],
                        "candidate_report": candidate_row["report"],
                    },
                    "elapsed": metrics["elapsed_ns"],
                    "process_rss": {
                        "peak_rss_bytes": metrics["peak_rss_bytes"],
                        "rss_delta_bytes": metrics["rss_delta_bytes"],
                    },
                    "gnu_time_max_rss_bytes": endpoint(control_row["gnu_time_max_rss_bytes"], candidate_row["gnu_time_max_rss_bytes"]),
                    "allocator": {
                        "status": "measured" if mode == "allocator" else "unavailable",
                        "metrics": allocation,
                    },
                    "bootstrap_median_elapsed_delta": bootstrap_median_delta(control_elapsed, candidate_elapsed, label),
                })
    return {
        "schema": "litchi-0457-odp-control-candidate-comparison-v1",
        "change": CHANGE,
        "inputs": {
            "control": {
                "path": control_path.as_posix(),
                "schema": control_summary["schema"],
                "selector": control_summary["matrix"]["selector"],
                "protocol_sha256": control_summary["protocol_sha256"],
                "authentication": "derived summary whose receipts/reports were checked by control/derive.py",
            },
            "candidate": {
                "path": candidate_path.as_posix(),
                "schema": candidate_summary["schema"],
                "selector": candidate_summary["matrix"]["selector"],
                "protocol_sha256": candidate_summary["protocol_sha256"],
                "authentication": "derived summary whose receipts/reports were checked by candidate/verify.py and candidate/derive.py",
            },
        },
        "matrix": {
            "rows": len(rows),
            "repeats": list(REPEATS),
            "modes": list(MODES),
            "shapes": list(SHAPES),
            "samples_per_row": SAMPLES,
        },
        "comparison_contract": {
            "status": "numeric_endpoint_comparison_only",
            "control": "odp_existing_append_lifecycle",
            "candidate": "odp_source_tail_append_lifecycle",
            "delta": "candidate minus control",
            "relative_percent": "100 * delta / abs(control); baseline-zero endpoints are explicitly flagged",
            "threshold_percent": THRESHOLD_PERCENT,
            "ordinary_commit_patch_speedup": "withheld",
            "general_crud_or_retained_result_equivalence": "withheld",
            "geometric_mean": "not computed",
        },
        "bootstrap_definition": {
            "resamples": BOOTSTRAP_RESAMPLES,
            "seed": BOOTSTRAP_SEED,
            "confidence": 0.95,
            "unit": "nanoseconds",
            "scope": "median elapsed delta per repeat/mode/shape row",
            "method": "independent nonparametric resampling of the two retained thirty-sample vectors; deterministic fixed LCG; nearest-rank interval endpoints",
        },
        "rows": rows,
        "claims": [
            "p50, p95, and p99 are recomputed descriptive endpoints from each authenticated retained vector",
            "allocation volume and regional peak metrics are compared only where both allocator lanes publish measured vectors",
            "normal allocator metrics remain unavailable and are never replaced with zero",
            "process peak/RSS delta and GNU-time maximum RSS deltas carry explicit over-5-percent or baseline-zero flags",
            "bootstrap intervals describe median elapsed deltas within each matched matrix row",
            "no ordinary Commit/Patch speedup, regression, geomean, causal, scaling, or retained-result equivalence claim is made",
        ],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--control", type=Path, default=ROOT / "control" / "measurements.json")
    parser.add_argument("--candidate", type=Path, default=ROOT / "candidate" / "summary.json")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--write", action="store_true", help="write the comparison JSON; refuses to overwrite an existing file")
    args = parser.parse_args(argv)
    try:
        result = compare(args.control.resolve(), args.candidate.resolve())
        raw = json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False).encode("utf-8") + b"\n"
        if args.write:
            output = (args.output or ROOT / "comparison.json").resolve()
            if output.exists() or output.is_symlink():
                fail("output", "refusing to overwrite an existing comparison")
            output.write_bytes(raw)
        print(raw.decode("utf-8"))
        return 0
    except (OSError, KeyError, TypeError, ValueError, ComparisonError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
