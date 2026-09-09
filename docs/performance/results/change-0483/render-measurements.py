#!/usr/bin/env python3
"""Render an independent measurement review for the 0483 capture matrix.

The formal analyzer writes ``summary.json`` from the same raw reports.  This
renderer deliberately reads the protocol, reports, and GNU ``time`` receipts
itself so that the review can compare a second derivation with that summary.
It does not run the harness or any profiler.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
from pathlib import Path
from typing import Any, Iterable, Mapping


ROOT = Path(__file__).resolve().parent
COUNTS = (64, 8_192, 131_072)
ARMS = ("a1", "b1", "b2", "a2")
INSTRUMENTATIONS = ("normal", "allocator")
ROUTES = ("materialized", "bounded")
ROUTE_NAMES = {
    "materialized": "materialized_paragraph_copy",
    "bounded": "bounded_plain_text_tail_append",
}
SAMPLES = 30
THRESHOLD_PERCENT = 5.0
STAT_NAMES = (
    "n", "min", "max", "mean", "stdev", "ci95_low", "ci95_high", "p50", "p95", "p99",
)
PROCESS_FIELDS = (
    "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
    "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
    "system_cpu_ticks", "clock_ticks_per_second", "voluntary_context_switches",
    "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes",
)
READ_FIELDS = ("calls", "requested_bytes", "returned_bytes")
SINK_FIELDS = ("accepted_bytes", "write_calls", "largest_write")
ALLOC_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
HISTOGRAM_FIELDS = (
    "bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
    "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536",
)


class ReviewError(ValueError):
    """Raw evidence is incomplete or malformed."""


def fail(message: str) -> None:
    raise ReviewError(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot read {path}: {error}")


def number(value: Any, label: str) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(f"{label}: expected a number")
    if isinstance(value, float) and not math.isfinite(value):
        fail(f"{label}: expected a finite number")
    if value < 0:
        fail(f"{label}: expected a non-negative number")
    return value


def integer(value: Any, label: str) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or value < 0:
        fail(f"{label}: expected a non-negative integer")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected a non-empty string")
    return value


def capture_specs(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    captures = protocol.get("captures")
    if not isinstance(captures, list) or len(captures) != 24:
        fail("protocol must contain exactly 24 formal captures")
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for index, raw in enumerate(captures):
        if not isinstance(raw, dict):
            fail(f"protocol.captures[{index}] is malformed")
        spec = {
            key: raw.get(key)
            for key in ("label", "arm", "repeat", "route", "route_name", "instrumentation", "count")
        }
        label = text(spec["label"], f"capture[{index}].label")
        if label in seen:
            fail(f"duplicate capture label: {label}")
        seen.add(label)
        if spec["arm"] not in ARMS or spec["route"] not in ROUTES:
            fail(f"{label}: unknown arm or route")
        if spec["instrumentation"] not in INSTRUMENTATIONS:
            fail(f"{label}: unknown instrumentation")
        if spec["repeat"] not in (1, 2) or spec["count"] not in COUNTS:
            fail(f"{label}: unknown repeat or count")
        expected_route_name = ROUTE_NAMES[spec["route"]]
        if spec["route_name"] != expected_route_name:
            fail(f"{label}: route name differs from route")
        result.append(spec)
    expected: dict[str, dict[str, Any]] = {}
    for arm, route, repeat in (
        ("a1", "materialized", 1), ("b1", "bounded", 1),
        ("b2", "bounded", 2), ("a2", "materialized", 2),
    ):
        for instrumentation in INSTRUMENTATIONS:
            for count in COUNTS:
                label = f"{arm}-{instrumentation}-{count}-{route}"
                expected[label] = {
                    "arm": arm,
                    "repeat": repeat,
                    "route": route,
                    "route_name": ROUTE_NAMES[route],
                    "instrumentation": instrumentation,
                    "count": count,
                }
    if seen != set(expected):
        fail("protocol capture label matrix differs from the 24-case design")
    for spec in result:
        expected_spec = expected[spec["label"]]
        if any(spec[key] != value for key, value in expected_spec.items()):
            fail(f"{spec['label']}: protocol topology differs from the 24-case design")
    return sorted(result, key=lambda item: item["label"])


def histogram(value: Any, label: str) -> dict[str, int]:
    if not isinstance(value, dict) or set(value) != set(HISTOGRAM_FIELDS):
        fail(f"{label}: histogram fields differ")
    return {key: integer(value[key], f"{label}.{key}") for key in HISTOGRAM_FIELDS}


def read_counters(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: source-read counters are missing")
    result = {field: integer(value.get(field), f"{label}.{field}") for field in READ_FIELDS}
    result["request_histogram"] = histogram(value.get("request_histogram"), f"{label}.request_histogram")
    if sum(result["request_histogram"].values()) != result["calls"]:
        fail(f"{label}: request histogram does not sum to calls")
    if result["returned_bytes"] > result["requested_bytes"]:
        fail(f"{label}: returned bytes exceed requested bytes")
    return result


def sink_counters(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: sink counters are missing")
    result = {field: integer(value.get(field), f"{label}.{field}") for field in SINK_FIELDS}
    result["histogram"] = histogram(value.get("histogram"), f"{label}.histogram")
    if sum(result["histogram"].values()) != result["write_calls"]:
        fail(f"{label}: write histogram does not sum to calls")
    if result["largest_write"] > result["accepted_bytes"]:
        fail(f"{label}: largest write exceeds accepted bytes")
    result["sha256"] = text(value.get("sha256"), f"{label}.sha256")
    return result


def allocation_counters(value: Any, instrumentation: str, label: str) -> dict[str, Any] | None:
    if instrumentation == "normal":
        if value not in (None, {"status": "unavailable", "scope": "operation_global_system_allocator"}):
            fail(f"{label}: normal lane exposes allocator counters")
        return None
    if not isinstance(value, dict):
        fail(f"{label}: allocator counters are missing")
    result = {field: integer(value.get(field), f"{label}.{field}") for field in ALLOC_FIELDS}
    if value.get("status") != "measured":
        fail(f"{label}: allocator status is not measured")
    if result["live_bytes_before"] + result["allocated_bytes"] - result["deallocated_bytes"] != result["live_bytes_after"]:
        fail(f"{label}: allocator conservation fails")
    if result["peak_live_bytes_before"] > result["peak_live_bytes_after"]:
        fail(f"{label}: absolute peak regresses")
    if result["region_peak_live_bytes"] < max(result["live_bytes_before"], result["live_bytes_after"]):
        fail(f"{label}: region peak is below a region boundary")
    if result["region_peak_live_bytes"] > result["peak_live_bytes_after"]:
        fail(f"{label}: region peak exceeds absolute peak")
    result["incremental_peak_live_bytes"] = result["region_peak_live_bytes"] - result["live_bytes_before"]
    result["retention_delta_bytes"] = result["live_bytes_after"] - result["live_bytes_before"]
    return result


def process_counters(value: Any, label: str) -> dict[str, int] | None:
    if value is None:
        return None
    if not isinstance(value, dict):
        fail(f"{label}: process counters are malformed")
    result = {field: integer(value.get(field), f"{label}.{field}") for field in PROCESS_FIELDS}
    if result["peak_rss_bytes"] < result["rss_bytes"]:
        fail(f"{label}: absolute VmHWM is below RSS delta")
    return result


def resource_rss(path: Path, label: str) -> int:
    if not path.is_file():
        fail(f"{label}: GNU time resource receipt is missing")
    prefix = "Maximum resident set size (kbytes):"
    values: list[int] = []
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line.startswith(prefix):
            continue
        value = line[len(prefix):].strip()
        try:
            parsed = int(value)
        except ValueError:
            fail(f"{label}: malformed GNU time RSS value")
        if parsed < 0:
            fail(f"{label}: GNU time RSS value is negative")
        values.append(parsed)
    if len(values) != 1:
        fail(f"{label}: expected exactly one GNU time RSS value")
    return values[0] * 1024


def load_capture(root: Path, spec: Mapping[str, Any]) -> dict[str, Any]:
    label = spec["label"]
    path = root / "captures" / f"{label}.report.json"
    report = read_json(path)
    if not isinstance(report, dict) or report.get("version") != 1:
        fail(f"{label}: report schema/version is malformed")
    config = report.get("config")
    if not isinstance(config, dict) or config.get("samples") != SAMPLES:
        fail(f"{label}: report sample configuration differs")
    cases = report.get("cases")
    if not isinstance(cases, list) or len(cases) != 1:
        fail(f"{label}: expected one report case")
    case = cases[0]
    if not isinstance(case, dict) or case.get("count") != spec["count"]:
        fail(f"{label}: report count differs")
    corpus = case.get("corpus")
    routes = case.get("routes")
    if not isinstance(corpus, dict) or not isinstance(routes, list) or len(routes) != 1:
        fail(f"{label}: report corpus/route shape is malformed")
    route_report = routes[0]
    if not isinstance(route_report, dict) or route_report.get("route") != spec["route_name"]:
        fail(f"{label}: report route differs")
    route_corpus = route_report.get("corpus")
    samples = route_report.get("samples")
    if not isinstance(route_corpus, dict) or not isinstance(samples, list) or len(samples) != SAMPLES:
        fail(f"{label}: report route samples are malformed")
    checked: list[dict[str, Any]] = []
    for index, raw in enumerate(samples):
        if not isinstance(raw, dict) or raw.get("sample") != index or raw.get("route") != spec["route_name"]:
            fail(f"{label}.samples[{index}]: sample identity differs")
        elapsed = integer(raw.get("elapsed_ns"), f"{label}.samples[{index}].elapsed_ns")
        if elapsed == 0:
            fail(f"{label}.samples[{index}]: elapsed time is zero")
        checked.append({
            "elapsed_ns": elapsed,
            "source_reads": read_counters(raw.get("source_reads"), f"{label}.samples[{index}].source_reads"),
            "sink": sink_counters(raw.get("sink"), f"{label}.samples[{index}].sink"),
            "allocation": allocation_counters(raw.get("allocation"), spec["instrumentation"], f"{label}.samples[{index}].allocation"),
            "process": process_counters(raw.get("process"), f"{label}.samples[{index}].process"),
        })
    return {
        "spec": dict(spec),
        "report_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "source": corpus,
        "route": route_corpus,
        "samples": checked,
        "gnu_time_max_rss_bytes": resource_rss(root / "captures" / f"{label}.resource", label),
    }


def percentile(values: Iterable[int | float], fraction: float) -> float:
    ordered = sorted(float(value) for value in values)
    if not ordered:
        fail("cannot calculate a percentile over an empty vector")
    if len(ordered) == 1:
        return ordered[0]
    position = (len(ordered) - 1) * fraction
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return ordered[lower]
    ratio = position - lower
    return ordered[lower] + (ordered[upper] - ordered[lower]) * ratio


def stats(values: Iterable[int | float]) -> dict[str, int | float]:
    values = list(values)
    if not values:
        fail("cannot calculate statistics over an empty vector")
    mean = statistics.fmean(values)
    stdev = statistics.stdev(values) if len(values) > 1 else 0.0
    sem = stdev / math.sqrt(len(values))
    return {
        "n": len(values),
        "min": min(values),
        "max": max(values),
        "mean": mean,
        "stdev": stdev,
        "ci95_low": mean - 1.96 * sem,
        "ci95_high": mean + 1.96 * sem,
        "p50": percentile(values, 0.50),
        "p95": percentile(values, 0.95),
        "p99": percentile(values, 0.99),
    }


def sample_metric(sample: Mapping[str, Any], metric: str) -> int | float:
    if metric == "elapsed_ns":
        return sample["elapsed_ns"]
    if metric == "output_throughput_bytes_per_second":
        return sample["sink"]["accepted_bytes"] * 1_000_000_000 / sample["elapsed_ns"]
    if metric == "source_requested_throughput_bytes_per_second":
        return sample["source_reads"]["requested_bytes"] * 1_000_000_000 / sample["elapsed_ns"]
    if metric.startswith("source_reads."):
        return sample["source_reads"][metric.split(".", 1)[1]]
    if metric.startswith("sink."):
        return sample["sink"][metric.split(".", 1)[1]]
    if metric.startswith("process."):
        process = sample.get("process")
        if process is None:
            fail(f"{metric}: process observer is unavailable")
        return process[metric.split(".", 1)[1]]
    if metric.startswith("allocation."):
        allocation = sample.get("allocation")
        if allocation is None:
            fail(f"{metric}: allocator observer is unavailable")
        return allocation[metric.split(".", 1)[1]]
    fail(f"unknown raw metric: {metric}")


def row_metrics(capture: Mapping[str, Any]) -> dict[str, Any]:
    spec = capture["spec"]
    samples = capture["samples"]
    metrics = {
        "elapsed_ns": stats(sample_metric(sample, "elapsed_ns") for sample in samples),
        "output_throughput_bytes_per_second": stats(sample_metric(sample, "output_throughput_bytes_per_second") for sample in samples),
        "source_requested_throughput_bytes_per_second": stats(sample_metric(sample, "source_requested_throughput_bytes_per_second") for sample in samples),
    }
    for field in READ_FIELDS:
        metrics[f"source_reads.{field}"] = stats(sample_metric(sample, f"source_reads.{field}") for sample in samples)
    for field in SINK_FIELDS:
        metrics[f"sink.{field}"] = stats(sample_metric(sample, f"sink.{field}") for sample in samples)
    process_available = all(sample["process"] is not None for sample in samples)
    if process_available:
        for field in PROCESS_FIELDS:
            metrics[f"process.{field}"] = stats(sample_metric(sample, f"process.{field}") for sample in samples)
    if spec["instrumentation"] == "allocator":
        for field in (*ALLOC_FIELDS, "incremental_peak_live_bytes", "retention_delta_bytes"):
            metrics[f"allocation.{field}"] = stats(sample_metric(sample, f"allocation.{field}") for sample in samples)
        metrics["total_peak_live_bytes"] = metrics["allocation.region_peak_live_bytes"]
    first = samples[0]
    return {
        "label": spec["label"],
        "arm": spec["arm"],
        "repeat": spec["repeat"],
        "route": spec["route"],
        "route_name": spec["route_name"],
        "instrumentation": spec["instrumentation"],
        "count": spec["count"],
        "report_sha256": capture["report_sha256"],
        "source_archive_sha256": capture["source"].get("source_archive_sha256"),
        "source_main_xml_sha256": capture["source"].get("source_main_xml_sha256"),
        "source_archive_bytes": capture["source"].get("source_archive_bytes"),
        "source_main_xml_bytes": capture["source"].get("source_main_xml_bytes"),
        "output_archive_bytes": capture["route"].get("output_archive_bytes"),
        "output_archive_sha256": capture["route"].get("output_archive_sha256"),
        "output_main_xml_bytes": capture["route"].get("output_main_xml_bytes"),
        "output_main_xml_sha256": capture["route"].get("output_main_xml_sha256"),
        "gnu_time_max_rss_bytes": capture["gnu_time_max_rss_bytes"],
        "metrics": metrics,
        "process_observer": process_available,
        "request_histogram": first["source_reads"]["request_histogram"],
        "write_histogram": first["sink"]["histogram"],
    }


def metric_mean(row: Mapping[str, Any], metric: str) -> float:
    value = row["metrics"].get(metric)
    if not isinstance(value, Mapping):
        fail(f"{row['label']}: metric missing: {metric}")
    result = value.get("mean")
    if not isinstance(result, (int, float)) or isinstance(result, bool):
        fail(f"{row['label']}: metric mean is not numeric: {metric}")
    return float(result)


def percent_change(before: float, after: float) -> tuple[float | None, bool, str | None]:
    if before == 0:
        if after > 0:
            # A percentage relative to zero is undefined. Keep the explicit
            # status so a consumer cannot mistake it for a measured percent.
            return None, True, "undefined_zero_baseline"
        # Preserve the analyzer's established 0 -> 0 representation. There
        # is no drift to flag, and no undefined status is needed.
        return 0.0, True, None
    return (after - before) / before * 100.0, False, None


def compare_rows(first: Mapping[str, Any], second: Mapping[str, Any], kind: str) -> dict[str, Any]:
    metrics = sorted(set(first["metrics"]) & set(second["metrics"]))
    changes = []
    for metric in metrics:
        before = metric_mean(first, metric)
        after = metric_mean(second, metric)
        change, zero_baseline, percent_status = percent_change(before, after)
        change_record = {
            "metric": metric,
            "first_mean": before,
            "second_mean": after,
            "percent": change,
            "zero_baseline": zero_baseline,
            "flagged": change is None or abs(change) > THRESHOLD_PERCENT,
            "adverse": after < before if metric == "output_throughput_bytes_per_second" else after > before,
        }
        if percent_status is not None:
            change_record["percent_status"] = percent_status
        changes.append(change_record)
    first_rss = float(first["gnu_time_max_rss_bytes"])
    second_rss = float(second["gnu_time_max_rss_bytes"])
    change, zero_baseline, percent_status = percent_change(first_rss, second_rss)
    change_record = {
        "metric": "gnu_time_max_rss_bytes",
        "first_mean": first_rss,
        "second_mean": second_rss,
        "percent": change,
        "zero_baseline": zero_baseline,
        "flagged": change is None or abs(change) > THRESHOLD_PERCENT,
        "adverse": second_rss > first_rss,
    }
    if percent_status is not None:
        change_record["percent_status"] = percent_status
    changes.append(change_record)
    return {
        "kind": kind,
        "first": first["label"],
        "second": second["label"],
        "instrumentation": first["instrumentation"],
        "route": first["route"],
        "count": first["count"],
        "changes": changes,
        "flags": [change for change in changes if change["flagged"]],
    }


def comparisons(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    by_key = {(row["arm"], row["instrumentation"], row["count"]): row for row in rows}
    route_comparisons: list[dict[str, Any]] = []
    repeat_comparisons: list[dict[str, Any]] = []
    for instrumentation in INSTRUMENTATIONS:
        for count in COUNTS:
            for first_arm, second_arm, route, repeat in (
                ("a1", "b1", "candidate_b1_vs_control_a1", 1),
                ("a2", "b2", "candidate_b2_vs_control_a2", 2),
            ):
                first = by_key[(first_arm, instrumentation, count)]
                second = by_key[(second_arm, instrumentation, count)]
                route_comparisons.append(compare_rows(first, second, route))
            for first_arm, second_arm, route in (
                ("a1", "a2", "materialized_repeat_a2_vs_a1"),
                ("b1", "b2", "bounded_repeat_b2_vs_b1"),
            ):
                first = by_key[(first_arm, instrumentation, count)]
                second = by_key[(second_arm, instrumentation, count)]
                repeat_comparisons.append(compare_rows(first, second, route))
    return route_comparisons, repeat_comparisons


def corpus_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for count in COUNTS:
        selected = [row for row in rows if row["count"] == count]
        if len(selected) != 8:
            fail(f"count {count}: expected eight report rows")
        for field in (
            "source_archive_sha256", "source_main_xml_sha256", "source_archive_bytes", "source_main_xml_bytes",
        ):
            if len({row[field] for row in selected}) != 1:
                fail(f"count {count}: {field} differs between reports")
        route_records = []
        for route in ROUTES:
            route_selected = [row for row in selected if row["route"] == route]
            for field in ("output_archive_sha256", "output_archive_bytes", "output_main_xml_sha256", "output_main_xml_bytes"):
                if len({row[field] for row in route_selected}) != 1:
                    fail(f"count {count}/{route}: {field} differs between reports")
            route_records.append({
                "route": route,
                "route_name": ROUTE_NAMES[route],
                **{field: route_selected[0][field] for field in ("output_archive_sha256", "output_archive_bytes", "output_main_xml_sha256", "output_main_xml_bytes")},
            })
        result.append({
            "count": count,
            **{field: selected[0][field] for field in ("source_archive_sha256", "source_archive_bytes", "source_main_xml_sha256", "source_main_xml_bytes")},
            "routes": route_records,
        })
    return result


def summary_path(metric: str) -> tuple[str, ...] | None:
    """Map an independent metric name to the analyzer's summary shape."""
    if metric in {
        "source_requested_throughput_bytes_per_second",
        "allocation.incremental_peak_live_bytes",
        "allocation.retention_delta_bytes",
    }:
        return None
    if metric == "output_throughput_bytes_per_second":
        return ("throughput_bytes_per_second",)
    if metric == "gnu_time_max_rss_bytes":
        return ("rss_bytes",)
    if metric == "total_peak_live_bytes":
        return ("total_peak_live_bytes",)
    return tuple(metric.split("."))


def summary_value(row: Mapping[str, Any], path: tuple[str, ...]) -> Any:
    value: Any = row
    for part in path:
        if not isinstance(value, Mapping) or part not in value:
            return None
        value = value[part]
    return value


def same_number(first: Any, second: Any) -> bool:
    if isinstance(first, bool) or isinstance(second, bool):
        return first == second
    if isinstance(first, (int, float)) and isinstance(second, (int, float)):
        if isinstance(first, float) and not math.isfinite(first):
            return False
        if isinstance(second, float) and not math.isfinite(second):
            return False
        return math.isclose(float(first), float(second), rel_tol=1e-12, abs_tol=1e-9)
    return first == second


def compare_stat_vector(
    differences: list[dict[str, Any]],
    label: str,
    metric: str,
    raw_stats: Mapping[str, Any],
    summary_stats: Any,
) -> None:
    if not isinstance(summary_stats, Mapping):
        differences.append({"label": label, "field": metric, "raw": raw_stats, "summary": summary_stats})
        return
    for stat_name in STAT_NAMES:
        raw_value = raw_stats.get(stat_name)
        summary_value_item = summary_stats.get(stat_name)
        if not same_number(raw_value, summary_value_item):
            differences.append({
                "label": label,
                "field": f"{metric}.{stat_name}",
                "raw": raw_value,
                "summary": summary_value_item,
            })


def canonical_summary_metric(metric: str) -> str:
    if metric in {
        "source_requested_throughput_bytes_per_second",
        "allocation.incremental_peak_live_bytes",
        "allocation.retention_delta_bytes",
    }:
        return ""
    if metric == "output_throughput_bytes_per_second":
        return "throughput_bytes_per_second"
    if metric == "gnu_time_max_rss_bytes":
        return "rss_bytes"
    return metric


def compare_summary_flags(
    summary: Mapping[str, Any],
    route_comparisons: list[dict[str, Any]],
    repeat_comparisons: list[dict[str, Any]],
    differences: list[dict[str, Any]],
) -> bool:
    summary_comparisons = summary.get("comparisons")
    if not isinstance(summary_comparisons, list):
        differences.append({"field": "comparisons", "raw": "present", "summary": summary_comparisons})
        return False
    expected_records = route_comparisons + repeat_comparisons
    by_key: dict[tuple[Any, Any, Any], dict[str, Any]] = {}
    for record in summary_comparisons:
        if not isinstance(record, Mapping):
            continue
        key = (record.get("name"), record.get("instrumentation"), record.get("count"))
        if key in by_key:
            differences.append({"field": "comparisons.duplicate", "summary": key})
        by_key[key] = dict(record)
    matched = True
    for expected in expected_records:
        key = (expected["kind"], expected["instrumentation"], expected["count"])
        actual = by_key.get(key)
        if actual is None:
            differences.append({"field": "comparisons.row", "raw": key, "summary": "missing"})
            matched = False
            continue
        raw_flag_map = {
            summary_metric: (flag.get("percent"), flag.get("adverse"))
            for flag in expected["flags"]
            for summary_metric in (canonical_summary_metric(flag["metric"]),)
            if summary_metric
        }
        summary_flags = actual.get("flags")
        if not isinstance(summary_flags, list):
            differences.append({"field": "comparisons.flags", "raw": raw_flag_map, "summary": summary_flags, "key": key})
            matched = False
            continue
        summary_flag_map = {
            canonical_summary_metric(flag.get("metric")): (flag.get("percent"), flag.get("adverse"))
            for flag in summary_flags
            if isinstance(flag, Mapping) and canonical_summary_metric(flag.get("metric"))
        }
        if set(raw_flag_map) != set(summary_flag_map):
            differences.append({"field": "comparisons.flags.metrics", "raw": sorted(raw_flag_map), "summary": sorted(summary_flag_map), "key": key})
            matched = False
        for metric in set(raw_flag_map) & set(summary_flag_map):
            raw_percent, raw_adverse = raw_flag_map[metric]
            summary_percent, summary_adverse = summary_flag_map[metric]
            if not same_number(raw_percent, summary_percent):
                differences.append({"field": f"comparisons.flags.{metric}", "raw": raw_percent, "summary": summary_percent, "key": key})
                matched = False
            if raw_adverse != summary_adverse:
                differences.append({"field": f"comparisons.flags.{metric}.adverse", "raw": raw_adverse, "summary": summary_adverse, "key": key})
                matched = False
    expected_keys = {
        (record["kind"], record["instrumentation"], record["count"])
        for record in expected_records
    }
    if set(by_key) != expected_keys:
        differences.append({"field": "comparisons.keys", "raw": sorted(expected_keys), "summary": sorted(by_key)})
        matched = False
    return matched


def summary_comparison(root: Path, rows: list[dict[str, Any]], summary: Any | None) -> dict[str, Any]:
    if summary is None:
        return {"available": False, "matches": None, "differences": [], "reason": "summary.json was not supplied"}
    if not isinstance(summary, dict):
        fail("summary.json is not an object")
    summary_rows = summary.get("rows")
    if not isinstance(summary_rows, list):
        fail("summary.json rows are missing")
    by_label = {row.get("label"): row for row in summary_rows if isinstance(row, dict)}
    differences: list[dict[str, Any]] = []
    compared = 0
    for row in rows:
        other = by_label.get(row["label"])
        if other is None:
            differences.append({"label": row["label"], "field": "row", "raw": "present", "summary": "missing"})
            continue
        compared += 1
        for field in (
            "source_archive_sha256", "source_main_xml_sha256", "gnu_time_max_rss_bytes",
        ):
            summary_field = "rss_bytes" if field == "gnu_time_max_rss_bytes" else field
            raw_value = row[field]
            summary_scalar = other.get(summary_field)
            if raw_value != summary_scalar:
                differences.append({"label": row["label"], "field": field, "raw": raw_value, "summary": summary_scalar})
        for metric in row["metrics"]:
            path = summary_path(metric)
            if path is None:
                continue
            compare_stat_vector(
                differences,
                row["label"],
                metric,
                row["metrics"][metric],
                summary_value(other, path),
            )
    unrepresented = sorted(
        metric
        for row in rows
        for metric in row["metrics"]
        if summary_path(metric) is None
    )
    unrepresented = sorted(set(unrepresented))
    # The comparison lists are checked below by derive(), after they have
    # been independently recomputed from the same rows.
    return {
        "available": True,
        "summary_schema": summary.get("schema"),
        "summary_rows": len(summary_rows),
        "compared_rows": compared,
        "unrepresented_raw_metrics": unrepresented,
        "matches": not differences and len(summary_rows) == len(rows),
        "differences": differences,
    }


def derive(root: Path, summary_path: Path | None) -> dict[str, Any]:
    protocol_path = root / "protocol.json"
    protocol = read_json(protocol_path)
    if not isinstance(protocol, dict):
        fail("protocol.json is not an object")
    specs = capture_specs(protocol)
    captures = [load_capture(root, spec) for spec in specs]
    rows = [row_metrics(capture) for capture in captures]
    route_comparison, repeat_comparison = comparisons(rows)
    summary = read_json(summary_path) if summary_path is not None and summary_path.is_file() else None
    summary_review = summary_comparison(root, rows, summary)
    if summary is not None:
        flags_match = compare_summary_flags(
            summary,
            route_comparison,
            repeat_comparison,
            summary_review["differences"],
        )
        summary_review["comparison_flags_match"] = flags_match
        summary_review["matches"] = summary_review["matches"] and flags_match
    else:
        summary_review["comparison_flags_match"] = None
    return {
        "schema": "docx-tail-append-measurement-review-v1",
        "protocol_sha256": hashlib.sha256(protocol_path.read_bytes()).hexdigest(),
        "summary_comparison": summary_review,
        "method": {
            "sample_count": SAMPLES,
            "capture_count": len(rows),
            "percentile": "linear interpolation at position (n - 1) * q",
            "confidence_interval": "normal approximate mean +/- 1.96 * sample standard deviation / sqrt(n)",
            "repeat_percent_change": "(repeat_2_mean - repeat_1_mean) / repeat_1_mean * 100",
            "flag_rule": "absolute change strictly greater than 5 percent; improvements are retained as flags",
            "elapsed_unit": "milliseconds in the rendered document; nanoseconds in review JSON",
            "rss_scope": "one /usr/bin/time -v maximum resident set size per fresh process, converted from KiB to bytes",
            "allocator_increment": "region_peak_live_bytes - live_bytes_before, computed independently per sample",
        },
        "matrix": {
            "counts": list(COUNTS),
            "arms": list(ARMS),
            "routes": list(ROUTES),
            "instrumentations": list(INSTRUMENTATIONS),
            "phase_samples": protocol.get("phase_samples"),
        },
        "corpus": corpus_rows(rows),
        "rows": rows,
        "route_comparisons": route_comparison,
        "repeat_comparisons": repeat_comparison,
        "caveats": [
            "The materialized and bounded route arms select two paths in the same executable for each instrumentation lane; normal and allocator lanes use separate instrumentation binaries and are not a direct speed comparison.",
            "Each timed lifecycle appends exactly one plain paragraph at the tail. The bounded caller text is borrowed from storage prepared outside the timed region.",
            "requested allocation bytes are callback accounting: reallocation increments allocation_calls and reallocation_calls, and allocated_bytes includes the full requested new_size. No physical copy-byte counter is present.",
            "region_peak_live_bytes is an absolute live-byte high-water value. The operation peak shown here subtracts live_bytes_before; independent attribution-region peaks must not be summed.",
            "Total elapsed time includes source/package admission, publication, sink digest finalization, and owner drops. Corpus construction and independent source/candidate oracles remain outside the timed lifecycle.",
            "process.rss_bytes is a saturating RSS delta and process.peak_rss_bytes is an absolute VmHWM endpoint. GNU time RSS is a broader whole-process maximum covering setup, oracles, serialization, and teardown.",
            "The explicit bounded-window memory goal remains open; this one-append matrix does not establish a constant-memory or arbitrary DOCX scaling property.",
        ],
    }


def fmt(value: Any, digits: int = 6) -> str:
    if isinstance(value, float):
        if math.isinf(value):
            return "inf" if value > 0 else "-inf"
        return f"{value:.{digits}f}"
    if isinstance(value, int):
        return f"{value:,}"
    return str(value)


def metric_display(metric: str) -> tuple[str, str, float]:
    if metric == "elapsed_ns":
        return "elapsed", "ms", 1e-6
    if metric.endswith("throughput_bytes_per_second"):
        return metric.replace("_bytes_per_second", ""), "bytes/s", 1.0
    if metric.endswith("_bytes") or ".bytes" in metric:
        return metric, "bytes", 1.0
    if metric in {
        "process.rchar", "process.wchar", "process.read_bytes", "process.write_bytes",
        "process.cancelled_write_bytes",
    }:
        return metric, "bytes", 1.0
    if metric.endswith("_calls") or metric.endswith(".calls"):
        return metric, "calls", 1.0
    if metric == "process.clock_ticks_per_second":
        return metric, "ticks/s", 1.0
    if metric.endswith("_ticks") or metric.endswith("_ticks_per_second"):
        return metric, "ticks", 1.0
    return metric, "events", 1.0


def stat_table(lines: list[str], rows: list[dict[str, Any]], metric_names: list[str], title: str, scale: float = 1.0) -> None:
    lines.extend([
        f"### {title}",
        "",
        "| lane | count | repeat | metric | mean | 95% CI | p50 | p95 | p99 |",
        "| :--- | ---: | ---: | :--- | ---: | :--- | ---: | ---: | ---: |",
    ])
    for row in rows:
        for metric in metric_names:
            item = row["metrics"].get(metric)
            if item is None:
                continue
            label, unit, _ = metric_display(metric)
            lines.append(
                f"| {row['instrumentation']}/{row['route']}/{row['arm']} | {row['count']:,} | {row['repeat']} | {label} ({unit}) | "
                f"{fmt(item['mean'] * scale)} | [{fmt(item['ci95_low'] * scale)}, {fmt(item['ci95_high'] * scale)}] | "
                f"{fmt(item['p50'] * scale)} | {fmt(item['p95'] * scale)} | {fmt(item['p99'] * scale)} |"
            )
    lines.append("")


def histogram_text(value: Mapping[str, int]) -> str:
    return ", ".join(f"{key}={value[key]}" for key in HISTOGRAM_FIELDS if value[key]) or "all zero"


def histogram_table(lines: list[str], rows: list[dict[str, Any]]) -> None:
    lines.extend([
        "### Read and write histogram shapes",
        "",
        "The histogram counters are per sample; the table shows the first sample for each lane and the full vectors are retained in the JSON review.",
        "",
        "| lane | count | repeat | source request histogram | sink write histogram |",
        "| :--- | ---: | ---: | :--- | :--- |",
    ])
    for row in rows:
        lines.append(
            f"| {row['instrumentation']}/{row['route']}/{row['arm']} | {row['count']:,} | {row['repeat']} | "
            f"{histogram_text(row['request_histogram'])} | {histogram_text(row['write_histogram'])} |"
        )
    lines.append("")


def comparison_lines(lines: list[str], comparisons_value: list[dict[str, Any]], title: str) -> None:
    lines.extend([
        f"### {title}",
        "",
        "Every metric is retained in `measurement-review.json`. The table lists every >5% review flag and every undefined zero-baseline increase; an explicit no-flag row is emitted for stable lanes.",
        "",
        "| lane | count | metric | change | adverse |",
        "| :--- | ---: | :--- | ---: | :---: |",
    ])
    emitted = False
    for comparison in comparisons_value:
        for change in comparison["flags"]:
            emitted = True
            change_text = (
                "undefined (zero baseline)"
                if change["percent"] is None
                else f"{fmt(change['percent'], 2)}%"
            )
            lines.append(
                f"| {comparison['instrumentation']}/{comparison.get('route', '')} ({comparison['first']} → {comparison['second']}) | "
                f"{comparison['count']:,} | {change['metric']} | {change_text} | {'yes' if change['adverse'] else 'no'} |"
            )
    if not emitted:
        lines.append("| all lanes | — | no metric exceeded 5% | 0.00% | no |")
    lines.append("")


def render_markdown(review: Mapping[str, Any]) -> str:
    rows = list(review["rows"])
    lines = [
        "# Change 0483 measurements",
        "",
        "This document independently recomputes the formal 0483 measurements from the 24 retained raw reports and their GNU `time -v` receipts (30 samples per process, 720 samples total). Percentiles use linear interpolation at `(n - 1) * q`; 95% intervals are the normal approximation `mean ± 1.96 × SEM`. The raw derivation and every comparison remain in [measurement-review.json](measurement-review.json).",
        "",
        "For each instrumentation lane, materialized and bounded are route selections in the same executable, using a fresh process per capture. Every timed lifecycle appends one plain paragraph at the tail. The normal and allocator lanes use separate binaries, so their elapsed values are reported as separate lanes.",
        "",
        "## Corpus and protocol",
        "",
        "| paragraphs | source XML bytes | source archive bytes | materialized output bytes | bounded output bytes |",
        "| ---: | ---: | ---: | ---: | ---: |",
    ]
    for corpus in review["corpus"]:
        outputs = {item["route"]: item for item in corpus["routes"]}
        lines.append(
            f"| {corpus['count']:,} | {fmt(corpus['source_main_xml_bytes'])} | {fmt(corpus['source_archive_bytes'])} | "
            f"{fmt(outputs['materialized']['output_archive_bytes'])} | {fmt(outputs['bounded']['output_archive_bytes'])} |"
        )
    lines.extend([
        "",
        "The report-level source and output hashes are retained in the JSON review and were required to be constant for each count and route across the eight corresponding reports. The protocol records `phase_samples: false`, so this matrix has no phase retention or phase I/O rows.",
        "",
        "## Total lifecycle",
        "",
        "`elapsed` is the timed total lifecycle. `output_throughput` uses sink accepted bytes; `source_requested_throughput` uses logical source requested bytes, so the two rates describe different byte domains.",
        "",
        "| lane | count | repeat | elapsed mean ms | 95% CI ms | p50 ms | p95 ms | p99 ms | GNU time max RSS KiB | process RSS delta mean | process VmHWM mean |",
        "| :--- | ---: | ---: | ---: | :--- | ---: | ---: | ---: | ---: | ---: | ---: |",
    ])
    for row in rows:
        elapsed = row["metrics"]["elapsed_ns"]
        process_rss = row["metrics"].get("process.rss_bytes")
        process_hwm = row["metrics"].get("process.peak_rss_bytes")
        lines.append(
            f"| {row['instrumentation']}/{row['route']}/{row['arm']} | {row['count']:,} | {row['repeat']} | "
            f"{fmt(elapsed['mean'] * 1e-6)} | [{fmt(elapsed['ci95_low'] * 1e-6)}, {fmt(elapsed['ci95_high'] * 1e-6)}] | "
            f"{fmt(elapsed['p50'] * 1e-6)} | {fmt(elapsed['p95'] * 1e-6)} | {fmt(elapsed['p99'] * 1e-6)} | "
            f"{fmt(row['gnu_time_max_rss_bytes'] / 1024, 0)} | "
            f"{fmt(process_rss['mean'] if process_rss else 'unavailable')} | "
            f"{fmt(process_hwm['mean'] if process_hwm else 'unavailable')} |"
        )
    lines.append("")
    stat_table(lines, rows, ["output_throughput_bytes_per_second", "source_requested_throughput_bytes_per_second"], "Throughput")
    stat_table(lines, rows, ["process.rss_bytes", "process.peak_rss_bytes"], "In-process RSS observers")
    stat_table(lines, rows, [
        "source_reads.calls", "source_reads.requested_bytes", "source_reads.returned_bytes",
        "sink.accepted_bytes", "sink.write_calls", "sink.largest_write",
    ], "Source read and sink write dimensions")
    histogram_table(lines, rows)
    allocator_rows = [row for row in rows if row["instrumentation"] == "allocator"]
    stat_table(lines, allocator_rows, [
        "allocation.allocation_calls", "allocation.deallocation_calls", "allocation.reallocation_calls",
        "allocation.failed_allocation_calls", "allocation.allocated_bytes", "allocation.deallocated_bytes",
        "allocation.live_bytes_before", "allocation.live_bytes_after", "allocation.peak_live_bytes_before",
        "allocation.peak_live_bytes_after", "allocation.region_peak_live_bytes",
        "allocation.incremental_peak_live_bytes", "allocation.retention_delta_bytes",
    ], "Allocator counters and peaks")
    lines.extend([
        "The allocator `region_peak_live_bytes` values are absolute high-water endpoints. The operation incremental peak in the table is calculated per sample as `region_peak_live_bytes - live_bytes_before`; it is the reported operation quantity. Allocation bytes are requested callback accounting, including the full requested `new_size` for reallocations, and do not measure physical bytes copied.",
        "",
        "## Repeat drift and route flags",
        "",
        "Repeat drift is `(repeat 2 mean - repeat 1 mean) / repeat 1 mean`; route comparisons pair materialized and bounded arms within the same repeat, count, and instrumentation. All changes, including values below 5%, are in the JSON review. Flags use absolute change strictly greater than 5% and retain improvements for review. A zero baseline has no percentage; its JSON `percent` is `null` and carries an explicit zero-baseline status/flag.",
        "",
    ])
    comparison_lines(lines, review["repeat_comparisons"], "Repeat drift flags")
    comparison_lines(lines, review["route_comparisons"], "Materialized/bounded route flags")
    lines.extend([
        "## Scope caveats",
        "",
    ])
    for caveat in review["caveats"]:
        lines.append(f"- {caveat}")
    lines.extend([
        "",
        "The review does not turn these one-append observations into a production performance claim; the explicit bounded-window objective remains open.",
        "",
    ])
    comparison = review["summary_comparison"]
    if comparison["available"]:
        lines.append(f"Independent comparison with `summary.json`: **{'matched' if comparison['matches'] else 'differences found'}** ({len(comparison['differences'])} differences).")
    else:
        lines.append("No `summary.json` was supplied to this rendering run; the review remains raw-report derived.")
    lines.append("")
    return "\n".join(lines)


def write_or_check(path: Path, content: str, check: bool) -> None:
    if check:
        if not path.is_file() or path.read_text(encoding="utf-8") != content:
            fail(f"{path}: rendered content differs")
        return
    path.write_text(content, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--summary", type=Path, default=None)
    parser.add_argument("--markdown", type=Path, default=None)
    parser.add_argument("--review", type=Path, default=None)
    parser.add_argument("--check", action="store_true")
    options = parser.parse_args()
    root = options.root.resolve()
    summary_path = options.summary.resolve() if options.summary is not None else root / "summary.json"
    markdown_path = options.markdown.resolve() if options.markdown is not None else root / "measurements.md"
    review_path = options.review.resolve() if options.review is not None else root / "measurement-review.json"
    review = derive(root, summary_path)
    review_text = json.dumps(review, indent=2, sort_keys=True, allow_nan=False) + "\n"
    markdown = render_markdown(review)
    write_or_check(review_path, review_text, options.check)
    write_or_check(markdown_path, markdown, options.check)
    print(f"rendered {len(review['rows'])} reports / {review['method']['sample_count'] * len(review['rows'])} samples")


if __name__ == "__main__":
    main()
