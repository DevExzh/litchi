#!/usr/bin/env python3
"""Recompute the retained 0464 PPTX pair-lifecycle measurements.

This module follows the frozen ``protocol.json`` and the Rust report schema
exactly.  Reports contain ``samples_raw`` rows.  The timed API clocks are
``open_source_ns``, ``open_destination_ns``, ``plan_ns``, and
``publication_ns``; ``api_sum_ns`` is their checked sum.  The nine phase rows
are diagnostics, each with explicit provider/cache/budget/RSS availability.
Normal reports intentionally omit allocation fields; allocator reports carry
one operation-scoped allocation object for each timed API phase.

The emitted summary is descriptive evidence for this named pair and provider
matrix.  It makes no performance comparison or optimization claim.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import statistics
from pathlib import Path
from typing import Any, Iterable


ROOT = Path(__file__).resolve().parent
CHANGE = 464
SCHEMA = "litchi-0464-summary-v1"
PROTOCOL_SCHEMA = "litchi-0464-pptx-pair-capture-v1"
REPORT_SCHEMA = "pptx_pair_lifecycle_v1"
PHASES = (
    "baseline",
    "opened",
    "planned",
    "published",
    "drop_result",
    "drop_plan",
    "drop_view",
    "drop_caller_sources",
    "drop_sink",
)
TIMINGS = ("open_source_ns", "open_destination_ns", "open_ns", "plan_ns", "publication_ns", "api_sum_ns")
ALLOC_TIMINGS = ("open_source", "open_destination", "plan", "publication")
PROVIDERS = {"bytes", "range"}
INSTRUMENTS = {"normal", "allocator"}
REPEATS = {"R1", "R2"}
BOOTSTRAP_RESAMPLES = 2000


class SummaryError(ValueError):
    pass


def fail(message: str) -> None:
    raise SummaryError(message)


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON ({error})")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty string")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def finite(value: Any, label: str, minimum: float = 0) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or value < minimum:
        fail(f"{label}: expected finite number >= {minimum}")
    return value


def sha_file(path: Path) -> str:
    hasher = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                hasher.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return hasher.hexdigest()


def digest(value: Any, label: str) -> str:
    value = text(value, label).lower()
    if len(value) != 64 or any(character not in "0123456789abcdef" for character in value):
        fail(f"{label}: expected lowercase SHA-256")
    return value


def artifact(path: Path, root: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        fail(f"artifact is missing, symlinked, or non-regular: {path}")
    return {"path": path.relative_to(root).as_posix(), "bytes": path.stat().st_size, "sha256": sha_file(path)}


def check_artifact_ref(path: Path, root: Path, expected: dict[str, Any], label: str) -> dict[str, Any]:
    row = obj(expected, label)
    expected_path = text(row.get("path"), f"{label}.path")
    expected_bytes = integer(row.get("bytes"), f"{label}.bytes")
    expected_sha = digest(row.get("sha256"), f"{label}.sha256")
    if path.is_symlink() or not path.is_file() or path.stat().st_size != expected_bytes or sha_file(path) != expected_sha:
        fail(f"{label}: artifact identity differs")
    if Path(expected_path).name != path.name:
        fail(f"{label}.path: filename differs")
    return {"path": path.relative_to(root).as_posix(), "bytes": expected_bytes, "sha256": expected_sha}


def protocol(root: Path) -> tuple[Path, dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    path = root / "protocol.json"
    value = obj(load(path, "protocol.json"), "protocol.json")
    if value.get("schema") != PROTOCOL_SCHEMA or value.get("change") != CHANGE or value.get("status") != "frozen":
        fail("protocol identity differs")
    if integer(value.get("samples"), "protocol.samples") != 30 or integer(value.get("warmups"), "protocol.warmups") != 3:
        fail("protocol sample configuration differs")
    matrix = obj(value.get("matrix"), "protocol.matrix")
    if matrix.get("reports_total") != 8 or matrix.get("retained_samples_per_report") != 30 or matrix.get("warmups_per_report") != 3:
        fail("protocol matrix counts differ")
    orders = obj(value.get("orders"), "protocol.orders")
    lanes: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    for repeat in ("R1", "R2"):
        rows = orders.get(repeat)
        if not isinstance(rows, list) or len(rows) != 4:
            fail(f"protocol.orders.{repeat}: expected four lanes")
        for index, lane_value in enumerate(rows):
            lane = obj(lane_value, f"protocol.orders.{repeat}[{index}]")
            lane_id = text(lane.get("lane"), f"protocol.orders.{repeat}[{index}].lane")
            provider = text(lane.get("provider"), f"protocol.orders.{repeat}[{index}].provider")
            instrument = text(lane.get("instrumentation"), f"protocol.orders.{repeat}[{index}].instrumentation")
            if provider not in PROVIDERS or instrument not in INSTRUMENTS or lane.get("repeat") != repeat:
                fail(f"protocol lane dimensions differ: {lane_id}")
            dimensions = (provider, instrument, repeat)
            if dimensions in seen:
                fail(f"protocol lane is repeated: {lane_id}")
            seen.add(dimensions)
            lanes.append({"id": lane_id, "provider": provider, "instrumentation": instrument, "repeat": repeat})
    expected = {(provider, instrument, repeat) for provider in PROVIDERS for instrument in INSTRUMENTS for repeat in REPEATS}
    if seen != expected:
        fail(f"protocol lane coverage differs: {sorted(expected - seen)}")
    inputs = obj(value.get("inputs"), "protocol.inputs")
    limits = obj(inputs.get("limits"), "protocol.inputs.limits")
    range_spec = obj(value.get("range"), "protocol.range")
    if range_spec.get("max_range_bytes") != limits.get("max_range_bytes") or range_spec.get("delay_us") != limits.get("delay_us"):
        fail("protocol range limits disagree")
    return path, value, lanes, {"inputs": inputs, "limits": limits, "range": range_spec, "matrix": matrix}


def phase_point(value: Any, label: str) -> dict[str, Any]:
    point = obj(value, label)
    availability = point.get("availability")
    if availability not in {"available", "unavailable"}:
        fail(f"{label}.availability: expected available or unavailable")
    reason = point.get("unavailable_reason")
    if availability == "unavailable":
        if not isinstance(reason, str) or not reason:
            fail(f"{label}: unavailable point needs a reason")
        if any(key not in {"availability", "unavailable_reason"} and item is not None for key, item in point.items()):
            fail(f"{label}: unavailable point contains fabricated metrics")
    elif reason is not None:
        fail(f"{label}: available point has an unavailable reason")
    return point


def numeric_fields(value: dict[str, Any]) -> dict[str, list[int | float]]:
    result: dict[str, list[int | float]] = {}
    for key, item in value.items():
        if isinstance(item, bool) or item is None:
            continue
        if isinstance(item, (int, float)) and math.isfinite(float(item)):
            result.setdefault(key, []).append(item)
    return result


def output_path(report_path: Path, artifact_value: dict[str, Any]) -> Path:
    reported = Path(text(artifact_value.get("path"), "output_artifact.path"))
    if reported.name != "output.pptx":
        fail("output_artifact.path must name output.pptx")
    if reported.parent.name != report_path.parent.name:
        fail("output_artifact.path must remain in its lane directory")
    return report_path.parent / "output.pptx"


def validate_lane(root: Path, protocol_value: dict[str, Any], config: dict[str, Any], lane: dict[str, Any]) -> dict[str, Any]:
    lane_dir = root / "captures" / lane["repeat"] / lane["id"]
    report_path = lane_dir / "report.json"
    receipt_path = lane_dir / "receipt.json"
    output = lane_dir / "output.pptx"
    for path in (report_path, receipt_path, output):
        if path.is_symlink() or not path.is_file():
            fail(f"{lane['id']}: missing or symlinked retained artifact {path.relative_to(root)}")
    report = obj(load(report_path, str(report_path.relative_to(root))), str(report_path.relative_to(root)))
    receipt = obj(load(receipt_path, str(receipt_path.relative_to(root))), str(receipt_path.relative_to(root)))
    label = str(report_path.relative_to(root))
    if report.get("schema") != REPORT_SCHEMA or report.get("pair_id") != protocol_value["pair_id"]:
        fail(f"{label}: report schema or pair identity differs")
    if report.get("provider") != lane["provider"] or report.get("repeat") != lane["repeat"]:
        fail(f"{label}: provider/repeat differs")
    if report.get("samples") != 30 or report.get("warmup") != 3 or report.get("checked_iteration_count") != 33:
        fail(f"{label}: report sample configuration differs")
    if report.get("source_revision") != protocol_value["pair_manifest"]["source_revision"]:
        fail(f"{label}: source revision differs")
    expected_instrumentation = "system_allocator_operation_scoped" if lane["instrumentation"] == "allocator" else "none"
    if report.get("instrumentation") != expected_instrumentation:
        fail(f"{label}: instrumentation differs")
    for scope in ("provider_scope", "timing_scope", "api_sum_scope", "allocation_scope", "range_scope", "oracle_scope"):
        if not isinstance(report.get(scope), str) or not report[scope]:
            fail(f"{label}: declared {scope} is missing")
    for role in ("source", "destination"):
        identity = obj(report.get(role), f"{label}.{role}")
        expected = config["inputs"][role]
        if identity.get("sha256") != expected.get("sha256") or identity.get("bytes") != expected.get("bytes"):
            fail(f"{label}: {role} input identity differs from protocol")
    if receipt.get("status") != "pass" or receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
        fail(f"{label}: capture receipt is not a successful workload and oracle receipt")
    receipt_lane = obj(receipt.get("lane"), f"{label}.receipt.lane")
    if (receipt_lane.get("lane"), receipt_lane.get("provider"), receipt_lane.get("instrumentation"), receipt_lane.get("repeat")) != (lane["id"], lane["provider"], lane["instrumentation"], lane["repeat"]):
        fail(f"{label}: receipt lane differs")
    if receipt.get("inputs_unchanged") is not True or receipt.get("source_unchanged") is not True:
        fail(f"{label}: receipt does not prove input/source custody")
    output_identity = check_artifact_ref(output, root, obj(report.get("output_artifact"), f"{label}.output_artifact"), f"{label}.output_artifact")
    if output_path(report_path, report["output_artifact"]) != output:
        fail(f"{label}: output path does not point beside report")
    receipt_output = obj(receipt.get("output"), f"{label}.receipt.output")
    if receipt_output.get("sha256") != output_identity["sha256"] or receipt_output.get("bytes") != output_identity["bytes"]:
        fail(f"{label}: receipt output identity differs")
    rows = report.get("samples_raw")
    if not isinstance(rows, list) or len(rows) != 30:
        fail(f"{label}: samples_raw must contain 30 rows")
    configured = obj(report.get("configured_limits"), f"{label}.configured_limits")
    expected_limits = config["limits"]
    for actual_key, protocol_key in (("memory_bytes", "memory_bytes"), ("input_bytes", "input_bytes"), ("output_bytes", "output_bytes"), ("max_write_bytes", "max_write_bytes")):
        if configured.get(actual_key) != expected_limits.get(protocol_key):
            fail(f"{label}: configured limit {actual_key} differs")
    if configured.get("max_range_bytes") != config["range"]["max_range_bytes"]:
        fail(f"{label}: max range limit differs")
    if configured.get("fixed_delay_us") != config["range"]["delay_us"]:
        fail(f"{label}: range delay differs")
    phase_descriptions = report.get("phases")
    if not isinstance(phase_descriptions, list) or [obj(item, f"{label}.phases[{i}]").get("label") for i, item in enumerate(phase_descriptions)] != list(PHASES):
        fail(f"{label}: phase descriptions differ")
    timing_values: dict[str, list[int | float]] = {name: [] for name in TIMINGS}
    allocation_values: dict[str, dict[str, list[int | float]]] = {name: {} for name in ALLOC_TIMINGS}
    phase_values: dict[str, dict[str, list[int | float]]] = {}
    phase_rss: dict[str, dict[str, list[int | float]]] = {}
    phase_unavailable: dict[str, dict[str, str]] = {}
    output_hashes: set[tuple[str, int]] = set()
    for index, row_value in enumerate(rows):
        row = obj(row_value, f"{label}.samples_raw[{index}]")
        if row.get("sample_index") != index:
            fail(f"{label}: sample indices are not exactly 0..29")
        if row.get("output_sha256") != output_identity["sha256"] or row.get("output_bytes") != output_identity["bytes"]:
            fail(f"{label}.samples_raw[{index}]: output identity differs")
        output_hashes.add((row["output_sha256"], row["output_bytes"]))
        semantic = obj(row.get("semantic_oracle"), f"{label}.samples_raw[{index}].semantic_oracle")
        for key in ("output_semantics_verified", "destination_slide_order_verified", "source_direct_images_verified", "raw_untouched_destination_records_verified", "copied_source_payloads_verified", "source_unchanged_after_publication"):
            if semantic.get(key) is not True:
                fail(f"{label}.samples_raw[{index}].semantic_oracle.{key} is not true")
        if semantic.get("independent_external_oracle_required") is not True:
            fail(f"{label}.samples_raw[{index}]: external oracle requirement missing")
        raw_oracle = obj(row.get("raw_oracle"), f"{label}.samples_raw[{index}].raw_oracle")
        for key in ("untouched_destination_records_verified", "copied_source_payloads_verified"):
            if raw_oracle.get(key) is not True:
                fail(f"{label}.samples_raw[{index}].raw_oracle.{key} is not true")
        timings = obj(row.get("timings"), f"{label}.samples_raw[{index}].timings")
        for key in TIMINGS:
            timing_values[key].append(integer(timings.get(key), f"{label}.samples_raw[{index}].timings.{key}"))
        if timing_values["open_ns"][-1] != timing_values["open_source_ns"][-1] + timing_values["open_destination_ns"][-1]:
            fail(f"{label}.samples_raw[{index}]: open timing sum differs")
        if timing_values["api_sum_ns"][-1] != timing_values["open_ns"][-1] + timing_values["plan_ns"][-1] + timing_values["publication_ns"][-1]:
            fail(f"{label}.samples_raw[{index}]: api_sum_ns differs from API phase sum")
        present_alloc = [key for key in timings if key.endswith("_allocation_metrics")]
        if lane["instrumentation"] == "normal":
            if present_alloc or "allocator_counter_revision" in report:
                fail(f"{label}: normal report contains allocation fields")
        else:
            expected_alloc = sorted(f"{name}_allocation_metrics" for name in ALLOC_TIMINGS)
            if report.get("allocator_counter_revision") in (None, "") or sorted(present_alloc) != expected_alloc:
                fail(f"{label}: allocator report allocation fields are incomplete")
            for name in ALLOC_TIMINGS:
                metrics = obj(timings.get(f"{name}_allocation_metrics"), f"{label}.samples_raw[{index}].timings.{name}_allocation_metrics")
                if metrics.get("status") != "measured":
                    fail(f"{label}.samples_raw[{index}].timings.{name}_allocation_metrics: status is not measured")
                for metric, values in numeric_fields(metrics).items():
                    allocation_values[name].setdefault(metric, []).extend(values)
        phase_rows = row.get("phases")
        if not isinstance(phase_rows, list) or len(phase_rows) != len(PHASES):
            fail(f"{label}.samples_raw[{index}]: expected nine phase diagnostics")
        if [obj(item, f"{label}.samples_raw[{index}].phases[{j}]").get("label") for j, item in enumerate(phase_rows)] != list(PHASES):
            fail(f"{label}.samples_raw[{index}]: phase labels differ")
        for phase_value in phase_rows:
            phase = obj(phase_value, f"{label}.phase")
            name = text(phase.get("label"), f"{label}.phase.label")
            values = phase_values.setdefault(name, {})
            for owner in ("source_reads", "destination_reads", "source_cache", "destination_cache"):
                point = phase_point(phase.get(owner), f"{label}.{name}.{owner}")
                if point["availability"] == "available":
                    for metric, items in numeric_fields(point).items():
                        values.setdefault(owner + "." + metric, []).extend(items)
                else:
                    phase_unavailable.setdefault(name, {})[owner] = point["unavailable_reason"]
            for owner in ("source_budget", "destination_budget"):
                point = obj(phase.get(owner), f"{label}.{name}.{owner}")
                for metric, items in numeric_fields(point).items():
                    values.setdefault(owner + "." + metric, []).extend(items)
            rss = phase_point(phase.get("rss"), f"{label}.{name}.rss")
            if rss["availability"] == "available":
                for metric, items in numeric_fields(rss).items():
                    phase_rss.setdefault(name, {}).setdefault(metric, []).extend(items)
            else:
                phase_unavailable.setdefault(name, {})["rss"] = rss["unavailable_reason"]
    if output_hashes != {(output_identity["sha256"], output_identity["bytes"])}:
        fail(f"{label}: retained output identity coverage differs")
    return {"id": lane["id"], "provider": lane["provider"], "instrumentation": lane["instrumentation"], "repeat": lane["repeat"], "report": artifact(report_path, root), "output": output_identity, "timing_values": timing_values, "allocation_values": allocation_values, "phase_values": phase_values, "phase_rss": phase_rss, "phase_unavailable": phase_unavailable, "provider_metadata": {"provider": lane["provider"], "instrumentation": report["instrumentation"], "scope": report["provider_scope"]}, "limits": configured}


def statistic(values: Iterable[int | float], seed: int, resamples: int = BOOTSTRAP_RESAMPLES) -> dict[str, Any]:
    clean = [finite(value, "metric") for value in values]
    if not clean:
        fail("empty metric")
    ordered = sorted(clean)
    rng = random.Random(seed)
    boot = sorted(statistics.fmean(rng.choices(clean, k=len(clean))) for _ in range(resamples))
    return {"count": len(clean), "min": ordered[0], "p50": ordered[max(0, math.ceil(0.50 * len(ordered)) - 1)], "p95": ordered[max(0, math.ceil(0.95 * len(ordered)) - 1)], "p99": ordered[max(0, math.ceil(0.99 * len(ordered)) - 1)], "max": ordered[-1], "mean": statistics.fmean(clean), "mean_bootstrap_95": {"low": boot[max(0, math.ceil(0.025 * resamples) - 1)], "high": boot[min(resamples - 1, math.ceil(0.975 * resamples) - 1)], "seed": seed, "resamples": resamples, "method": "fixed-seed iid bootstrap of the arithmetic mean"}}


def summarize(root: Path = ROOT) -> dict[str, Any]:
    protocol_path, protocol_value, lanes, config = protocol(root)
    validated = [validate_lane(root, protocol_value, config, lane) for lane in lanes]
    seed = integer(protocol_value.get("bootstrap_seed", CHANGE * 10000 + 1), "protocol.bootstrap_seed")
    summary_lanes: list[dict[str, Any]] = []
    for lane_index, lane in enumerate(validated):
        timings = {name: statistic(values, seed + lane_index * 1000 + offset) for offset, (name, values) in enumerate(lane["timing_values"].items())}
        allocation: dict[str, dict[str, Any]] = {}
        if lane["instrumentation"] == "allocator":
            for phase_index, (phase, metrics) in enumerate(lane["allocation_values"].items()):
                allocation[phase] = {name: statistic(values, seed + lane_index * 1000 + 100 + phase_index * 31 + metric_index) for metric_index, (name, values) in enumerate(sorted(metrics.items()))}
        phase_summary: dict[str, Any] = {}
        for phase_index, phase in enumerate(PHASES):
            providers = lane["phase_values"].get(phase, {})
            rss = lane["phase_rss"].get(phase, {})
            phase_summary[phase] = {"provider_counters": {name: statistic(values, seed + lane_index * 1000 + 300 + phase_index * 31 + metric_index) for metric_index, (name, values) in enumerate(sorted(providers.items()))}, "rss": {name: statistic(values, seed + lane_index * 1000 + 600 + phase_index * 31 + metric_index) for metric_index, (name, values) in enumerate(sorted(rss.items()))}, "declared_unavailable": lane["phase_unavailable"].get(phase, {})}
        summary_lanes.append({"id": lane["id"], "provider": lane["provider"], "instrumentation": lane["instrumentation"], "repeat": lane["repeat"], "report": lane["report"], "output": lane["output"], "samples": 30, "warmup": 3, "timings": timings, "allocation_counters": allocation, "phase_diagnostics": phase_summary, "provider_metadata": lane["provider_metadata"], "provider_limits": {"max_range_bytes": config["range"]["max_range_bytes"], "delay_us": config["range"]["delay_us"]}, "configured_limits": lane["limits"]})
    return {"schema": SCHEMA, "change": CHANGE, "claims": ["descriptive PPTX pair-lifecycle timing, provider diagnostics, operation-scoped allocator counters where instrumented, and phase/process RSS"], "claims_excluded": ["performance comparison, optimization, speedup, causality, physical/network I/O, native Office acceptance, and normal-lane allocator totals"], "capture_protocol": {"path": protocol_path.relative_to(root).as_posix(), "bytes": protocol_path.stat().st_size, "sha256": sha_file(protocol_path)}, "pair_id": protocol_value["pair_id"], "samples": 30, "warmup": 3, "phase_order": list(PHASES), "bootstrap": {"seed": seed, "resamples": BOOTSTRAP_RESAMPLES, "method": "fixed-seed iid bootstrap of the arithmetic mean"}, "lanes": summary_lanes}


def write_summary(root: Path = ROOT, output: Path | None = None) -> dict[str, Any]:
    value = summarize(root)
    target = output or root / "summary.json"
    target.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return value


def check_summary(root: Path = ROOT, summary_path: Path | None = None) -> dict[str, Any]:
    expected = summarize(root)
    path = summary_path or root / "summary.json"
    if obj(load(path, str(path)), str(path)) != expected:
        fail("summary.json is not the deterministic recomputation")
    return expected


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    try:
        if args.check:
            check_summary(args.root, args.output)
        else:
            write_summary(args.root, args.output)
        print(json.dumps({"status": "pass", "schema": SCHEMA}, sort_keys=True))
        return 0
    except (SummaryError, OSError, ValueError) as error:
        print(json.dumps({"status": "failed", "error": str(error)}, sort_keys=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
