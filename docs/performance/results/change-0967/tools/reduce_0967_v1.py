#!/usr/bin/env python3
"""Fail-closed offline reducer for 0967 matched packets.

Only retained JSON packets are read. No executable, source, Cargo command, or
workload is launched. Missing evidence yields ``pending_incomplete`` and never
silently becomes a candidate rejection or qualification.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import random
import sys
from pathlib import Path
from typing import Any, Callable

SCHEMA = "litchi.performance.0967.reduction.v1"
RUNNER_SCHEMA = "litchi.performance.0967.runner.v1"
CONTROL_SCHEMA = "litchi.performance.0967.doc_revision_positive_control.v1"


def load_runner() -> Any:
    path = Path(__file__).with_name("run_0967_v1.py")
    spec = importlib.util.spec_from_file_location("run_0967_v1", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load runner validator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


runner = load_runner()


def load(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"{path}: expected object")
    return value


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def percentile(values: list[float], q: float) -> float:
    if not values:
        raise ValueError("percentile of empty values")
    ordered = sorted(float(v) for v in values)
    if len(ordered) == 1:
        return ordered[0]
    position = (len(ordered) - 1) * q
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def bootstrap_median(values: list[float], seed: int, resamples: int) -> dict[str, float]:
    if not values:
        raise ValueError("bootstrap of empty values")
    rng = random.Random(seed)
    n = len(values)
    medians = []
    for _ in range(resamples):
        sample = [values[rng.randrange(n)] for _ in range(n)]
        medians.append(percentile(sample, 0.5))
    return {"median": percentile(values, 0.5), "lower": percentile(medians, 0.025), "upper": percentile(medians, 0.975)}


def ratio(numerator: float, denominator: float, label: str) -> float:
    if denominator <= 0:
        raise ValueError(f"{label}: non-positive baseline metric")
    return numerator / denominator


def packet_path_list(manifest: dict[str, Any], manifest_path: Path) -> list[Path]:
    paths = manifest.get("packets")
    if not isinstance(paths, list):
        raise ValueError("capture manifest has no packet list")
    result = []
    for raw in paths:
        path = Path(raw)
        if not path.is_absolute():
            path = manifest_path.parent / path
        result.append(path)
    return result


def check_packet(plan: dict[str, Any], packet: dict[str, Any]) -> None:
    if packet.get("schema") != RUNNER_SCHEMA:
        raise ValueError(f"packet {packet.get('packet_id')}: schema mismatch")
    terminal = packet.get("terminal") or {}
    if terminal.get("returncode") != 0 or terminal.get("timed_out") or terminal.get("launch_error"):
        raise ValueError(f"packet {packet.get('packet_id')}: terminal failure")
    if packet.get("validation_error") is not None:
        raise ValueError(f"packet {packet.get('packet_id')}: {packet['validation_error']}")
    role = packet.get("role")
    binary_role = "native_baseline" if role == "aa_baseline" else role
    binary = plan.get("binaries", {}).get(binary_role)
    if not binary:
        raise ValueError(f"packet {packet.get('packet_id')}: unknown role {role}")
    observed_binary = packet.get("binary") or {}
    if observed_binary.get("path") != binary.get("path") or observed_binary.get("sha256") != binary.get("sha256"):
        raise ValueError(f"packet {packet.get('packet_id')}: binary custody mismatch")
    context = packet.get("context") or {}
    if context.get("source_revision") != plan["source"]["revision"]:
        raise ValueError(f"packet {packet.get('packet_id')}: source revision mismatch")
    report = packet.get("report")
    if not isinstance(report, dict):
        raise ValueError(f"packet {packet.get('packet_id')}: missing parsed report")
    if str(role).startswith("timed_control_"):
        context = packet.get("context") or {}
        cell_id = (context.get("cell") or {}).get("id")
        cells = {row.get("id"): row for row in plan["controls"]["timed_scope"]["cells"]}
        cell = cells.get(cell_id)
        if cell is None:
            raise ValueError(f"packet {packet.get('packet_id')}: unknown timed-control cell")
        input_row = plan["controls"]["timed_scope"]["inputs"].get(cell["control"])
        if not input_row or input_row.get("available") is not True:
            raise ValueError(f"packet {packet.get('packet_id')}: timed-control input not bound")
        runner.control_profile_validator(cell, report, context.get("variant"), str(role).startswith("timed_control_allocator"), input_row)
    elif role in {"native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate", "aa_baseline"}:
        check_real_report(plan, packet)


def child_metric(packet: dict[str, Any], metric: str) -> float:
    report = packet["report"]
    if metric in {"p50", "p95", "p99"}:
        elapsed = report.get("elapsed_ns")
        if not isinstance(elapsed, list) or not elapsed:
            raise ValueError(f"{packet['packet_id']}: elapsed_ns unavailable")
        values = [float(item) for item in elapsed]
        if metric == "p50":
            return percentile(values, 0.50)
        if metric == "p95":
            return percentile(values, 0.95)
        return max(values)
    allocation = report.get("allocation")
    if not isinstance(allocation, dict):
        raise ValueError(f"{packet['packet_id']}: allocation metrics unavailable")
    if metric == "rss_kb":
        rss = packet.get("rss", {}).get("kb")
        if not isinstance(rss, int) or rss <= 0:
            raise ValueError(f"{packet['packet_id']}: RSS unavailable")
        return float(rss)
    value = allocation.get(metric)
    if not isinstance(value, int) or value <= 0:
        raise ValueError(f"{packet['packet_id']}: invalid {metric}")
    return float(value)


def check_real_report(plan: dict[str, Any], packet: dict[str, Any]) -> None:
    context = packet.get("context") or {}
    cell = context.get("cell") or {}
    report = packet.get("report") or {}
    if report.get("schema") != runner.PROBE_SCHEMA:
        raise ValueError(f"{packet['packet_id']}: real report schema mismatch")
    case = cell.get("case")
    fixture = plan["fixtures"].get(case)
    if not fixture or report.get("case") != case or report.get("edits") != cell.get("edits") or report.get("source_sha256") != fixture["sha256"] or report.get("source_bytes") != fixture["bytes"]:
        raise ValueError(f"{packet['packet_id']}: real cell/source binding mismatch")
    retention = {"zero": "zero", "default": "default_8mib", "release": "release_8mib"}[cell["retention"]]
    if report.get("retention") != retention:
        raise ValueError(f"{packet['packet_id']}: real retention mismatch")
    allocator = str(packet.get("role", "")).startswith("allocator_")
    if report.get("warmups") != (0 if allocator else 5) or report.get("samples_requested") != (1 if allocator else 80):
        raise ValueError(f"{packet['packet_id']}: real schedule mismatch")
    if bool(report.get("allocator_instrumented")) != allocator or bool(report.get("timing_claim")) == allocator:
        raise ValueError(f"{packet['packet_id']}: real lane flags mismatch")
    elapsed = report.get("elapsed_ns")
    if not isinstance(elapsed, list) or len(elapsed) != (0 if allocator else 80):
        raise ValueError(f"{packet['packet_id']}: real elapsed count mismatch")
    rows = report.get("samples")
    if not isinstance(rows, list) or len(rows) != (1 if allocator else 80):
        raise ValueError(f"{packet['packet_id']}: real sample rows mismatch")
    correctness = report.get("correctness")
    if not isinstance(correctness, dict) or any(value is False for value in correctness.values() if isinstance(value, bool)):
        raise ValueError(f"{packet['packet_id']}: real correctness gate failed")
    for row in rows:
        if not isinstance(row, dict) or row.get("exact_output_equal_reference") is not True:
            raise ValueError(f"{packet['packet_id']}: real output oracle failed")
        if allocator:
            allocation = row.get("allocation")
            if not isinstance(allocation, dict) or any(not isinstance(allocation.get(field), int) for field in ("allocated_bytes", "deallocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes")):
                raise ValueError(f"{packet['packet_id']}: real allocation fields incomplete")
        elif row.get("elapsed_ns") is None:
            raise ValueError(f"{packet['packet_id']}: real elapsed sample missing")


def expected_pair_order(plan: dict[str, Any], lane: str, pair: int) -> str:
    schedule = plan["real_matrix"][lane]["schedule"]
    rows = {int(row["pair"]): row for row in schedule}
    try:
        return rows[pair]["relative_order"]
    except KeyError as exc:
        raise ValueError(f"{lane}: pair {pair} is absent from the declared schedule") from exc


def group_real(plan: dict[str, Any], packets: list[dict[str, Any]], lane: str) -> dict[tuple[str, int, str], dict[str, dict[int, dict[str, Any]]]]:
    grouped: dict[tuple[str, int, str], dict[str, dict[int, dict[str, Any]]]] = {}
    for packet in packets:
        context = packet.get("context") or {}
        cell = context.get("cell") or {}
        key = (cell.get("case", ""), int(cell.get("edits", -1)), cell.get("retention", ""))
        expected_cell_id = f"{key[0]}/e{key[1]}/{key[2]}"
        if cell.get("cell_id") != expected_cell_id:
            raise ValueError(f"{packet.get('packet_id')}: real cell identity mismatch")
        variant = context.get("variant")
        pair = int(context.get("pair", -1))
        if variant not in {"baseline", "candidate"} or pair < 1:
            raise ValueError(f"{packet.get('packet_id')}: invalid real packet context")
        relative_order = context.get("relative_order")
        if relative_order != expected_pair_order(plan, lane, pair):
            raise ValueError(f"{packet.get('packet_id')}: declared relative order does not match {lane} pair {pair}")
        grouped.setdefault(key, {}).setdefault(variant, {})
        if pair in grouped[key][variant]:
            raise ValueError(f"duplicate {lane} packet for {key} {variant} pair {pair}")
        grouped[key][variant][pair] = packet
    return grouped


def pair_metrics(group: dict[str, dict[int, dict[str, Any]]], metric: str, pairs: int, seed: int, resamples: int) -> dict[str, Any]:
    baseline = group.get("baseline", {})
    candidate = group.get("candidate", {})
    if set(baseline) != set(range(1, pairs + 1)) or set(candidate) != set(range(1, pairs + 1)):
        raise ValueError("incomplete or extra paired children")
    bvalues = [child_metric(baseline[index], metric) for index in range(1, pairs + 1)]
    cvalues = [child_metric(candidate[index], metric) for index in range(1, pairs + 1)]
    ratios = [ratio(c, b, metric) for c, b in zip(cvalues, bvalues)]
    ci = bootstrap_median(ratios, seed, resamples)
    return {"baseline_child": bvalues, "candidate_child": cvalues, "ratios": ratios, "bootstrap": ci, "max_ratio": max(ratios)}


def aa_metrics(packets: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    grouped: dict[str, dict[str, dict[int, dict[str, Any]]]] = {}
    for packet in packets:
        context = packet.get("context") or {}
        cell = context.get("cell", {}).get("cell_id")
        order = context.get("aa_order")
        position = context.get("aa_position")
        pair = int(context.get("pair", -1))
        if cell not in plan["aa_noise"]["cells"] or order not in {"AB", "BA"} or position not in {"first", "second"}:
            raise ValueError(f"invalid A/A packet context: {packet.get('packet_id')}")
        expected_order = "AB" if pair <= plan["aa_noise"]["native_pairs_per_cell"] // 2 else "BA"
        if pair < 1 or order != expected_order:
            raise ValueError(f"invalid A/A pair order: {packet.get('packet_id')}")
        grouped.setdefault(cell, {}).setdefault(order, {}).setdefault(pair, {})[position] = packet
    result: dict[str, Any] = {}
    aa = plan["aa_noise"]
    expected_pairs = set(range(1, aa["native_pairs_per_cell"] + 1))
    for cell in aa["cells"]:
        if cell not in grouped:
            raise ValueError(f"missing A/A cell {cell}")
        cell_result = {}
        for order in ("AB", "BA"):
            rows = grouped[cell].get(order, {})
            expected_order_pairs = {pair for pair in expected_pairs if ("AB" if pair <= aa["native_pairs_per_cell"] // 2 else "BA") == order}
            if set(rows) != expected_order_pairs or any(set(item) != {"first", "second"} for item in rows.values()):
                raise ValueError(f"incomplete A/A {cell} {order}")
            directional = {}
            for metric in ("p50", "p95"):
                values = [ratio(child_metric(rows[index]["second"], metric), child_metric(rows[index]["first"], metric), metric) for index in sorted(rows)]
                ci = bootstrap_median(values, aa["bootstrap_seed"], aa.get("bootstrap_resamples", 10000))
                symmetric_upper = max(ci["upper"], 1.0 / ci["lower"] if ci["lower"] > 0 else float("inf"))
                directional[metric] = {**ci, "symmetric_upper": symmetric_upper, "ratios": values}
            cell_result[order] = directional
        result[cell] = cell_result
    return result


def control_semantics(plan: dict[str, Any], packets: list[dict[str, Any]]) -> dict[str, Any]:
    by_id: dict[str, dict[str, dict[str, Any]]] = {}
    for packet in packets:
        context = packet.get("context") or {}
        control_id = context.get("control_id")
        variant = context.get("variant")
        if not control_id or variant not in {"baseline", "candidate"}:
            raise ValueError(f"invalid control packet context: {packet.get('packet_id')}")
        if control_id in by_id and variant in by_id[control_id]:
            raise ValueError(f"duplicate control packet {control_id} {variant}")
        by_id.setdefault(control_id, {})[variant] = packet
    required = {row["id"] for row in plan["controls"]["matrix"]}
    if set(by_id) != required or any(set(rows) != {"baseline", "candidate"} for rows in by_id.values()):
        raise ValueError("control matrix incomplete")
    rows = {}
    for control in plan["controls"]["matrix"]:
        control_id = control["id"]
        checked = {}
        for variant, packet in by_id[control_id].items():
            report = packet["report"]
            if report.get("schema") != CONTROL_SCHEMA or report.get("variant") != variant or report.get("timing_claim") is not False or report.get("error") is not None:
                raise ValueError(f"{control_id}/{variant}: control report failed basic gates")
            pre = report.get("pre") or {}
            post = report.get("post") or {}
            before = pre.get("revisions") or []
            after = post.get("revisions") or []
            if control_id == "real-bookmark-delete-redline-query":
                query = report.get("query") or {}
                if report.get("mode") != "query" or len(before) != 20 or query.get("replacement_succeeded") is not False or query.get("source_unchanged_after_refusal") is not True:
                    raise ValueError(f"{control_id}/{variant}: query refusal gate failed")
            else:
                edit = report.get("edit") or {}
                if report.get("mode") != "run" or edit.get("replacement_succeeded") is not True or edit.get("reopen_succeeded") is not True or edit.get("changed") is not True:
                    raise ValueError(f"{control_id}/{variant}: lifecycle gate failed")
                witness = report.get("branch_witness") or {}
                if control.get("require_revision_survival") and (not before or after != before or witness.get("positive_branch_required") is not True):
                    raise ValueError(f"{control_id}/{variant}: positive revisions did not survive")
                if control.get("require_false_marker") and (not before or after != before or witness.get("false_marker_present") is not True):
                    raise ValueError(f"{control_id}/{variant}: true revision/false marker gate failed")
            checked[variant] = {"source_sha256": report.get("source_sha256"), "fixture": report.get("fixture"), "revision_count": len(before)}
        if checked["baseline"]["source_sha256"] != checked["candidate"]["source_sha256"]:
            raise ValueError(f"{control_id}: baseline/candidate generated source differs")
        rows[control_id] = checked
    return rows


def group_timed_controls(plan: dict[str, Any], packets: list[dict[str, Any]]) -> dict[str, dict[str, dict[str, dict[int, dict[str, Any]]]]]:
    grouped: dict[str, dict[str, dict[str, dict[int, dict[str, Any]]]]] = {}
    for packet in packets:
        context = packet.get("context") or {}
        cell = (context.get("cell") or {}).get("id")
        variant = context.get("variant")
        lane = context.get("lane")
        pair = int(context.get("pair", -1))
        if not cell or lane not in {"control_native", "control_allocator"} or variant not in {"baseline", "candidate"} or pair < 1:
            raise ValueError(f"invalid timed control packet context: {packet.get('packet_id')}")
        pairs = plan["controls"]["timed_scope"]["native_pairs_per_control"] if lane == "control_native" else plan["controls"]["timed_scope"]["allocator_pairs_per_control"]
        expected_order = "baseline_then_candidate" if pair <= pairs // 2 else "candidate_then_baseline"
        if context.get("relative_order") != expected_order:
            raise ValueError(f"invalid timed control pair order: {packet.get('packet_id')}")
        grouped.setdefault(cell, {}).setdefault(lane, {}).setdefault(variant, {})
        if pair in grouped[cell][lane][variant]:
            raise ValueError(f"duplicate timed control child {cell}/{lane}/{variant}/{pair}")
        grouped[cell][lane][variant][pair] = packet
    return grouped


TIMED_TOP_IDENTITY_FIELDS = (
    "reference_output_sha256",
    "reference_output_bytes",
    "reference_logical_vec_capacity_bytes",
    "reference_revision_count",
    "reference_paragraph_count",
    "reference_semantic_change_count",
)
TIMED_SAMPLE_IDENTITY_FIELDS = (
    "index",
    "output_sha256",
    "output_bytes",
    "logical_vec_capacity_bytes",
    "exact_output_equal_reference",
    "revision_count",
    "paragraph_count",
    "semantic_change_count",
    "paragraph_zero_sha256",
    "public_reopen_succeeded",
    "source_unchanged",
)


def timed_output_semantic_identity(report: dict[str, Any]) -> tuple[Any, ...]:
    """Return the output/correctness identity independent of measurements.

    The baseline and candidate are allowed to differ in elapsed time and
    allocator counters.  Their output bytes and all semantic observations must
    still agree for each timed control cell.  Keep this list explicit so a new
    measurement field cannot silently become part of the cross-variant gate.
    """
    top = tuple((field, report.get(field)) for field in TIMED_TOP_IDENTITY_FIELDS)
    samples = report.get("samples")
    if not isinstance(samples, list):
        raise ValueError("timed report samples are not a list")
    rows: list[tuple[tuple[str, Any], ...]] = []
    for sample in samples:
        if not isinstance(sample, dict):
            raise ValueError("timed report sample is not an object")
        rows.append(tuple((field, sample.get(field)) for field in TIMED_SAMPLE_IDENTITY_FIELDS))
    return (top, tuple(rows))


def require_timed_variant_identity(cell_id: str, lane: str, group: dict[str, dict[int, dict[str, Any]]]) -> None:
    """Require baseline/candidate output and semantic identity in one lane."""
    if set(group) != {"baseline", "candidate"}:
        raise ValueError(f"timed control {cell_id}/{lane}: baseline/candidate reports are incomplete")
    # Each report contains the whole warmup/sample series.  The pair number is
    # only the child schedule; compare every paired child, in pair order.
    baseline_pairs = group["baseline"]
    candidate_pairs = group["candidate"]
    if set(baseline_pairs) != set(candidate_pairs):
        raise ValueError(f"timed control {cell_id}/{lane}: baseline/candidate pair sets differ")
    for pair in sorted(baseline_pairs):
        baseline_identity = timed_output_semantic_identity(baseline_pairs[pair]["report"])
        candidate_identity = timed_output_semantic_identity(candidate_pairs[pair]["report"])
        if baseline_identity != candidate_identity:
            raise ValueError(f"timed control {cell_id}/{lane}/pair-{pair}: baseline/candidate output or semantic identity differs")


def reduce_timed_controls(plan: dict[str, Any], packets: list[dict[str, Any]]) -> dict[str, Any]:
    grouped = group_timed_controls(plan, packets)
    result: dict[str, Any] = {}
    native_pairs = plan["controls"]["timed_scope"]["native_pairs_per_control"]
    allocator_pairs = plan["controls"]["timed_scope"]["allocator_pairs_per_control"]
    seed = plan["controls"]["timed_scope"].get("bootstrap_seed", 967101)
    resamples = plan["controls"]["timed_scope"].get("bootstrap_resamples", 10000)
    expected = {row["id"]: row for row in plan["controls"]["timed_scope"]["cells"]}
    if set(grouped) != set(expected):
        raise ValueError("timed control matrix incomplete")
    for cell_id, cell in expected.items():
        row: dict[str, Any] = {"control": cell["control"], "scope": cell["scope"]}
        native_group = grouped[cell_id].get("control_native")
        allocator_group = grouped[cell_id].get("control_allocator")
        if native_group:
            require_timed_variant_identity(cell_id, "native", native_group)
            ng = native_group
            metrics = {metric: pair_metrics(ng, metric, native_pairs, seed, resamples) for metric in ("p50", "p95", "p99")}
            metrics["noninferior"] = all(metrics[name]["bootstrap"]["upper"] <= 1.05 for name in ("p50", "p95"))
            metrics["median_over_1_05_review"] = metrics["p50"]["bootstrap"]["median"] > 1.05
            metrics["p99_ratio_median"] = metrics["p99"]["bootstrap"]["median"]
            metrics["p99_descriptive_flag"] = metrics["p99_ratio_median"] > 1.05
            row["native"] = metrics
        if allocator_group:
            require_timed_variant_identity(cell_id, "allocator", allocator_group)
            ag = allocator_group
            metrics = {metric: pair_metrics(ag, metric, allocator_pairs, seed + 1, resamples) for metric in ("allocated_bytes", "peak_live_bytes", "allocation_calls", "rss_kb")}
            metrics["guard_pass"] = metrics["allocated_bytes"]["max_ratio"] <= 1.01 and metrics["peak_live_bytes"]["max_ratio"] <= 1.05 and metrics["rss_kb"]["max_ratio"] <= 1.05
            row["allocator"] = metrics
        if "native" not in row or "allocator" not in row:
            raise ValueError(f"timed control {cell_id}: native or allocator lane missing")
        result[cell_id] = row
    return result


def reduce(plan: dict[str, Any], manifest_path: Path, output: Path, selection: Path) -> int:
    manifest = load(manifest_path)
    packet_paths = packet_path_list(manifest, manifest_path)
    packets = [load(path) for path in packet_paths]
    errors = []
    for packet in packets:
        try:
            check_packet(plan, packet)
        except (OSError, ValueError, KeyError, TypeError) as exc:
            errors.append(str(exc))
    if errors:
        result = {"schema": SCHEMA, "status": "rejected_packet_validation", "errors": errors, "build_started": False, "workload_started": True, "timing_claim": False}
        output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
        selection.write_text(json.dumps({"schema": "litchi.performance.0967.selection.v1", "status": "no_decision", "reason": "packet validation failed", "errors": errors}, indent=2) + "\n", encoding="utf-8")
        return 2
    controls = [packet for packet in packets if packet.get("role") in {"control_baseline", "control_candidate"}]
    timed_controls = [packet for packet in packets if str(packet.get("role", "")).startswith("timed_control_")]
    aa = [packet for packet in packets if packet.get("role") == "aa_baseline"]
    native = [packet for packet in packets if packet.get("role") in {"native_baseline", "native_candidate"}]
    allocator = [packet for packet in packets if packet.get("role") in {"allocator_baseline", "allocator_candidate"}]
    reasons = []
    control_result = None
    try:
        control_result = control_semantics(plan, controls)
    except (ValueError, KeyError, TypeError) as exc:
        reasons.append(f"controls: {exc}")
    scenario_path = Path(plan["controls"]["scenario_audit"]["output"])
    scenario = None
    if scenario_path.is_file():
        scenario = load(scenario_path)
        if scenario.get("status") != "pass":
            reasons.append(f"scenario audit status {scenario.get('status')!r}")
    else:
        reasons.append("scenario audit receipt is missing")
    # Timed generated controls are deliberately separate from the correctness probe.
    timed_result = None
    if plan["controls"].get("timed_scope", {}).get("status") != "bound":
        reasons.append("timed positive/dense/false control scope is not bound")
    else:
        try:
            timed_result = reduce_timed_controls(plan, timed_controls)
            if any(not row["native"]["noninferior"] for row in timed_result.values()):
                reasons.append("timed control native noninferiority failed")
            if any(not row["allocator"]["guard_pass"] for row in timed_result.values()):
                reasons.append("timed control allocator guard failed")
        except (ValueError, KeyError, TypeError) as exc:
            reasons.append(f"timed controls: {exc}")
    native_result: dict[str, Any] = {}
    allocator_result: dict[str, Any] = {}
    try:
        ng = group_real(plan, native, "native")
        for cell in plan["real_matrix"]["cells"]:
            key = (cell["case"], cell["edits"], cell["retention"])
            group = ng.get(key)
            if group is None:
                raise ValueError(f"missing native cell {cell['cell_id']}")
            metric_result = {}
            for metric in ("p50", "p95", "p99"):
                if metric == "p99":
                    metric_result[metric] = pair_metrics(group, metric, 12, plan["real_matrix"]["native"]["bootstrap_seed"], plan["real_matrix"]["native"]["bootstrap_resamples"])
                else:
                    metric_result[metric] = pair_metrics(group, metric, 12, plan["real_matrix"]["native"]["bootstrap_seed"], plan["real_matrix"]["native"]["bootstrap_resamples"])
            metric_result["noninferior"] = all(metric_result[m]["bootstrap"]["upper"] <= 1.05 for m in ("p50", "p95"))
            metric_result["median_over_1_05_review"] = metric_result["p50"]["bootstrap"]["median"] > 1.05
            metric_result["p99_ratio_median"] = metric_result["p99"]["bootstrap"]["median"]
            metric_result["p99_descriptive_flag"] = metric_result["p99_ratio_median"] > 1.05
            native_result[cell["cell_id"]] = metric_result
        if len(ng) != 24:
            raise ValueError("extra native cells")
    except (ValueError, KeyError, TypeError) as exc:
        reasons.append(f"native: {exc}")
    try:
        ag = group_real(plan, allocator, "allocator")
        for cell in plan["real_matrix"]["cells"]:
            key = (cell["case"], cell["edits"], cell["retention"])
            group = ag.get(key)
            if group is None:
                raise ValueError(f"missing allocator cell {cell['cell_id']}")
            metrics = {name: pair_metrics(group, name, 3, plan["real_matrix"]["allocator"]["bootstrap_seed"], plan["real_matrix"]["allocator"]["bootstrap_resamples"]) for name in ("allocated_bytes", "peak_live_bytes", "allocation_calls", "rss_kb")}
            metrics["guard_pass"] = metrics["allocated_bytes"]["max_ratio"] <= 1.01 and metrics["peak_live_bytes"]["max_ratio"] <= 1.05 and metrics["rss_kb"]["max_ratio"] <= 1.05
            allocator_result[cell["cell_id"]] = metrics
        if len(ag) != 24:
            raise ValueError("extra allocator cells")
    except (ValueError, KeyError, TypeError) as exc:
        reasons.append(f"allocator: {exc}")
    aa_result = None
    try:
        aa_result = aa_metrics(aa, plan)
        for cell, metrics in aa_result.items():
            for metric in ("p50", "p95"):
                if metrics["AB"][metric]["symmetric_upper"] > 1.05 or metrics["BA"][metric]["symmetric_upper"] > 1.05:
                    reasons.append(f"A/A readiness failed {cell} {metric}")
    except (ValueError, KeyError, TypeError, ZeroDivisionError) as exc:
        reasons.append(f"A/A: {exc}")
    complete = not reasons
    native_pass = complete and all(row.get("noninferior") for row in native_result.values())
    memory_pass = complete and all(row.get("guard_pass") for row in allocator_result.values())
    calls_guard = complete and all(
        row["allocation_calls"]["bootstrap"]["median"] <= 1.01
        for cell, row in allocator_result.items()
        if cell not in {"docfloat/e32/default", "docnohf/e32/default"}
    )
    allocation_benefit = complete and allocator_result.get("docfloat/e32/default", {}).get("allocation_calls", {}).get("bootstrap", {}).get("median", math.inf) <= 0.90 and allocator_result.get("docnohf/e32/default", {}).get("allocation_calls", {}).get("bootstrap", {}).get("median", math.inf) <= 0.97
    native_benefit = complete and any(
        row.get("p50", {}).get("bootstrap", {}).get("median", math.inf) <= 0.95 and row.get("p50", {}).get("bootstrap", {}).get("upper", math.inf) < 1.0
        for cell, row in native_result.items()
        if cell in {"docfloat/e32/default", "docnohf/e32/default"}
    )
    if complete and not native_pass:
        reasons.append("native noninferiority failed")
    if complete and not memory_pass:
        reasons.append("allocator memory/RSS guard failed")
    if complete and not calls_guard:
        reasons.append("non-endpoint allocation_calls guard failed")
    if complete and not (allocation_benefit or native_benefit):
        reasons.append("no predeclared benefit branch passed")
    if not complete:
        status = "pending_incomplete"
    elif not (native_pass and memory_pass and calls_guard and (allocation_benefit or native_benefit)):
        status = "rejected"
    else:
        status = "qualified_candidate"
    result = {
        "schema": SCHEMA,
        "status": status,
        "created_utc": __import__("datetime").datetime.now(__import__("datetime").timezone.utc).isoformat(),
        "plan": str(manifest.get("plan", "")),
        "timing_claim": False,
        "build_started": False,
        "workload_started": True,
        "packet_count": len(packets),
        "reasons": reasons,
        "scenario_audit": scenario,
        "controls": control_result,
        "timed_controls": timed_result,
        "aa_noise": aa_result,
        "native": native_result,
        "allocator": allocator_result,
        "gates": {"native_noninferiority": native_pass, "allocator_memory": memory_pass, "allocation_calls_guard": calls_guard, "allocation_benefit": allocation_benefit, "native_benefit": native_benefit},
        "raw_packets_retained": True,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    selection_result = {"schema": "litchi.performance.0967.selection.v1", "status": "no_decision" if status == "pending_incomplete" else ("qualify" if status == "qualified_candidate" else "reject"), "reduction": str(output), "decision_rule": "qualification requires every fixed gate and either allocation or native benefit branch; incomplete evidence is no decision", "reasons": reasons, "timing_claim": False}
    selection.parent.mkdir(parents=True, exist_ok=True)
    selection.write_text(json.dumps(selection_result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": status, "reasons": reasons, "output": str(output), "selection": str(selection)}))
    return 0 if status in {"pending_incomplete", "qualified_candidate"} else 2


def self_test() -> int:
    if percentile([1, 2, 3, 4], 0.5) != 2.5:
        raise SystemExit("percentile test failed")
    ci = bootstrap_median([1.0, 1.0, 1.0], 967001, 100)
    if ci != {"median": 1.0, "lower": 1.0, "upper": 1.0}:
        raise SystemExit("bootstrap equality test failed")
    if max(1.2, 1 / 0.8) != 1.25:
        raise SystemExit("symmetric A/A test failed")
    identity_plan = {
        "real_matrix": {
            "native": {"schedule": [{"pair": 1, "relative_order": "baseline_then_candidate"}]},
            "allocator": {"schedule": [{"pair": 1, "relative_order": "baseline_then_candidate"}]},
        },
        "aa_noise": {"cells": ["aa"], "native_pairs_per_cell": 2, "bootstrap_seed": 967004, "bootstrap_resamples": 100},
        "controls": {"timed_scope": {"native_pairs_per_control": 2, "allocator_pairs_per_control": 2}},
    }
    real_context = {
        "cell": {"cell_id": "case/e1/default", "case": "case", "edits": 1, "retention": "default"},
        "variant": "baseline",
        "pair": 1,
        "relative_order": "baseline_then_candidate",
    }
    real_packet = {"packet_id": "real", "context": real_context}
    try:
        group_real(identity_plan, [{**real_packet, "context": {**real_context, "relative_order": "candidate_then_baseline"}}], "native")
    except ValueError:
        pass
    else:
        raise SystemExit("real relative-order identity test failed")
    try:
        group_real(identity_plan, [real_packet, real_packet], "native")
    except ValueError:
        pass
    else:
        raise SystemExit("real duplicate identity test failed")
    aa_packet = {"packet_id": "aa", "context": {"cell": {"cell_id": "aa"}, "aa_order": "BA", "aa_position": "first", "pair": 1}}
    try:
        aa_metrics([aa_packet], identity_plan)
    except ValueError:
        pass
    else:
        raise SystemExit("A/A order identity test failed")
    timed_packet = {"packet_id": "timed", "context": {"cell": {"id": "cell"}, "lane": "control_native", "variant": "baseline", "pair": 1, "relative_order": "candidate_then_baseline"}}
    try:
        group_timed_controls(identity_plan, [timed_packet])
    except ValueError:
        pass
    else:
        raise SystemExit("timed-control order identity test failed")
    identity_top = {
        "reference_output_sha256": "output",
        "reference_output_bytes": 10,
        "reference_logical_vec_capacity_bytes": 16,
        "reference_revision_count": 2,
        "reference_paragraph_count": 2,
        "reference_semantic_change_count": 1,
    }
    identity_sample = {
        "index": 0,
        "output_sha256": "output",
        "output_bytes": 10,
        "logical_vec_capacity_bytes": 16,
        "exact_output_equal_reference": True,
        "revision_count": 2,
        "paragraph_count": 2,
        "semantic_change_count": 1,
        "paragraph_zero_sha256": "paragraph",
        "public_reopen_succeeded": True,
        "source_unchanged": False,
        "elapsed_ns": [10],
        "allocation": None,
    }
    baseline_report = {**identity_top, "samples": [identity_sample], "elapsed_ns": [10]}
    candidate_report = {**identity_top, "samples": [{**identity_sample, "elapsed_ns": [11]}], "elapsed_ns": [11]}
    identity_group = {"baseline": {1: {"report": baseline_report}}, "candidate": {1: {"report": candidate_report}}}
    require_timed_variant_identity("cell", "native", identity_group)
    different_report = {**candidate_report, "samples": [{**candidate_report["samples"][0], "semantic_change_count": 2}]}
    try:
        require_timed_variant_identity("cell", "native", {"baseline": {1: {"report": baseline_report}}, "candidate": {1: {"report": different_report}}})
    except ValueError:
        pass
    else:
        raise SystemExit("timed output/semantic identity test failed")
    print(json.dumps({"schema": SCHEMA, "status": "self_test_pass"}))
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", type=Path, default=Path("/home/zhuhe/litchi-performance-0967/plans/plan-0967-v1.json"))
    ap.add_argument("--manifest", type=Path)
    ap.add_argument("--output", type=Path, default=Path("/home/zhuhe/litchi-performance-0967/analysis/reduction-v1.json"))
    ap.add_argument("--selection", type=Path, default=Path("/home/zhuhe/litchi-performance-0967/analysis/selection-v1.json"))
    ap.add_argument("--check-plan", action="store_true")
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--reduce", action="store_true")
    args = ap.parse_args()
    if args.self_test:
        return self_test()
    try:
        plan = load(args.plan.resolve())
        if args.check_plan:
            checked = runner.check_plan(args.plan.resolve(), allow_unbound=True)
            print(json.dumps({"schema": SCHEMA, "status": "plan_check_pass", "plan_status": checked.get("status")}))
            if not args.reduce:
                return 0
        if not args.reduce:
            args.reduce = True
        if not args.manifest:
            raise ValueError("--manifest is required for --reduce")
        return reduce(plan, args.manifest.resolve(), args.output.resolve(), args.selection.resolve())
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(f"reducer rejected input: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
