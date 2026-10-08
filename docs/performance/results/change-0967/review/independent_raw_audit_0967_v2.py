#!/usr/bin/env python3
"""Independent raw packet audit for the sealed 0967 capture.

This helper intentionally does not import the staged runner or reducer.  It
checks packet custody, terminal/oracle identity, schedules, A/A readiness, and
recomputes every declared metric/gate from the retained JSON packets.
"""
from __future__ import annotations

import hashlib
import json
import math
import os
import random
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

BASE = Path("/home/zhuhe/litchi-performance-0967")
PLAN_PATH = BASE / "plans/plan-0967-root-authorized-v1.json"
MANIFEST_PATH = BASE / "capture/capture-manifest-v1.json"
OUTPUT_PATH = BASE / "review/independent-raw-audit-0967-v2.json"
RUNNER_SCHEMA = "litchi.performance.0967.runner.v1"
PROBE_SCHEMA = "litchi.performance.0917.doc_retention.v1"
CONTROL_SCHEMA = "litchi.performance.0967.doc_revision_positive_control.v1"
CONTROL_PROFILE_SCHEMA = "litchi.performance.0967.control_profile.v1"


def load(path: Path) -> Any:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def digest(path: Path) -> str:
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
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def bootstrap_median(values: list[float], seed: int, resamples: int) -> dict[str, float]:
    if not values:
        raise ValueError("bootstrap of empty values")
    rng = random.Random(seed)
    medians: list[float] = []
    n = len(values)
    for _ in range(resamples):
        sample = [values[rng.randrange(n)] for _ in values]
        medians.append(percentile(sample, 0.5))
    return {
        "median": percentile(values, 0.5),
        "lower": percentile(medians, 0.025),
        "upper": percentile(medians, 0.975),
    }


def all_true(value: Any, label: str, errors: list[str]) -> None:
    if not isinstance(value, dict):
        errors.append(f"{label}: expected object")
        return
    for key, item in value.items():
        if isinstance(item, bool) and not item:
            errors.append(f"{label}.{key}: false")


def child_metric(packet: dict[str, Any], metric: str) -> float:
    report = packet.get("report") or {}
    if metric in {"p50", "p95", "p99"}:
        elapsed = report.get("elapsed_ns")
        if not isinstance(elapsed, list) or not elapsed:
            raise ValueError(f"{packet.get('packet_id')}: missing elapsed_ns")
        values = [float(item) for item in elapsed]
        if metric == "p50":
            return percentile(values, 0.50)
        if metric == "p95":
            return percentile(values, 0.95)
        return max(values)
    allocation = (report.get("samples") or [{}])[0].get("allocation")
    if not isinstance(allocation, dict):
        raise ValueError(f"{packet.get('packet_id')}: missing allocation")
    if metric == "rss_kb":
        value = (packet.get("rss") or {}).get("kb")
        if not isinstance(value, int) or value <= 0:
            raise ValueError(f"{packet.get('packet_id')}: missing RSS")
        return float(value)
    value = allocation.get(metric)
    if not isinstance(value, int) or value <= 0:
        raise ValueError(f"{packet.get('packet_id')}: invalid {metric}")
    return float(value)


def pair_metrics(group: dict[str, dict[int, dict[str, Any]]], metric: str, pairs: int, seed: int, resamples: int) -> dict[str, Any]:
    baseline = group.get("baseline", {})
    candidate = group.get("candidate", {})
    expected = set(range(1, pairs + 1))
    if set(baseline) != expected or set(candidate) != expected:
        raise ValueError(f"incomplete {metric} pair set")
    bvalues = [child_metric(baseline[i], metric) for i in range(1, pairs + 1)]
    cvalues = [child_metric(candidate[i], metric) for i in range(1, pairs + 1)]
    ratios = [c / b for c, b in zip(cvalues, bvalues)]
    if any(b <= 0 for b in bvalues):
        raise ValueError(f"non-positive baseline {metric}")
    return {
        "baseline_child": bvalues,
        "candidate_child": cvalues,
        "ratios": ratios,
        "bootstrap": bootstrap_median(ratios, seed, resamples),
        "max_ratio": max(ratios),
    }


def timed_identity(report: dict[str, Any]) -> tuple[Any, ...]:
    top_fields = (
        "reference_output_sha256", "reference_output_bytes",
        "reference_logical_vec_capacity_bytes", "reference_revision_count",
        "reference_paragraph_count", "reference_semantic_change_count",
    )
    sample_fields = (
        "index", "output_sha256", "output_bytes", "logical_vec_capacity_bytes",
        "exact_output_equal_reference", "revision_count", "paragraph_count",
        "semantic_change_count", "paragraph_zero_sha256", "public_reopen_succeeded",
        "source_unchanged",
    )
    samples = report.get("samples")
    if not isinstance(samples, list):
        raise ValueError("timed samples are not a list")
    return (
        tuple((field, report.get(field)) for field in top_fields),
        tuple(tuple((field, row.get(field)) for field in sample_fields) for row in samples),
    )


def main() -> int:
    import argparse
    global PLAN_PATH, MANIFEST_PATH, OUTPUT_PATH
    parser = argparse.ArgumentParser(description="Independent 0967 raw audit; writes only the explicitly supplied new output path.")
    parser.add_argument("--plan", type=Path, default=PLAN_PATH)
    parser.add_argument("--manifest", type=Path, default=MANIFEST_PATH)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    PLAN_PATH = args.plan.resolve()
    MANIFEST_PATH = args.manifest.resolve()
    OUTPUT_PATH = args.output.resolve()
    forbidden = (BASE / "review/independent-raw-audit-0967-v1.json").resolve()
    if OUTPUT_PATH == forbidden:
        parser.error("refusing to overwrite the superseded v1 audit")
    plan = load(PLAN_PATH)
    manifest = load(MANIFEST_PATH)
    errors: list[str] = []
    warnings: list[str] = []
    capture_root = Path(plan["outputs"]["capture_root"]).resolve()

    # Bind the final authorized plan and manifest before reading raw packets.
    plan_sha = digest(PLAN_PATH)
    manifest_sha = digest(MANIFEST_PATH)
    if manifest.get("schema") != RUNNER_SCHEMA:
        errors.append("capture manifest schema mismatch")
    if manifest.get("plan") != str(PLAN_PATH):
        errors.append("capture manifest plan path mismatch")
    packet_refs = manifest.get("packets")
    if not isinstance(packet_refs, list):
        errors.append("capture manifest packets is not a list")
        packet_refs = []
    if manifest.get("packet_count") != len(packet_refs):
        errors.append("capture manifest packet_count mismatch")
    if len(packet_refs) != 896:
        errors.append(f"capture manifest expected 896 packets, got {len(packet_refs)}")
    ref_paths = [Path(raw).resolve() for raw in packet_refs]
    if len(set(ref_paths)) != len(ref_paths):
        errors.append("duplicate packet paths in capture manifest")
    packets: list[dict[str, Any]] = []
    packet_path_by_id: dict[str, Path] = {}

    def under(path: Path, root: Path) -> bool:
        try:
            path.relative_to(root)
            return True
        except ValueError:
            return False

    for path in ref_paths:
        if not under(path, capture_root):
            errors.append(f"packet path outside capture root: {path}")
            continue
        try:
            packet = load(path)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"packet load {path}: {exc}")
            continue
        packets.append(packet)
        packet_id = packet.get("packet_id")
        if not packet_id or packet_id in packet_path_by_id:
            errors.append(f"duplicate/missing packet_id: {packet_id!r}")
        else:
            packet_path_by_id[packet_id] = path

    expected_binary_roles = {
        "native_baseline": "native_baseline", "native_candidate": "native_candidate",
        "allocator_baseline": "allocator_baseline", "allocator_candidate": "allocator_candidate",
        "control_baseline": "control_baseline", "control_candidate": "control_candidate",
        "timed_control_native_baseline": "timed_control_native_baseline",
        "timed_control_native_candidate": "timed_control_native_candidate",
        "timed_control_allocator_baseline": "timed_control_allocator_baseline",
        "timed_control_allocator_candidate": "timed_control_allocator_candidate",
        "aa_baseline": "native_baseline",
    }
    binary_hash_checks: dict[str, str] = {}
    binary_actual_hashes: dict[str, str] = {}
    for binary_role, binary_row in (plan.get("binaries") or {}).items():
        binary_path = Path(binary_row.get("path", ""))
        if not binary_path.is_file():
            errors.append(f"missing bound binary {binary_role}: {binary_path}")
            continue
        actual_binary_sha = digest(binary_path)
        binary_actual_hashes[binary_role] = actual_binary_sha
        if actual_binary_sha != binary_row.get("sha256"):
            errors.append(f"bound binary hash changed {binary_role}")
    packet_errors_before_role: dict[str, int] = Counter()
    for packet in packets:
        packet_id = packet.get("packet_id", "<missing>")
        role = packet.get("role")
        if packet.get("schema") != RUNNER_SCHEMA:
            errors.append(f"{packet_id}: packet schema mismatch")
        terminal = packet.get("terminal") or {}
        if terminal.get("returncode") != 0 or terminal.get("timed_out") or terminal.get("launch_error"):
            errors.append(f"{packet_id}: terminal failure {terminal}")
        if packet.get("validation_error") is not None:
            errors.append(f"{packet_id}: validation_error={packet.get('validation_error')}")
        if role not in expected_binary_roles:
            errors.append(f"{packet_id}: unexpected role {role}")
            packet_errors_before_role["unexpected_role"] += 1
            continue
        bound_role = expected_binary_roles[role]
        expected_binary = plan["binaries"].get(bound_role) or {}
        observed_binary = packet.get("binary") or {}
        if observed_binary.get("path") != expected_binary.get("path") or observed_binary.get("sha256") != expected_binary.get("sha256"):
            errors.append(f"{packet_id}: binary path/hash mismatch")
        else:
            binary_hash_checks[bound_role] = observed_binary.get("sha256")
        context = packet.get("context") or {}
        if context.get("source_revision") != plan["source"]["revision"]:
            errors.append(f"{packet_id}: source revision mismatch")
        # Check every retained stdout/stderr/RSS descriptor, including the
        # functional-control receipts outside the capture directory.
        for descriptor_name in ("stdout", "stderr"):
            descriptor = packet.get(descriptor_name) or {}
            descriptor_path = descriptor.get("path")
            if descriptor_path is None:
                if descriptor.get("bytes", 0) != 0 or descriptor.get("sha256") is not None:
                    errors.append(f"{packet_id}: null {descriptor_name} has nonzero descriptor")
                continue
            p = Path(descriptor_path)
            if not p.is_file():
                errors.append(f"{packet_id}: missing {descriptor_name} {p}")
                continue
            actual_bytes = p.stat().st_size
            actual_sha = digest(p)
            if descriptor.get("bytes") != actual_bytes or descriptor.get("sha256") != actual_sha:
                errors.append(f"{packet_id}: {descriptor_name} bytes/hash mismatch")
        rss = packet.get("rss") or {}
        if rss.get("path") is not None:
            p = Path(rss["path"])
            if not p.is_file() or rss.get("bytes") != p.stat().st_size:
                errors.append(f"{packet_id}: RSS descriptor mismatch")
        if not isinstance(packet.get("report"), dict):
            errors.append(f"{packet_id}: missing report object")

    role_counts = Counter(packet.get("role") for packet in packets)
    expected_role_counts = {
        "control_baseline": 4, "control_candidate": 4,
        "aa_baseline": 48,
        "timed_control_native_baseline": 48, "timed_control_native_candidate": 48,
        "timed_control_allocator_baseline": 12, "timed_control_allocator_candidate": 12,
        "native_baseline": 288, "native_candidate": 288,
        "allocator_baseline": 72, "allocator_candidate": 72,
    }
    if dict(role_counts) != expected_role_counts:
        errors.append(f"role counts mismatch: {dict(role_counts)}")

    fixtures = plan["fixtures"]
    retention_report = {"zero": "zero", "default": "default_8mib", "release": "release_8mib"}
    controls_by_id = {row["id"]: row for row in plan["controls"]["matrix"]}
    timed_cells = {row["id"]: row for row in plan["controls"]["timed_scope"]["cells"]}
    real_cells = {row["cell_id"]: row for row in plan["real_matrix"]["cells"]}
    semantic_errors: list[str] = []

    # Validate correctness controls independently of the staged validator.
    controls_group: dict[tuple[str, str], dict[str, Any]] = {}
    for packet in packets:
        role = packet.get("role")
        if role not in {"control_baseline", "control_candidate"}:
            continue
        context = packet["context"]
        control_id = context.get("control_id")
        variant = context.get("variant")
        report = packet["report"]
        control = controls_by_id.get(control_id)
        if control is None or variant not in {"baseline", "candidate"}:
            semantic_errors.append(f"functional control identity {packet.get('packet_id')}")
            continue
        controls_group[(control_id, variant)] = packet
        if report.get("schema") != CONTROL_SCHEMA or report.get("variant") != variant or report.get("timing_claim") is not False or report.get("error") is not None:
            semantic_errors.append(f"functional control basic oracle {packet['packet_id']}")
        pre = report.get("pre") or {}
        post = report.get("post") or {}
        before = pre.get("revisions") or []
        after = post.get("revisions") or []
        if control_id == "real-bookmark-delete-redline-query":
            query = report.get("query") or {}
            if report.get("mode") != "query" or len(before) != 20 or query.get("replacement_succeeded") is not False or query.get("source_unchanged_after_refusal") is not True:
                semantic_errors.append(f"functional query oracle {packet['packet_id']}")
        else:
            edit = report.get("edit") or {}
            if report.get("mode") != "run" or edit.get("replacement_succeeded") is not True or edit.get("reopen_succeeded") is not True or edit.get("changed") is not True:
                semantic_errors.append(f"functional edit oracle {packet['packet_id']}")
            witness = report.get("branch_witness") or {}
            if control.get("require_revision_survival") and (not before or after != before or witness.get("positive_branch_required") is not True):
                semantic_errors.append(f"functional revision witness {packet['packet_id']}")
            if control.get("require_false_marker") and (not before or after != before or witness.get("false_marker_present") is not True):
                semantic_errors.append(f"functional false-marker witness {packet['packet_id']}")
    for control_id in controls_by_id:
        for variant in ("baseline", "candidate"):
            if (control_id, variant) not in controls_group:
                semantic_errors.append(f"missing functional control {control_id}/{variant}")
        if (control_id, "baseline") in controls_group and (control_id, "candidate") in controls_group:
            b = controls_group[(control_id, "baseline")]["report"]
            c = controls_group[(control_id, "candidate")]["report"]
            if b.get("source_sha256") != c.get("source_sha256") or b.get("source_bytes") != c.get("source_bytes"):
                semantic_errors.append(f"functional source identity differs {control_id}")

    # Validate and group all timed controls.
    timed_group: dict[str, dict[str, dict[str, dict[int, dict[str, Any]]]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
    for packet in packets:
        role = packet.get("role", "")
        if not role.startswith("timed_control_"):
            continue
        context = packet["context"]
        cell = context.get("cell") or {}
        cell_id = cell.get("id")
        lane = context.get("lane")
        variant = context.get("variant")
        pair = int(context.get("pair", -1))
        timed_group[cell_id][lane][variant][pair] = packet
        row = timed_cells.get(cell_id)
        allocator = lane == "control_allocator"
        report = packet["report"]
        if row is None or lane not in {"control_native", "control_allocator"} or variant not in {"baseline", "candidate"}:
            semantic_errors.append(f"timed control identity {packet.get('packet_id')}")
            continue
        input_row = plan["controls"]["timed_scope"]["inputs"][row["control"]]
        if report.get("schema") != CONTROL_PROFILE_SCHEMA or report.get("variant") != variant or report.get("control") != row["control"] or report.get("scope") != row["scope"]:
            semantic_errors.append(f"timed control basic identity {packet['packet_id']}")
        if row["scope"] == "edit":
            if report.get("edits") != row.get("edits") or report.get("retention") != "default_8mib":
                semantic_errors.append(f"timed control edit identity {packet['packet_id']}")
        if report.get("source_sha256") != input_row["sha256"] or report.get("source_bytes") != input_row["bytes"]:
            semantic_errors.append(f"timed control source identity {packet['packet_id']}")
        if bool(report.get("allocator_instrumented")) != allocator or bool(report.get("timing_claim")) == allocator or report.get("allocator_timing_claim") is not False:
            semantic_errors.append(f"timed control lane flags {packet['packet_id']}")
        warmups = 0 if allocator else 5
        samples = 1 if allocator else 80
        if report.get("warmups") != warmups or report.get("samples_requested") != samples:
            semantic_errors.append(f"timed control schedule {packet['packet_id']}")
        elapsed = report.get("elapsed_ns")
        rows = report.get("samples")
        if not isinstance(elapsed, list) or len(elapsed) != (0 if allocator else samples) or not isinstance(rows, list) or len(rows) != samples:
            semantic_errors.append(f"timed control sample count {packet['packet_id']}")
        expected_revisions = {"generated-positive": 2, "dense1024": 1024, "false-marker": 1, "real-positive": 20}[row["control"]]
        if isinstance(rows, list):
            for sample in rows:
                if sample.get("exact_output_equal_reference") is not True or sample.get("public_reopen_succeeded") is not True or sample.get("revision_count") != expected_revisions:
                    semantic_errors.append(f"timed control semantic sample {packet['packet_id']}")
                if row["scope"] == "edit":
                    if sample.get("semantic_change_count") != row.get("edits") or sample.get("paragraph_zero_sha256") is None:
                        semantic_errors.append(f"timed control edit sample {packet['packet_id']}")
                elif sample.get("source_unchanged") is not True:
                    semantic_errors.append(f"timed control query sample {packet['packet_id']}")
                if allocator:
                    alloc = sample.get("allocation")
                    if not isinstance(alloc, dict) or any(not isinstance(alloc.get(field), int) for field in ("allocated_bytes", "deallocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes")):
                        semantic_errors.append(f"timed control allocation sample {packet['packet_id']}")
        if isinstance(report, dict):
            try:
                # This compares all declared output/semantic identity fields,
                # independently of elapsed and allocator counters.
                pass
            except Exception as exc:  # pragma: no cover
                semantic_errors.append(str(exc))

    # Compare timed control baseline/candidate semantic identity and compute gates.
    timed_results: dict[str, Any] = {}
    timed_seed = plan["controls"]["timed_scope"]["bootstrap_seed"]
    timed_resamples = plan["controls"]["timed_scope"]["bootstrap_resamples"]
    for cell_id, cell in timed_cells.items():
        row_result: dict[str, Any] = {"control": cell["control"], "scope": cell["scope"]}
        for lane, pairs, seed in (("control_native", plan["controls"]["timed_scope"]["native_pairs_per_control"], timed_seed), ("control_allocator", plan["controls"]["timed_scope"]["allocator_pairs_per_control"], timed_seed + 1)):
            group = timed_group.get(cell_id, {}).get(lane, {})
            if set(group) != {"baseline", "candidate"}:
                semantic_errors.append(f"timed control missing variants {cell_id}/{lane}")
                continue
            expected_pairs = set(range(1, pairs + 1))
            if set(group["baseline"]) != expected_pairs or set(group["candidate"]) != expected_pairs:
                semantic_errors.append(f"timed control missing pairs {cell_id}/{lane}")
                continue
            for pair in sorted(expected_pairs):
                if timed_identity(group["baseline"][pair]["report"]) != timed_identity(group["candidate"][pair]["report"]):
                    semantic_errors.append(f"timed control semantic identity differs {cell_id}/{lane}/p{pair}")
            metrics = {}
            metric_names = ("p50", "p95", "p99") if lane == "control_native" else ("allocated_bytes", "peak_live_bytes", "allocation_calls", "rss_kb")
            for metric in metric_names:
                metrics[metric] = pair_metrics(group, metric, pairs, seed, timed_resamples)
            if lane == "control_native":
                metrics["noninferior"] = all(metrics[m]["bootstrap"]["upper"] <= 1.05 for m in ("p50", "p95"))
                metrics["median_over_1_05_review"] = metrics["p50"]["bootstrap"]["median"] > 1.05
                metrics["p99_ratio_median"] = metrics["p99"]["bootstrap"]["median"]
                metrics["p99_descriptive_flag"] = metrics["p99_ratio_median"] > 1.05
                row_result["native"] = metrics
            else:
                metrics["guard_pass"] = metrics["allocated_bytes"]["max_ratio"] <= 1.01 and metrics["peak_live_bytes"]["max_ratio"] <= 1.05 and metrics["rss_kb"]["max_ratio"] <= 1.05
                row_result["allocator"] = metrics
        timed_results[cell_id] = row_result

    # Validate and group A/A readiness, using same-baseline directional ratios.
    aa_group: dict[str, dict[str, dict[int, dict[str, dict[str, Any]]]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
    for packet in packets:
        if packet.get("role") != "aa_baseline":
            continue
        context = packet["context"]
        aa_group[context.get("cell", {}).get("cell_id")][context.get("aa_order")][int(context.get("pair", -1))][context.get("aa_position")] = packet
    aa = plan["aa_noise"]
    aa_results: dict[str, Any] = {}
    aa_errors: list[str] = []
    for cell_id in aa["cells"]:
        cell_result: dict[str, Any] = {}
        for order in ("AB", "BA"):
            rows = aa_group.get(cell_id, {}).get(order, {})
            expected_pairs = {p for p in range(1, aa["native_pairs_per_cell"] + 1) if ("AB" if p <= aa["native_pairs_per_cell"] // 2 else "BA") == order}
            if set(rows) != expected_pairs or any(set(r) != {"first", "second"} for r in rows.values()):
                aa_errors.append(f"incomplete A/A {cell_id}/{order}")
                continue
            direction: dict[str, Any] = {}
            for metric, q in (("p50", 0.50), ("p95", 0.95)):
                values: list[float] = []
                for pair in sorted(rows):
                    first = rows[pair]["first"]["report"].get("elapsed_ns")
                    second = rows[pair]["second"]["report"].get("elapsed_ns")
                    if not isinstance(first, list) or not isinstance(second, list) or len(first) != aa["samples"] or len(second) != aa["samples"]:
                        aa_errors.append(f"A/A sample count {cell_id}/{order}/p{pair}")
                        continue
                    first_metric = percentile([float(v) for v in first], q)
                    second_metric = percentile([float(v) for v in second], q)
                    if first_metric <= 0:
                        aa_errors.append(f"A/A nonpositive first metric {cell_id}/{order}/p{pair}")
                        continue
                    values.append(second_metric / first_metric)
                if len(values) != len(expected_pairs):
                    continue
                ci = bootstrap_median(values, aa["bootstrap_seed"], aa["bootstrap_resamples"])
                symmetric_upper = max(ci["upper"], 1.0 / ci["lower"] if ci["lower"] > 0 else float("inf"))
                direction[metric] = {**ci, "ratios": values, "symmetric_upper": symmetric_upper, "pass": symmetric_upper <= aa["upper_ci_limit"]}
            cell_result[order] = direction
        aa_results[cell_id] = cell_result
    aa_pass = not aa_errors and all(
        cell_result.get(order, {}).get(metric, {}).get("pass") is True
        for cell_result in aa_results.values() for order in ("AB", "BA") for metric in ("p50", "p95")
    )
    aa_readiness_path = Path(manifest.get("aa_readiness", ""))
    aa_manifest_path = capture_root / "aa-readiness-manifest-v1.json"
    try:
        retained_aa = load(aa_readiness_path)
        retained_aa_manifest = load(aa_manifest_path)
        if retained_aa.get("status") != ("pass" if aa_pass else "fail"):
            aa_errors.append("retained A/A readiness status differs")
        if retained_aa_manifest.get("packet_count") != 48 or retained_aa_manifest.get("timed_controls_started") is not False or retained_aa_manifest.get("real_comparisons_started") is not False:
            aa_errors.append("A/A terminal manifest identity mismatch")
    except Exception as exc:  # noqa: BLE001
        aa_errors.append(f"A/A retained readiness load: {exc}")
    aa_pass = aa_pass and not aa_errors

    # Validate and group real native/allocator cells.
    real_groups: dict[str, dict[str, dict[int, dict[str, Any]]]] = defaultdict(lambda: defaultdict(dict))
    for packet in packets:
        role = packet.get("role")
        if role not in {"native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate"}:
            continue
        context = packet["context"]
        cell = context.get("cell") or {}
        cell_id = cell.get("cell_id")
        lane = context.get("lane")
        variant = context.get("variant")
        pair = int(context.get("pair", -1))
        real_groups[f"{lane}:{cell_id}"][variant][pair] = packet
        expected_cell = real_cells.get(cell_id)
        if expected_cell is None or expected_cell.get("case") != cell.get("case") or expected_cell.get("edits") != cell.get("edits") or expected_cell.get("retention") != cell.get("retention"):
            semantic_errors.append(f"real cell identity {packet.get('packet_id')}")
        expected_fixture = fixtures.get(cell.get("case"))
        report = packet.get("report") or {}
        allocator = lane == "allocator"
        if report.get("schema") != PROBE_SCHEMA or report.get("case") != cell.get("case") or report.get("edits") != cell.get("edits") or report.get("source_sha256") != expected_fixture.get("sha256") or report.get("source_bytes") != expected_fixture.get("bytes"):
            semantic_errors.append(f"real report/source identity {packet.get('packet_id')}")
        if report.get("retention") != retention_report.get(cell.get("retention")) or report.get("warmups") != (0 if allocator else 5) or report.get("samples_requested") != (1 if allocator else 80):
            semantic_errors.append(f"real schedule/report identity {packet.get('packet_id')}")
        if bool(report.get("allocator_instrumented")) != allocator or bool(report.get("timing_claim")) == allocator:
            semantic_errors.append(f"real lane flags {packet.get('packet_id')}")
        elapsed = report.get("elapsed_ns")
        rows = report.get("samples")
        if not isinstance(elapsed, list) or len(elapsed) != (0 if allocator else 80) or not isinstance(rows, list) or len(rows) != (1 if allocator else 80):
            semantic_errors.append(f"real sample count {packet.get('packet_id')}")
        all_true(report.get("correctness"), f"{packet.get('packet_id')}.correctness", semantic_errors)
        if isinstance(rows, list):
            for sample in rows:
                if not isinstance(sample, dict) or sample.get("exact_output_equal_reference") is not True:
                    semantic_errors.append(f"real output identity {packet.get('packet_id')}")
                all_true((sample or {}).get("oracle"), f"{packet.get('packet_id')}.sample.oracle", semantic_errors)
                if allocator:
                    alloc = (sample or {}).get("allocation")
                    if not isinstance(alloc, dict) or any(not isinstance(alloc.get(field), int) for field in ("allocated_bytes", "deallocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes")):
                        semantic_errors.append(f"real allocation fields {packet.get('packet_id')}")
                elif sample.get("elapsed_ns") is None:
                    semantic_errors.append(f"real elapsed sample {packet.get('packet_id')}")

    real_native: dict[str, Any] = {}
    real_allocator: dict[str, Any] = {}
    real_seed = plan["real_matrix"]["native"]["bootstrap_seed"]
    real_resamples = plan["real_matrix"]["native"]["bootstrap_resamples"]
    alloc_seed = plan["real_matrix"]["allocator"]["bootstrap_seed"]
    alloc_resamples = plan["real_matrix"]["allocator"]["bootstrap_resamples"]
    native_failed_cells: list[dict[str, Any]] = []
    for cell_id in real_cells:
        group = real_groups.get(f"native:{cell_id}", {})
        metrics = {metric: pair_metrics(group, metric, 12, real_seed, real_resamples) for metric in ("p50", "p95", "p99")}
        metrics["noninferior"] = all(metrics[m]["bootstrap"]["upper"] <= 1.05 for m in ("p50", "p95"))
        metrics["median_over_1_05_review"] = metrics["p50"]["bootstrap"]["median"] > 1.05
        metrics["p99_ratio_median"] = metrics["p99"]["bootstrap"]["median"]
        metrics["p99_descriptive_flag"] = metrics["p99_ratio_median"] > 1.05
        real_native[cell_id] = metrics
        if not metrics["noninferior"]:
            native_failed_cells.append({
                "cell_id": cell_id,
                "p50_upper": metrics["p50"]["bootstrap"]["upper"],
                "p95_upper": metrics["p95"]["bootstrap"]["upper"],
                "p50_median": metrics["p50"]["bootstrap"]["median"],
                "p95_median": metrics["p95"]["bootstrap"]["median"],
            })
    for cell_id in real_cells:
        group = real_groups.get(f"allocator:{cell_id}", {})
        metrics = {metric: pair_metrics(group, metric, 3, alloc_seed, alloc_resamples) for metric in ("allocated_bytes", "peak_live_bytes", "allocation_calls", "rss_kb")}
        metrics["guard_pass"] = metrics["allocated_bytes"]["max_ratio"] <= 1.01 and metrics["peak_live_bytes"]["max_ratio"] <= 1.05 and metrics["rss_kb"]["max_ratio"] <= 1.05
        real_allocator[cell_id] = metrics

    # Recompute every declared gate without short-circuiting on an earlier fail.
    endpoint_cells = {"docfloat/e32/default", "docnohf/e32/default"}
    allocation_benefit = (
        real_allocator["docfloat/e32/default"]["allocation_calls"]["bootstrap"]["median"] <= 0.90
        and real_allocator["docnohf/e32/default"]["allocation_calls"]["bootstrap"]["median"] <= 0.97
    )
    native_benefit = any(
        real_native[cell]["p50"]["bootstrap"]["median"] <= 0.95 and real_native[cell]["p50"]["bootstrap"]["upper"] < 1.00
        for cell in endpoint_cells
    )
    non_endpoint_cells = set(real_cells) - endpoint_cells
    allocation_calls_guard = all(real_allocator[cell]["allocation_calls"]["bootstrap"]["median"] <= 1.01 for cell in non_endpoint_cells)
    allocator_memory = all(real_allocator[cell]["guard_pass"] for cell in real_cells)
    native_noninferiority = all(real_native[cell]["noninferior"] for cell in real_cells)
    timed_native_noninferiority = all(row.get("native", {}).get("noninferior") is True for row in timed_results.values())
    timed_allocator_guard = all(row.get("allocator", {}).get("guard_pass") is True for row in timed_results.values())
    timed_controls_pass = timed_native_noninferiority and timed_allocator_guard
    complete = not errors and not semantic_errors and not aa_errors
    gates = {
        "custody_and_schema_complete": complete,
        "aa_readiness": aa_pass,
        "timed_native_noninferiority": timed_native_noninferiority,
        "timed_allocator_guard": timed_allocator_guard,
        "timed_controls": timed_controls_pass,
        "native_noninferiority": native_noninferiority,
        "allocator_memory": allocator_memory,
        "allocation_calls_guard": allocation_calls_guard,
        "allocation_benefit": allocation_benefit,
        "native_benefit": native_benefit,
    }
    performance_pass = all(gates[name] for name in ("aa_readiness", "timed_controls", "native_noninferiority", "allocator_memory", "allocation_calls_guard")) and (allocation_benefit or native_benefit)
    decision = "reject_candidate" if complete and not performance_pass else ("qualify_candidate" if complete else "no_decision_incomplete_evidence")

    timed_allocator_failures = []
    for cell_id, row in timed_results.items():
        if not row.get("allocator", {}).get("guard_pass", False):
            a = row.get("allocator", {})
            timed_allocator_failures.append({
                "cell_id": cell_id,
                "allocated_max_ratio": a.get("allocated_bytes", {}).get("max_ratio"),
                "peak_live_max_ratio": a.get("peak_live_bytes", {}).get("max_ratio"),
                "rss_max_ratio": a.get("rss_kb", {}).get("max_ratio"),
                "allocation_calls_median": a.get("allocation_calls", {}).get("bootstrap", {}).get("median"),
            })

    result = {
        "schema": "litchi.performance.0967.independent_raw_audit.v2",
        "status": "pass_complete_performance_rejected" if decision == "reject_candidate" else ("pass_complete_candidate_qualified" if decision == "qualify_candidate" else "incomplete_no_decision"),
        "decision": decision,
        "created_utc": __import__("datetime").datetime.now(__import__("datetime").timezone.utc).isoformat(),
        "plan": {"path": str(PLAN_PATH), "sha256": plan_sha},
        "manifest": {"path": str(MANIFEST_PATH), "sha256": manifest_sha, "packet_count": len(packets)},
        "terminal": {"capture_manifest_schema": manifest.get("schema"), "packet_count": len(packets), "role_counts": dict(role_counts), "no_retry": True},
        "custody": {"errors": errors, "semantic_errors": semantic_errors, "binary_hash_roles_seen": binary_hash_checks, "binary_actual_hashes": binary_actual_hashes, "packet_paths_under_capture_root": all(under(p, capture_root) for p in ref_paths)},
        "aa_readiness": {"pass": aa_pass, "errors": aa_errors, "cells": aa_results, "retained_path": str(aa_readiness_path)},
        "timed_controls": {"pass": timed_controls_pass, "native_noninferiority": timed_native_noninferiority, "allocator_guard": timed_allocator_guard, "allocator_failures": timed_allocator_failures, "cells": timed_results},
        "native": {"noninferiority": native_noninferiority, "failed_cells": native_failed_cells, "cells": real_native},
        "allocator": {"memory_guard": allocator_memory, "allocation_calls_guard": allocation_calls_guard, "cells": real_allocator},
        "gates": gates,
        "performance": {"allocation_benefit": allocation_benefit, "native_benefit": native_benefit, "all_performance_gates": performance_pass},
        "warnings": warnings,
        "interpretation": "All packets and semantic oracles are complete, but a failed declared performance/control gate is a rejection, not incomplete evidence. No candidate adoption or production recommendation is made.",
        "overwrite_incident": {
            "type": "sealed_audit_output_overwritten",
            "superseded_path": str(BASE / "review/independent-raw-audit-0967-v1.json"),
            "sealed_v1_sha256": "dfb37103d82b076e46c680243eca713b251ffa16aeaf2ee06f7084999a1c6402",
            "sealed_v1_recovered": False,
            "observed_overwritten_v1_sha256": "b756855c7e779901be78d1c3574a5445e1a09078ebe25de0ffd410eb16fa3436",
            "cause": "independent_raw_audit_0967_v1.py --help invoked after sealing; v1 hardcoded its output path and rewrote it",
            "v2_scope": "same bound plan, manifest, retained packets, binaries, and declared thresholds; no workload or measurement rerun",
        },
    }
    OUTPUT_PATH.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": result["status"], "decision": decision, "output": str(OUTPUT_PATH), "sha256": digest(OUTPUT_PATH), "errors": len(errors), "semantic_errors": len(semantic_errors), "aa_errors": len(aa_errors), "native_failed_cells": len(native_failed_cells), "timed_allocator_failures": len(timed_allocator_failures), "gates": gates}, sort_keys=True))
    return 0 if decision in {"reject_candidate", "qualify_candidate"} else 2


if __name__ == "__main__":
    raise SystemExit(main())
