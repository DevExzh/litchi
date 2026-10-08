#!/usr/bin/env python3
"""Serial 0967 packet runner.

The default operation is a fail-closed plan/binding check. Capture requires an
explicitly authorized plan with all executable hashes and a sealed timing
control binding. It performs no retries, keeps terminal/stdout/stderr packets,
and never starts Cargo.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SCHEMA = "litchi.performance.0967.runner.v1"
PROBE_SCHEMA = "litchi.performance.0917.doc_retention.v1"
CONTROL_SCHEMA = "litchi.performance.0967.doc_revision_positive_control.v1"
CONTROL_PROFILE_BUILD_STATUSES = {"passed", "sealed", "matched_ordinary_native_and_allocator_control_builds_passed"}
MAX_OUTPUT_BYTES = 8 * 1024 * 1024
TIMEOUT_SECONDS = 900.0
EXPECTED_BINARY_ROLES = {
    "native_baseline": "ordinary_uninstrumented",
    "native_candidate": "ordinary_uninstrumented",
    "allocator_baseline": "counting_allocator",
    "allocator_candidate": "counting_allocator",
    "control_baseline": "correctness_control",
    "control_candidate": "correctness_control",
    "timed_control_native_baseline": "control_profile_native",
    "timed_control_native_candidate": "control_profile_native",
    "timed_control_allocator_baseline": "control_profile_allocator",
    "timed_control_allocator_candidate": "control_profile_allocator",
}


def utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def percentile(values: list[float], quantile: float) -> float:
    if not values:
        raise ValueError("percentile of empty values")
    ordered = sorted(float(value) for value in values)
    if len(ordered) == 1:
        return ordered[0]
    position = (len(ordered) - 1) * quantile
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def bootstrap_median(values: list[float], seed: int, resamples: int) -> dict[str, float]:
    if not values:
        raise ValueError("bootstrap of empty values")
    import random

    rng = random.Random(seed)
    medians = []
    for _ in range(resamples):
        sample = [values[rng.randrange(len(values))] for _ in values]
        medians.append(percentile(sample, 0.5))
    return {
        "median": percentile(values, 0.5),
        "lower": percentile(medians, 0.025),
        "upper": percentile(medians, 0.975),
    }


def load(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"{path}: expected object")
    return value


def fail(message: str) -> None:
    raise ValueError(message)


def check_hash(path: Path, expected: str, label: str) -> None:
    if not path.is_file():
        fail(f"{label}: missing {path}")
    if expected and sha(path) != expected:
        fail(f"{label}: SHA-256 mismatch for {path}")


def all_true(value: Any, path: str = "correctness") -> None:
    if not isinstance(value, dict):
        fail(f"{path}: expected object")
    bad = [key for key, item in value.items() if isinstance(item, bool) and not item]
    if bad:
        fail(f"{path}: failed gates {bad}")


def check_plan(plan_path: Path, *, allow_unbound: bool = False, for_capture: bool = False) -> dict[str, Any]:
    plan = load(plan_path)
    if plan.get("schema") != "litchi.performance.0967.plan.v1":
        fail("unexpected plan schema")
    if plan.get("source", {}).get("revision") != "893072657ad8b1d726bfce351b5ec2e98f084e94":
        fail("source revision is not the sealed 0967 revision")
    protocol = plan.get("protocol_document") or {}
    check_hash(Path(protocol.get("path", "")), protocol.get("sha256", ""), "protocol document")
    if plan.get("build_policy", {}).get("cargo_jobs") != 2 or plan.get("build_policy", {}).get("concurrency") != 1:
        fail("build policy drifted from serial cargo-process / -j2 policy")
    if plan.get("build_policy", {}).get("rustflags") != "":
        fail("RUSTFLAGS must be empty for matched uninstrumented binaries")
    control_profile_receipt = plan.get("build_policy", {}).get("control_profile_receipt")
    if control_profile_receipt:
        receipt_path = Path(control_profile_receipt.get("path", ""))
        check_hash(receipt_path, control_profile_receipt.get("sha256", ""), "control profile build receipt")
        receipt = load(receipt_path)
        if receipt.get("status") not in CONTROL_PROFILE_BUILD_STATUSES:
            fail("control profile build receipt is not sealed/passed")
    source = plan.get("source") or {}
    for tool_name, row in (plan.get("tools") or {}).items():
        tool_path = Path(row.get("path", ""))
        check_hash(tool_path, row.get("sha256", ""), f"bound tool {tool_name}")
    for role in ("baseline", "candidate"):
        row = source.get(role) or {}
        manifest = Path(row.get("manifest", ""))
        check_hash(manifest, row.get("manifest_sha256", ""), f"source.{role}.manifest")
        doc = load(manifest)
        if doc.get("source_revision") != plan["source"]["revision"]:
            fail(f"{role} source manifest revision mismatch")
        if role == "candidate" and doc.get("status") != "sealed_for_matched_build":
            fail("candidate source manifest is not sealed")
        if role == "baseline" and not doc.get("status"):
            fail("baseline source manifest has no status")
    for key in ("quality_receipt", "semantic_review"):
        row = source.get(key) or {}
        receipt = Path(row.get("path", ""))
        check_hash(receipt, row.get("sha256", ""), f"source.{key}")
        doc = load(receipt)
        if doc.get("status") not in {"passed", "passed_with_retained_baseline_failure"}:
            fail(f"{key}: status is not accepted")
    harness = plan.get("harness") or {}
    for field in ("candidate_files", "baseline_files", "candidate_timed_files", "baseline_timed_files"):
        for row in harness.get(field, []):
            root = Path(source["candidate" if field.startswith("candidate") else "baseline"]["root"])
            check_hash(root / row["path"], row["sha256"], f"harness {field}")
    fixtures = plan.get("fixtures") or {}
    if set(fixtures) != {"docnohf", "docfloat", "real-bookmark-delete-redline-query"}:
        fail("fixture set drifted")
    for case, row in fixtures.items():
        path = Path(row["path"])
        check_hash(path, row["sha256"], f"fixture {case}")
        if path.stat().st_size != row["bytes"]:
            fail(f"fixture {case}: byte count changed")
    matrix = plan.get("real_matrix") or {}
    cells = matrix.get("cells") or []
    if len(cells) != 24 or len({row.get("cell_id") for row in cells}) != 24:
        fail("real matrix must contain 24 unique cells")
    native = matrix.get("native") or {}
    if native.get("pairs_per_cell") != 12 or native.get("warmups") != 5 or native.get("samples") != 80:
        fail("native schedule drifted")
    expected_p99_descriptor = {"ratio_median_field": "p99_ratio_median", "flag_field": "p99_descriptive_flag", "flag_if_median_gt": 1.05, "report_only": True}
    if native.get("p99_descriptive") != expected_p99_descriptor:
        fail("native p99 descriptive report field drifted")
    allocator = matrix.get("allocator") or {}
    if allocator.get("pairs_per_cell") != 3 or allocator.get("warmups") != 0 or allocator.get("samples") != 1:
        fail("allocator schedule drifted")
    if native.get("relative_ratio") != "candidate_over_baseline" or allocator.get("relative_ratio") != "candidate_over_baseline":
        fail("relative orientation is not explicit candidate_over_baseline")
    controls = plan.get("controls") or {}
    if controls.get("scenario_audit", {}).get("status_required") != "pass":
        fail("scenario audit is not required")
    scenario_tool = Path(controls.get("scenario_audit", {}).get("path", ""))
    check_hash(scenario_tool, controls.get("scenario_audit", {}).get("sha256", ""), "scenario audit tool")
    if controls.get("probe_receipt"):
        probe_receipt = Path(controls["probe_receipt"].get("path", ""))
        check_hash(probe_receipt, controls["probe_receipt"].get("sha256", ""), "functional control receipt")
    timed = controls.get("timed_scope") or {}
    if for_capture:
        if plan.get("status") != "capture_authorized":
            fail("capture requires plan status capture_authorized")
        if plan.get("host_state", {}).get("status") != "bound":
            fail("capture requires root-bound before/after host state")
        if controls.get("status") != "bound":
            fail("capture requires sealed correctness controls")
        if timed.get("status") != "bound":
            fail("capture requires a timing-capable control binding")
    if timed.get("schema") != "litchi.performance.0967.control_profile.v1":
        fail("timed control schema is not bound")
    if timed.get("p99_descriptive") != expected_p99_descriptor:
        fail("timed-control p99 descriptive report field drifted")
    timed_root = Path(timed.get("harness_root", ""))
    for row in timed.get("harness_files", []):
        check_hash(timed_root / row["path"], row.get("sha256", ""), "timed control harness")
    expected_timed_cells = {"generated-positive-e1-default", "dense1024-e32-default", "false-marker-e1-default", "false-marker-e32-default"}
    if {row.get("id") for row in timed.get("cells", [])} != expected_timed_cells:
        fail("timed control cell matrix drifted")
    for control, row in (timed.get("inputs") or {}).items():
        if row.get("available") is not True:
            if for_capture:
                fail(f"timed control input is unavailable: {control}")
            continue
        input_path = Path(row.get("path", ""))
        check_hash(input_path, row.get("sha256", ""), f"timed control input {control}")
        if input_path.stat().st_size != row.get("bytes"):
            fail(f"timed control input byte count changed: {control}")
    if not for_capture and not allow_unbound:
        if plan.get("status") != "capture_authorized":
            fail("plan is prospective; use --allow-unbound for preflight")
    aa = plan.get("aa_noise") or {}
    if aa.get("coverage") != "endpoint_cells_only" or aa.get("same_binary") != "baseline" or aa.get("order") != ["AB", "BA"]:
        fail("A/A readiness scope drifted")
    if aa.get("cells") != ["docnohf/e32/default", "docfloat/e32/default"]:
        fail("A/A endpoint coverage drifted")
    if aa.get("native_pairs_per_cell") != 12 or aa.get("warmups") != 5 or aa.get("samples") != 80:
        fail("A/A schedule drifted")
    if aa.get("symmetric_ratio") != "max(r,1/r)":
        fail("A/A symmetric ratio is not bound")
    if aa.get("upper_ci_limit") != 1.05:
        fail("A/A upper CI threshold drifted")
    if not isinstance(aa.get("bootstrap_seed"), int) or aa.get("bootstrap_seed") <= 0:
        fail("A/A bootstrap seed is not bound")
    if not isinstance(aa.get("bootstrap_resamples"), int) or aa.get("bootstrap_resamples") < 1000:
        fail("A/A bootstrap resamples are not bound")
    binaries = plan.get("binaries") or {}
    for role, expected_kind in EXPECTED_BINARY_ROLES.items():
        row = binaries.get(role) or {}
        if row.get("kind") != expected_kind:
            fail(f"binary {role}: kind mismatch")
        path = Path(row.get("path", ""))
        digest = row.get("sha256")
        if not digest:
            if not allow_unbound:
                fail(f"binary {role}: SHA-256 is not bound")
        else:
            check_hash(path, digest, f"binary {role}")
    return plan


def read_json_line(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    if len(data) > MAX_OUTPUT_BYTES:
        fail(f"stdout exceeds {MAX_OUTPUT_BYTES} bytes: {path}")
    lines = [line for line in data.decode("utf-8", errors="strict").splitlines() if line.strip()]
    if not lines:
        fail(f"no JSON report on stdout: {path}")
    try:
        value = json.loads(lines[-1])
    except json.JSONDecodeError as exc:
        fail(f"invalid final JSON report {path}: {exc}")
    if not isinstance(value, dict):
        fail(f"report {path}: expected JSON object")
    return value


def check_report(report: dict[str, Any], *, case: str, fixture: dict[str, Any], edits: int, retention: str, allocator: bool, warmups: int, samples: int) -> None:
    if report.get("schema") != PROBE_SCHEMA:
        fail(f"{case}: report schema mismatch")
    if report.get("case") != case or report.get("edits") != edits:
        fail(f"{case}: report cell mismatch")
    expected_retention = {"zero": "zero", "default": "default_8mib", "release": "release_8mib"}[retention]
    if report.get("retention") != expected_retention:
        fail(f"{case}: retention mismatch")
    if report.get("source_sha256") != fixture["sha256"] or report.get("source_bytes") != fixture["bytes"]:
        fail(f"{case}: source fixture identity mismatch")
    if report.get("warmups") != warmups or report.get("samples_requested") != samples:
        fail(f"{case}: warmup/sample mismatch")
    if bool(report.get("allocator_instrumented")) != allocator or bool(report.get("timing_claim")) == allocator:
        fail(f"{case}: lane instrumentation/timing flags mismatch")
    all_true(report.get("correctness"))
    elapsed = report.get("elapsed_ns")
    if not isinstance(elapsed, list) or (len(elapsed) != samples if not allocator else len(elapsed) != 0):
        fail(f"{case}: elapsed sample count mismatch")
    samples_rows = report.get("samples")
    if not isinstance(samples_rows, list) or len(samples_rows) != samples:
        fail(f"{case}: sample row count mismatch")
    for item in samples_rows:
        if not isinstance(item, dict) or item.get("exact_output_equal_reference") is not True:
            fail(f"{case}: sample output oracle failed")
        all_true(item.get("oracle"), f"{case}.sample.oracle")
        if allocator:
            allocation = item.get("allocation")
            if not isinstance(allocation, dict):
                fail(f"{case}: allocator sample has no allocation metrics")
            for field in ("allocated_bytes", "deallocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes"):
                if not isinstance(allocation.get(field), int):
                    fail(f"{case}: missing allocation field {field}")
        elif item.get("elapsed_ns") is None:
            fail(f"{case}: native sample has no elapsed_ns")
    route_outputs = report.get("route_outputs")
    if not isinstance(route_outputs, list) or len(route_outputs) != 3:
        fail(f"{case}: route output set incomplete")
    for route in route_outputs:
        all_true(route.get("oracle"), f"{case}.route.oracle")


def kill_process(proc: subprocess.Popen[bytes]) -> None:
    try:
        os.killpg(proc.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass


def launch_packet(*, executable: Path, args: list[str], role: str, packet_dir: Path, context: dict[str, Any], validator: Any) -> dict[str, Any]:
    packet_dir.mkdir(parents=True, exist_ok=True)
    stem = context["packet_id"]
    stdout_path = packet_dir / f"{stem}.stdout"
    stderr_path = packet_dir / f"{stem}.stderr"
    rss_path = packet_dir / f"{stem}.rss"
    command = ["/usr/bin/time", "-f", "%M", "-o", str(rss_path), str(executable), *args]
    started = utc()
    start_ns = time.monotonic_ns()
    terminal = {"returncode": None, "timed_out": False, "launch_error": None}
    try:
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            proc = subprocess.Popen(command, stdout=stdout, stderr=stderr, start_new_session=True)
            try:
                proc.wait(timeout=TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                terminal["timed_out"] = True
                kill_process(proc)
                proc.wait()
            terminal["returncode"] = proc.returncode
    except OSError as exc:
        terminal["launch_error"] = str(exc)
    elapsed_wall_ns = time.monotonic_ns() - start_ns
    packet: dict[str, Any] = {
        "schema": SCHEMA,
        "packet_id": stem,
        "created_utc": started,
        "role": role,
        "command": command,
        "binary": {"path": str(executable), "sha256": sha(executable) if executable.is_file() else None},
        "context": context,
        "terminal": terminal,
        "wall_ns_excluded_from_metric": elapsed_wall_ns,
        "stdout": {"path": str(stdout_path), "bytes": stdout_path.stat().st_size if stdout_path.exists() else 0, "sha256": sha(stdout_path) if stdout_path.exists() else None},
        "stderr": {"path": str(stderr_path), "bytes": stderr_path.stat().st_size if stderr_path.exists() else 0, "sha256": sha(stderr_path) if stderr_path.exists() else None},
        "rss": {"path": str(rss_path), "bytes": rss_path.stat().st_size if rss_path.exists() else 0, "kb": None},
        "report": None,
        "validation_error": None,
    }
    if rss_path.exists():
        try:
            packet["rss"]["kb"] = int(rss_path.read_text(encoding="ascii").strip())
        except (ValueError, UnicodeDecodeError):
            packet["validation_error"] = "invalid /usr/bin/time RSS output"
    if terminal["returncode"] == 0 and not terminal["timed_out"] and terminal["launch_error"] is None:
        try:
            report = read_json_line(stdout_path)
            validator(report)
            packet["report"] = report
        except (OSError, UnicodeError, ValueError) as exc:
            packet["validation_error"] = str(exc)
    elif terminal["returncode"] is None:
        packet["validation_error"] = terminal["launch_error"] or "child did not launch"
    elif terminal["timed_out"]:
        packet["validation_error"] = "timeout; no retry permitted"
    else:
        packet["validation_error"] = f"child exit status {terminal['returncode']}"
    return packet


def write_packet(packet_dir: Path, packet: dict[str, Any]) -> None:
    packet_dir.mkdir(parents=True, exist_ok=True)
    path = packet_dir / f"{packet['packet_id']}.json"
    path.write_text(json.dumps(packet, indent=2) + "\n", encoding="utf-8")


def run_real(plan: dict[str, Any], lane: str, packet_root: Path) -> list[dict[str, Any]]:
    matrix = plan["real_matrix"]
    schedule = matrix[lane]["schedule"]
    pairs = matrix[lane]["pairs_per_cell"]
    warmups = matrix[lane]["warmups"]
    samples = matrix[lane]["samples"]
    packets: list[dict[str, Any]] = []
    fixtures = plan["fixtures"]
    for cell in matrix["cells"]:
        fixture = fixtures[cell["case"]]
        for pair in schedule:
            order = ["baseline", "candidate"] if pair["relative_order"] == "baseline_then_candidate" else ["candidate", "baseline"]
            for variant in order:
                role = f"{lane}_{variant}"
                binary = Path(plan["binaries"][f"{lane}_{variant}"]["path"])
                args = ["--case", cell["case"], "--input", fixture["path"], "--edits", str(cell["edits"]), "--retention", cell["retention"], "--warmup", str(warmups), "--samples", str(samples)]
                packet_id = f"{lane}-{cell['cell_id'].replace('/', '_')}-p{pair['pair']:02d}-{variant}"
                context = {"packet_id": packet_id, "lane": lane, "variant": variant, "cell": cell, "pair": pair["pair"], "relative_order": pair["relative_order"], "source_revision": plan["source"]["revision"], "fixture_sha256": fixture["sha256"], "warmups": warmups, "samples": samples}
                validator = lambda report, case=cell["case"], fixture=fixture, edits=cell["edits"], retention=cell["retention"], allocator=lane == "allocator", warmups=warmups, samples=samples: check_report(report, case=case, fixture=fixture, edits=edits, retention=retention, allocator=allocator, warmups=warmups, samples=samples)
                packet = launch_packet(executable=binary, args=args, role=role, packet_dir=packet_root / lane, context=context, validator=validator)
                write_packet(packet_root / lane, packet)
                packets.append(packet)
    return packets


def run_aa(plan: dict[str, Any], packet_root: Path) -> list[dict[str, Any]]:
    """Run the predeclared same-baseline AB/BA endpoint readiness cells."""
    aa = plan["aa_noise"]
    packets: list[dict[str, Any]] = []
    binary = Path(plan["binaries"]["native_baseline"]["path"])
    fixtures = plan["fixtures"]
    for cell_id in aa["cells"]:
        case, edit_part, retention = cell_id.split("/")
        fixture = fixtures[case]
        edits = int(edit_part[1:])
        for pair in range(1, aa["native_pairs_per_cell"] + 1):
            orientation = "AB" if pair <= aa["native_pairs_per_cell"] // 2 else "BA"
            for position in ("first", "second"):
                packet_id = f"aa-{cell_id.replace('/', '_')}-p{pair:02d}-{position}"
                context = {"packet_id": packet_id, "lane": "aa", "variant": "baseline", "cell": {"cell_id": cell_id, "case": case, "edits": edits, "retention": retention}, "pair": pair, "aa_order": orientation, "aa_position": position, "source_revision": plan["source"]["revision"], "fixture_sha256": fixture["sha256"], "warmups": aa["warmups"], "samples": aa["samples"]}
                args = ["--case", case, "--input", fixture["path"], "--edits", str(edits), "--retention", retention, "--warmup", str(aa["warmups"]), "--samples", str(aa["samples"])]
                validator = lambda report, case=case, fixture=fixture, edits=edits, retention=retention, warmups=aa["warmups"], samples=aa["samples"]: check_report(report, case=case, fixture=fixture, edits=edits, retention=retention, allocator=False, warmups=warmups, samples=samples)
                packet = launch_packet(executable=binary, args=args, role="aa_baseline", packet_dir=packet_root / "aa", context=context, validator=validator)
                write_packet(packet_root / "aa", packet)
                packets.append(packet)
    return packets


def aa_readiness(plan: dict[str, Any], packets: list[dict[str, Any]]) -> dict[str, Any]:
    """Reduce the frozen A/A stage before any timed or real comparison.

    The same baseline executable runs in AB and BA order.  Every direction
    must have all paired children and both p50/p95 symmetric bootstrap upper
    bounds at or below the frozen 1.05 limit.  This function deliberately
    reads only A/A packets, so a readiness failure cannot be masked by later
    candidate packets.
    """
    aa = plan["aa_noise"]
    expected_pairs = set(range(1, aa["native_pairs_per_cell"] + 1))
    grouped: dict[str, dict[str, dict[int, dict[str, dict[str, Any]]]]] = {}
    errors: list[str] = []
    for packet in packets:
        packet_id = packet.get("packet_id")
        if packet.get("role") != "aa_baseline":
            errors.append(f"{packet_id}: A/A role mismatch")
            continue
        if packet.get("validation_error") is not None:
            errors.append(f"{packet_id}: {packet['validation_error']}")
            continue
        context = packet.get("context") or {}
        cell = (context.get("cell") or {}).get("cell_id")
        order = context.get("aa_order")
        position = context.get("aa_position")
        try:
            pair = int(context.get("pair", -1))
        except (TypeError, ValueError):
            pair = -1
        if cell not in aa["cells"] or order not in {"AB", "BA"} or position not in {"first", "second"} or pair not in expected_pairs:
            errors.append(f"{packet_id}: invalid A/A context")
            continue
        expected_order = "AB" if pair <= aa["native_pairs_per_cell"] // 2 else "BA"
        if order != expected_order:
            errors.append(f"{packet_id}: A/A pair order mismatch")
            continue
        grouped.setdefault(cell, {}).setdefault(order, {}).setdefault(pair, {})[position] = packet
    if set(grouped) != set(aa["cells"]):
        errors.append("A/A cell coverage is incomplete")
    result: dict[str, Any] = {
        "schema": "litchi.performance.0967.aa_readiness.v1",
        "status": "fail" if errors else "pass",
        "threshold": aa["upper_ci_limit"],
        "bootstrap_seed": aa["bootstrap_seed"],
        "bootstrap_resamples": aa["bootstrap_resamples"],
        "pairs_per_cell": aa["native_pairs_per_cell"],
        "errors": errors,
        "cells": {},
    }
    if errors:
        return result
    for cell in aa["cells"]:
        cell_result: dict[str, Any] = {}
        for order in ("AB", "BA"):
            rows = grouped[cell].get(order, {})
            expected_order_pairs = {pair for pair in expected_pairs if ("AB" if pair <= aa["native_pairs_per_cell"] // 2 else "BA") == order}
            if set(rows) != expected_order_pairs or any(set(item) != {"first", "second"} for item in rows.values()):
                result["errors"].append(f"A/A {cell}/{order}: incomplete paired children")
                continue
            direction: dict[str, Any] = {}
            for metric, quantile in (("p50", 0.50), ("p95", 0.95)):
                values: list[float] = []
                for pair in sorted(rows):
                    first = (rows[pair]["first"].get("report") or {}).get("elapsed_ns")
                    second = (rows[pair]["second"].get("report") or {}).get("elapsed_ns")
                    if not isinstance(first, list) or not isinstance(second, list) or len(first) != aa["samples"] or len(second) != aa["samples"]:
                        result["errors"].append(f"A/A {cell}/{order}/pair-{pair}: elapsed samples incomplete")
                        break
                    baseline_metric = percentile([float(item) for item in first], quantile)
                    repeated_metric = percentile([float(item) for item in second], quantile)
                    if baseline_metric <= 0:
                        result["errors"].append(f"A/A {cell}/{order}/pair-{pair}: non-positive baseline metric")
                        break
                    values.append(repeated_metric / baseline_metric)
                if len(values) != len(expected_order_pairs):
                    continue
                interval = bootstrap_median(values, aa["bootstrap_seed"], aa["bootstrap_resamples"])
                symmetric_upper = max(interval["upper"], 1.0 / interval["lower"] if interval["lower"] > 0 else float("inf"))
                direction[metric] = {**interval, "ratios": values, "symmetric_upper": symmetric_upper, "pass": symmetric_upper <= aa["upper_ci_limit"]}
            cell_result[order] = direction
        result["cells"][cell] = cell_result
    if result["errors"]:
        result["status"] = "fail"
    else:
        result["status"] = "pass" if all(
            direction.get(metric, {}).get("pass") is True
            for cell_result in result["cells"].values()
            for direction in cell_result.values()
            for metric in ("p50", "p95")
        ) else "fail"
        if result["status"] == "fail":
            result["errors"].append("A/A symmetric bootstrap upper CI exceeded 1.05")
    return result


def write_aa_readiness(packet_root: Path, plan: dict[str, Any], packets: list[dict[str, Any]], readiness: dict[str, Any], terminal_status: str) -> None:
    packet_paths = [str(packet_root / "aa" / f"{packet['packet_id']}.json") for packet in packets]
    manifest = {
        "schema": "litchi.performance.0967.aa_readiness_manifest.v1",
        "status": terminal_status,
        "plan": str(plan.get("_path", "")),
        "packet_count": len(packets),
        "packets": packet_paths,
        "readiness": str(packet_root / "aa-readiness-v1.json"),
        "timed_controls_started": False,
        "real_comparisons_started": False,
        "no_retry": True,
    }
    (packet_root / "aa-readiness-v1.json").write_text(json.dumps(readiness, indent=2) + "\n", encoding="utf-8")
    (packet_root / "aa-readiness-manifest-v1.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")


def capture_sequence(
    plan: dict[str, Any],
    root: Path,
    lane: str,
    *,
    controls_runner: Any | None = None,
    aa_runner: Any | None = None,
    timed_runner: Any | None = None,
    real_runner: Any | None = None,
) -> dict[str, Any]:
    """Run the staged capture order; A/A readiness gates candidate work."""
    controls_runner = run_controls if controls_runner is None else controls_runner
    aa_runner = run_aa if aa_runner is None else aa_runner
    timed_runner = run_timed_controls if timed_runner is None else timed_runner
    real_runner = run_real if real_runner is None else real_runner
    plan = {**plan, "_path": plan.get("_path", "")}
    controls = controls_runner(plan, root)
    if any(packet.get("validation_error") for packet in controls):
        return {"status": "stopped_control_failure", "packets": len(controls)}
    all_packets: list[dict[str, Any]] = list(controls)
    try:
        aa_packets = aa_runner(plan, root)
        readiness = aa_readiness(plan, aa_packets)
    except (OSError, ValueError, KeyError, TypeError, ZeroDivisionError) as exc:
        aa_packets = []
        readiness = {
            "schema": "litchi.performance.0967.aa_readiness.v1",
            "status": "fail",
            "threshold": plan["aa_noise"].get("upper_ci_limit"),
            "bootstrap_seed": plan["aa_noise"].get("bootstrap_seed"),
            "bootstrap_resamples": plan["aa_noise"].get("bootstrap_resamples"),
            "pairs_per_cell": plan["aa_noise"].get("native_pairs_per_cell"),
            "errors": [f"A/A reduction error: {exc}"],
            "cells": {},
        }
    all_packets.extend(aa_packets)
    write_aa_readiness(root, plan, aa_packets, readiness, "aa_readiness_passed" if readiness.get("status") == "pass" else "stopped_aa_readiness")
    if readiness.get("status") != "pass" or any(packet.get("validation_error") for packet in aa_packets):
        return {"status": "stopped_aa_readiness", "packets": len(all_packets), "readiness": readiness}
    timed_packets = timed_runner(plan, root)
    all_packets.extend(timed_packets)
    if any(packet.get("validation_error") for packet in timed_packets):
        return {"status": "stopped_timed_control_failure", "packets": len(all_packets)}
    lanes = [lane] if lane != "all" else ["native", "allocator"]
    for current_lane in lanes:
        packets = real_runner(plan, current_lane, root)
        all_packets.extend(packets)
        if any(packet.get("validation_error") for packet in packets):
            return {"status": "stopped_packet_failure", "lane": current_lane, "packets": len(all_packets)}
    manifest = {
        "schema": SCHEMA,
        "plan": str(plan.get("_path", "")),
        "started_utc": utc(),
        "packet_count": len(all_packets),
        "aa_readiness": str(root / "aa-readiness-v1.json"),
        "packets": [str(root / ("controls" if (p["role"].startswith("control") or p["role"].startswith("timed_control")) else p["context"]["lane"]) / f"{p['packet_id']}.json") for p in all_packets],
    }
    (root / "capture-manifest-v1.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    return {"status": "capture_complete", "packet_count": len(all_packets), "manifest": str(root / "capture-manifest-v1.json")}


def self_test_capture_order() -> int:
    """Synthetic integration test for the AA-before-candidate stage gate."""
    import tempfile

    def plan_for_test() -> dict[str, Any]:
        return {
            "aa_noise": {
                "coverage": "endpoint_cells_only",
                "cells": ["cell/e1/default"],
                "same_binary": "baseline",
                "order": ["AB", "BA"],
                "native_pairs_per_cell": 2,
                "warmups": 5,
                "samples": 80,
                "symmetric_ratio": "max(r,1/r)",
                "upper_ci_limit": 1.05,
                "bootstrap_seed": 967004,
                "bootstrap_resamples": 100,
            }
        }

    def aa_packets(noisy: bool) -> list[dict[str, Any]]:
        packets = []
        for pair in (1, 2):
            order = "AB" if pair == 1 else "BA"
            for position in ("first", "second"):
                multiplier = 2.0 if noisy and order == "AB" and position == "second" else 1.0
                report = {"elapsed_ns": [100.0 if multiplier == 1.0 else 200.0] * 80}
                packets.append({
                    "packet_id": f"aa-{order}-{pair}-{position}",
                    "role": "aa_baseline",
                    "context": {"lane": "aa", "cell": {"cell_id": "cell/e1/default"}, "aa_order": order, "aa_position": position, "pair": pair},
                    "report": report,
                    "validation_error": None,
                })
        return packets

    def run_case(noisy: bool) -> tuple[dict[str, Any], dict[str, int]]:
        called = {"timed": 0, "real": 0}
        plan = plan_for_test()

        def controls(_plan: dict[str, Any], _root: Path) -> list[dict[str, Any]]:
            return []

        def aa(_plan: dict[str, Any], _root: Path) -> list[dict[str, Any]]:
            return aa_packets(noisy)

        def timed(_plan: dict[str, Any], _root: Path) -> list[dict[str, Any]]:
            called["timed"] += 1
            return []

        def real(_plan: dict[str, Any], _lane: str, _root: Path) -> list[dict[str, Any]]:
            called["real"] += 1
            return []

        with tempfile.TemporaryDirectory(prefix="litchi-0967-aa-self-test-") as directory:
            result = capture_sequence(plan, Path(directory), "all", controls_runner=controls, aa_runner=aa, timed_runner=timed, real_runner=real)
            readiness = load(Path(directory) / "aa-readiness-v1.json")
            result["readiness_status"] = readiness.get("status")
        return result, called

    passed, pass_calls = run_case(False)
    if passed.get("status") != "capture_complete" or passed.get("readiness_status") != "pass" or pass_calls != {"timed": 1, "real": 2}:
        raise SystemExit(f"AA pass/go integration test failed: {passed} {pass_calls}")
    failed, fail_calls = run_case(True)
    if failed.get("status") != "stopped_aa_readiness" or failed.get("readiness_status") != "fail" or fail_calls != {"timed": 0, "real": 0}:
        raise SystemExit(f"AA fail/stop integration test failed: {failed} {fail_calls}")
    print(json.dumps({"schema": SCHEMA, "status": "self_test_pass", "aa_pass_go": True, "aa_fail_stops_before_candidate": True}))
    return 0


def control_validator(control: dict[str, Any], report: dict[str, Any], variant: str) -> None:
    if report.get("schema") != CONTROL_SCHEMA or report.get("variant") != variant:
        fail(f"{control['id']}: control schema/variant mismatch")
    if report.get("timing_claim") is not False:
        fail(f"{control['id']}: correctness control timing claim must be false")
    if report.get("error") is not None:
        fail(f"{control['id']}: report error {report['error']}")
    mode = report.get("mode")
    pre = report.get("pre") or {}
    post = report.get("post") or {}
    revisions_before = pre.get("revisions") or []
    revisions_after = post.get("revisions") or []
    if control["id"] == "real-bookmark-delete-redline-query":
        query = report.get("query") or {}
        if mode != "query" or len(revisions_before) != 20 or query.get("replacement_succeeded") is not False or query.get("source_unchanged_after_refusal") is not True:
            fail("query control refusal/reopen oracle failed")
        return
    if mode != "run":
        fail(f"{control['id']}: expected run mode")
    edit = report.get("edit") or {}
    if edit.get("replacement_succeeded") is not True or edit.get("reopen_succeeded") is not True or edit.get("changed") is not True:
        fail(f"{control['id']}: lifecycle edit/reopen failed")
    witness = report.get("branch_witness") or {}
    if control.get("require_revision_survival"):
        if not revisions_before or revisions_after != revisions_before or witness.get("positive_branch_required") is not True:
            fail(f"{control['id']}: revision-bearing control did not survive")
    if control.get("require_false_marker"):
        if not revisions_before or revisions_after != revisions_before or witness.get("false_marker_present") is not True:
            fail(f"{control['id']}: true revision/false-marker witness did not survive")


def control_profile_validator(cell: dict[str, Any], report: dict[str, Any], variant: str, allocator: bool, input_row: dict[str, Any]) -> None:
    if report.get("schema") != "litchi.performance.0967.control_profile.v1":
        fail(f"{cell['id']}: timed control schema mismatch")
    if report.get("variant") != variant or report.get("control") != cell["control"] or report.get("scope") != cell["scope"]:
        fail(f"{cell['id']}: timed control identity mismatch")
    if cell["scope"] == "edit":
        if report.get("edits") != cell.get("edits") or report.get("retention") != "default_8mib":
            fail(f"{cell['id']}: timed control edit parameters mismatch")
    elif report.get("edits") is not None or report.get("retention") is not None:
        fail(f"{cell['id']}: timed control query carries edit parameters")
    if report.get("source_sha256") != input_row["sha256"] or report.get("source_bytes") != input_row["bytes"]:
        fail(f"{cell['id']}: timed control input identity mismatch")
    if bool(report.get("allocator_instrumented")) != allocator or bool(report.get("timing_claim")) == allocator or report.get("allocator_timing_claim") is not False:
        fail(f"{cell['id']}: timed control lane flags mismatch")
    warmups = 0 if allocator else 5
    samples = 1 if allocator else 80
    if report.get("warmups") != warmups or report.get("samples_requested") != samples:
        fail(f"{cell['id']}: timed control warmup/sample mismatch")
    elapsed = report.get("elapsed_ns")
    if not isinstance(elapsed, list) or len(elapsed) != (0 if allocator else samples):
        fail(f"{cell['id']}: timed control elapsed sample count mismatch")
    rows = report.get("samples")
    if not isinstance(rows, list) or len(rows) != samples:
        fail(f"{cell['id']}: timed control sample rows incomplete")
    expected_revisions = {"generated-positive": 2, "dense1024": 1024, "false-marker": 1, "real-positive": 20}[cell["control"]]
    for row in rows:
        if row.get("exact_output_equal_reference") is not True or row.get("public_reopen_succeeded") is not True:
            fail(f"{cell['id']}: timed control semantic/output oracle failed")
        if row.get("revision_count") != expected_revisions:
            fail(f"{cell['id']}: timed control revision count changed")
        if cell["scope"] == "edit":
            if row.get("semantic_change_count") != cell["edits"] or row.get("paragraph_zero_sha256") is None:
                fail(f"{cell['id']}: timed control edit oracle failed")
        elif row.get("source_unchanged") is not True:
            fail(f"{cell['id']}: timed control query mutated source")
        if allocator:
            allocation = row.get("allocation")
            if not isinstance(allocation, dict):
                fail(f"{cell['id']}: allocator control metrics missing")
            for field in ("allocated_bytes", "deallocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes"):
                if not isinstance(allocation.get(field), int):
                    fail(f"{cell['id']}: allocator control field missing {field}")
        elif row.get("elapsed_ns") is None:
            fail(f"{cell['id']}: native control elapsed sample missing")


def run_controls(plan: dict[str, Any], packet_root: Path) -> list[dict[str, Any]]:
    packets = []
    controls = plan["controls"]["matrix"]
    retained = plan["controls"].get("functional_reports") or {}
    if retained:
        # Functional controls are already sealed by the independent control
        # receipt. Reuse their exact JSON packets; this avoids switching the
        # probe's source-active symlink during a capture and launches nothing.
        for control in controls:
            control_id = control["id"]
            if control_id not in retained:
                continue
            for variant in ("baseline", "candidate"):
                row = retained[control_id].get(variant) or {}
                report_path = Path(row.get("path", ""))
                check_hash(report_path, row.get("sha256", ""), f"retained control {control_id}/{variant}")
                report = load(report_path)
                control_validator(control, report, variant)
                role = f"control_{variant}"
                binary = plan["binaries"][role]
                packet_id = f"control-{control_id}-{variant}"
                packet = {
                    "schema": SCHEMA,
                    "packet_id": packet_id,
                    "created_utc": utc(),
                    "role": role,
                    "command": [],
                    "binary": {"path": binary["path"], "sha256": binary["sha256"]},
                    "context": {"packet_id": packet_id, "control_id": control_id, "variant": variant, "source_revision": plan["source"]["revision"], "timed_scope": control.get("timed_scope", False), "retained_receipt": str(report_path)},
                    "terminal": {"returncode": 0, "timed_out": False, "launch_error": None},
                    "wall_ns_excluded_from_metric": None,
                    "stdout": {"path": str(report_path), "bytes": report_path.stat().st_size, "sha256": sha(report_path)},
                    "stderr": {"path": None, "bytes": 0, "sha256": None},
                    "rss": {"path": None, "bytes": 0, "kb": None},
                    "report": report,
                    "validation_error": None,
                    "retained": True,
                }
                write_packet(packet_root / "controls", packet)
                packets.append(packet)
    else:
        for control in controls:
            for variant in ("baseline", "candidate"):
                role = f"control_{variant}"
                binary = Path(plan["binaries"][role]["path"])
                args = ["--variant", variant]
                source_output: Path | None = None
                if control["probe_mode"] == "query-input":
                    args += ["--mode", control["probe_mode"], "--input", plan["fixtures"][control["id"]]["path"]]
                else:
                    args += ["--mode", control["probe_mode"]]
                    if "runs" in control:
                        args += ["--runs", str(control["runs"])]
                    source_output = packet_root / "controls" / f"{control['id']}-{variant}.input.doc"
                    args += ["--output", str(source_output)]
                packet_id = f"control-{control['id']}-{variant}"
                context = {"packet_id": packet_id, "control_id": control["id"], "variant": variant, "source_revision": plan["source"]["revision"], "timed_scope": control.get("timed_scope", False)}
                packet = launch_packet(executable=binary, args=args, role=role, packet_dir=packet_root / "controls", context=context, validator=lambda report, control=control, variant=variant: control_validator(control, report, variant))
                if source_output is not None and packet.get("validation_error") is None:
                    packet["generated_source"] = {"path": str(source_output), "bytes": source_output.stat().st_size if source_output.exists() else 0, "sha256": sha(source_output) if source_output.exists() else None}
                write_packet(packet_root / "controls", packet)
                packets.append(packet)
    return packets


def run_timed_controls(plan: dict[str, Any], packet_root: Path) -> list[dict[str, Any]]:
    timed = plan["controls"]["timed_scope"]
    inputs = timed["inputs"]
    packets: list[dict[str, Any]] = []
    for lane, pairs, warmups, samples, role_prefix in (
        ("control_native", timed["native_pairs_per_control"], timed["warmups"], timed["samples"], "timed_control_native"),
        ("control_allocator", timed["allocator_pairs_per_control"], timed["allocator_warmups"], timed["allocator_samples"], "timed_control_allocator"),
    ):
        schedule = [
            {"pair": pair, "relative_order": "baseline_then_candidate" if pair <= pairs // 2 else "candidate_then_baseline"}
            for pair in range(1, pairs + 1)
        ]
        for cell in timed["cells"]:
            input_row = inputs[cell["control"]]
            if input_row.get("available") is not True:
                raise ValueError(f"timed control input unavailable: {cell['control']}")
            for pair_row in schedule:
                order = ["baseline", "candidate"] if pair_row["relative_order"] == "baseline_then_candidate" else ["candidate", "baseline"]
                for variant in order:
                    role = f"{role_prefix}_{variant}"
                    binary = Path(plan["binaries"][role]["path"])
                    args = ["--variant", variant, "--control", cell["control"], "--input", input_row["path"], "--warmup", str(warmups), "--samples", str(samples)]
                    if cell["scope"] == "edit":
                        args += ["--edits", str(cell["edits"]), "--retention", cell["retention"]]
                    packet_id = f"{lane}-{cell['id']}-p{pair_row['pair']:02d}-{variant}"
                    context = {"packet_id": packet_id, "lane": lane, "variant": variant, "cell": cell, "control_id": cell["control"], "pair": pair_row["pair"], "relative_order": pair_row["relative_order"], "source_revision": plan["source"]["revision"], "fixture_sha256": input_row["sha256"], "warmups": warmups, "samples": samples}
                    validator = lambda report, cell=cell, variant=variant, allocator=lane == "control_allocator", input_row=input_row: control_profile_validator(cell, report, variant, allocator, input_row)
                    packet = launch_packet(executable=binary, args=args, role=role, packet_dir=packet_root / "controls", context=context, validator=validator)
                    write_packet(packet_root / "controls", packet)
                    packets.append(packet)
    return packets


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", type=Path, default=Path("/home/zhuhe/litchi-performance-0967/plans/plan-0967-v1.json"))
    ap.add_argument("--check-plan", action="store_true")
    ap.add_argument("--allow-unbound", action="store_true")
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--lane", choices=["native", "allocator", "all"], default="all")
    ap.add_argument("--output-root", type=Path)
    args = ap.parse_args()
    if args.self_test:
        return self_test_capture_order()
    if not args.check_plan and not args.run:
        args.check_plan = True
    try:
        if args.run and args.lane != "all":
            raise ValueError("capture requires --lane all; partial lanes cannot qualify and cannot be retried")
        plan = check_plan(args.plan.resolve(), allow_unbound=args.allow_unbound or not args.run, for_capture=args.run)
        if args.check_plan and not args.run:
            print(json.dumps({"schema": SCHEMA, "status": "preflight_pass", "plan": str(args.plan.resolve()), "plan_status": plan.get("status"), "binaries_bound": all(bool(row.get("sha256")) for row in plan["binaries"].values()), "controls_bound": plan["controls"].get("status") == "bound", "timed_controls_bound": plan["controls"].get("timed_scope", {}).get("status") == "bound", "timed_inputs_available": all(row.get("available") is True for row in plan["controls"].get("timed_scope", {}).get("inputs", {}).values()), "host_bound": plan.get("host_state", {}).get("status") == "bound"}))
            return 0
        root = (args.output_root or Path(plan["outputs"]["capture_root"])).resolve()
        if root.exists() and any(root.iterdir()):
            raise ValueError(f"capture output root is not empty; no retry/replacement is permitted: {root}")
        root.mkdir(parents=True, exist_ok=True)
        plan["_path"] = str(args.plan.resolve())
        result = capture_sequence(plan, root, args.lane)
        print(json.dumps(result))
        return 0 if result.get("status") == "capture_complete" else 2
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"runner rejected plan: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
