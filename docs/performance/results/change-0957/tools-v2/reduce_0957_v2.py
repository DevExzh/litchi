#!/usr/bin/env python3
"""Offline descriptive reducer for a completed 0957 allocator packet.

It validates the sealed packet before reading reports and emits only allocator
counter summaries. No elapsed-time, bootstrap, adoption, or speed claim is
computed.
"""
from __future__ import annotations

import argparse
import json
import math
import os
import sys
from pathlib import Path
from typing import Any

from common_0957_v2 import (
    ContractError,
    descriptor_matches,
    expected_children,
    exclusive_json,
    file_descriptor,
    finite_metric,
    load_json,
    mean,
    packet_path,
    require_dict,
    sha256_file,
    validate_plan_schedule,
    validate_report,
)
from capture_0957_v2 import child_argv, child_environment, external_descriptor

SUMMARY_METRICS = (
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
DERIVED_FIELDS = (
    "max_setup_core_peak_above_setup_live_before",
    "max_setup_drained_core_peak_above_setup_live_before",
)
SUMMARY_PHASES = (
    "setup",
    "ingress",
    "edit_commit",
    "publication",
    "commit_drop",
    "worker_drain",
    "core_combined",
    "core_combined_with_drain",
    "postdrop",
)


def fail(message: str) -> "NoReturn":
    raise ContractError(message)


def normalize_packet_path(value: Any, old_root: str, new_root: str) -> Any:
    if isinstance(value, str):
        if value == old_root:
            return new_root
        prefix = old_root.rstrip(os.sep) + os.sep
        if value.startswith(prefix):
            return new_root.rstrip(os.sep) + value[len(old_root.rstrip(os.sep)):]
        return value
    if isinstance(value, list):
        return [normalize_packet_path(item, old_root, new_root) for item in value]
    if isinstance(value, dict):
        return {key: normalize_packet_path(item, old_root, new_root) for key, item in value.items()}
    return value


def verify_packet_manifest(packet: Path) -> dict[str, Any]:
    manifest_path = packet / "packet-manifest.json"
    manifest = load_json(manifest_path, max_bytes=32 * 1024 * 1024)
    if not isinstance(manifest, dict) or manifest.get("schema") != "litchi.performance.0957.completed-packet-manifest.v1":
        fail("completed packet manifest schema is invalid")
    entries = manifest.get("entries")
    if not isinstance(entries, list) or len(entries) == 0:
        fail("completed packet manifest has no entries")
    listed: set[str] = set()
    for entry in entries:
        if not isinstance(entry, dict) or not isinstance(entry.get("path"), str):
            fail("packet manifest entry is malformed")
        rel = entry["path"]
        if rel in listed:
            fail(f"duplicate packet manifest path: {rel}")
        listed.add(rel)
        path = packet_path(packet, rel)
        descriptor_matches(path, entry)
    actual: set[str] = set()
    for path in packet.rglob("*"):
        if path == manifest_path:
            continue
        if path.is_symlink() or (path.exists() and not path.is_file() and not path.is_dir()):
            fail(f"packet contains unsupported artifact: {path}")
        if path.is_file():
            actual.add(str(path.relative_to(packet)))
    if actual != listed:
        missing = sorted(listed - actual)
        extra = sorted(actual - listed)
        fail(f"packet manifest census mismatch missing={missing[:3]} extra={extra[:3]}")
    return manifest


def _same_descriptor(actual: Any, expected: dict[str, Any], label: str) -> None:
    actual = actual if isinstance(actual, dict) else {}
    for key in ("path", "bytes", "sha256"):
        if actual.get(key) != expected.get(key):
            fail(f"{label}.{key} does not match expected custody")


def validate_child(packet: Path, plan: dict[str, Any], child: dict[str, Any], old_root: str) -> dict[str, Any]:
    child_root = packet / "children" / child["child_id"]
    report_dir = packet / "reports" / child["child_id"]
    start = load_json(child_root / "start.json")
    terminal = load_json(child_root / "terminal.json")
    validation = load_json(child_root / "validation.json")
    if start.get("child") != child or terminal.get("child") != child or validation.get("child") != child:
        fail(f"{child['child_id']} child identity mismatch")
    if start.get("index") != child["index"] or terminal.get("index") != child["index"]:
        fail(f"{child['child_id']} index mismatch")
    # The captured v1 receipt binds the complete child object above.  Its
    # redundant top-level index is optional; reject a contradictory one.
    if "index" in validation and validation["index"] != child["index"]:
        fail(f"{child['child_id']} validation index mismatch")
    if start.get("process_group") is not True:
        fail(f"{child['child_id']} start did not request a process group")

    historical_report_dir = Path(old_root) / "reports" / child["child_id"]
    expected_argv = child_argv(child, plan, historical_report_dir)
    expected_env = child_environment(plan)
    expected_cwd = str(Path(plan["execution"]["cwd"]))
    expected_report_path = str(historical_report_dir / "report.json")
    expected_output_dir = str(historical_report_dir)
    expected_plan_descriptor = {
        "path": "custody/plan.json",
        "bytes": (packet / "custody" / "plan.json").stat().st_size,
        "sha256": sha256_file(packet / "custody" / "plan.json"),
    }
    for label, record in (("start", start), ("terminal", terminal)):
        if record.get("argv") != expected_argv:
            fail(f"{child['child_id']} {label} argv does not match the frozen plan")
        if record.get("cwd") != expected_cwd:
            fail(f"{child['child_id']} {label} cwd does not match the frozen plan")
        if record.get("environment") != expected_env:
            fail(f"{child['child_id']} {label} environment does not match the frozen plan")
        if record.get("deadline_seconds") != plan["schedule"]["deadline_seconds"]:
            fail(f"{child['child_id']} {label} deadline mismatch")
        _same_descriptor(record.get("plan"), expected_plan_descriptor, f"{child['child_id']} {label}.plan")
    if start.get("output_directory") != expected_output_dir:
        fail(f"{child['child_id']} start output directory does not match the frozen plan")
    if terminal.get("report_path") != expected_report_path or terminal.get("report_directory") != expected_output_dir:
        fail(f"{child['child_id']} terminal report path does not match the frozen plan")
    if terminal.get("returncode") != 0 or terminal.get("timed_out") is not False:
        fail(f"{child['child_id']} terminal status is not successful")
    if terminal.get("leader_reaped") is not True or terminal.get("process_group_reaped") is not True:
        fail(f"{child['child_id']} process group was not successfully reaped")
    if terminal.get("forced_group_cleanup") is not False:
        fail(f"{child['child_id']} required forced process-group cleanup")
    if validation.get("status") != "passed":
        fail(f"{child['child_id']} validation receipt is not passed")
    _same_descriptor(validation.get("plan"), expected_plan_descriptor, f"{child['child_id']} validation.plan")
    terminal_descriptor = file_descriptor(child_root / "terminal.json", relative_to=packet)
    _same_descriptor(validation.get("terminal"), terminal_descriptor, f"{child['child_id']} validation.terminal")
    if validation.get("terminal_sha256") != terminal_descriptor["sha256"]:
        fail(f"{child['child_id']} validation terminal SHA does not match sealed terminal")
    for stream in ("stdout", "stderr"):
        path = child_root / f"{stream}.log"
        _same_descriptor(
            terminal.get(stream),
            file_descriptor(path, relative_to=packet),
            f"{child['child_id']} terminal.{stream}",
        )
    report_path = report_dir / "report.json"
    report_descriptor = file_descriptor(report_path, relative_to=packet)
    _same_descriptor(terminal.get("report"), report_descriptor, f"{child['child_id']} terminal.report")
    _same_descriptor(validation.get("report"), report_descriptor, f"{child['child_id']} validation.report")
    report = load_json(report_path)
    validate_report(report, plan, child)
    return report

def stats(values: list[int]) -> dict[str, Any]:
    if not values:
        fail("cannot summarize empty metric vector")
    return {"min": min(values), "max": max(values), "mean": mean(values), "count": len(values)}


def metric_summary(report: dict[str, Any], phase: str, metric: str) -> dict[str, Any]:
    rows = report["case"]["rows"]
    values = [finite_metric(row[phase], metric) for row in rows]
    return stats(values)


def derived_summary(report: dict[str, Any], field: str) -> dict[str, Any]:
    values = [finite_metric(row, field) for row in report["case"]["rows"]]
    return stats(values)


def ratio(numerator: float, denominator: float) -> dict[str, Any]:
    if denominator == 0:
        if numerator == 0:
            return {"status": "undefined_zero_zero", "ratio": None, "regression_over_5_percent": None}
        return {"status": "positive_over_zero", "ratio": None, "regression_over_5_percent": True}
    value = numerator / denominator
    return {"status": "finite", "ratio": value, "regression_over_5_percent": value > 1.05}


def reduce_packet(packet: Path, plan_path: Path, output: Path) -> dict[str, Any]:
    if output.resolve().is_relative_to(packet.resolve()):
        fail("reduction output must be outside the sealed packet")
    plan = load_json(plan_path, max_bytes=4 * 1024 * 1024)
    if not isinstance(plan, dict):
        fail("plan is not an object")
    if plan.get("status") != "authorized":
        fail("reduction requires the authorized formal plan")
    children = validate_plan_schedule(plan)
    for label in ("source", "build", "binary", "runner"):
        external_descriptor(label, plan.get(label))
    manifest = verify_packet_manifest(packet)
    capture_start = load_json(packet / "custody" / "capture-start.json")
    plan_descriptor = file_descriptor(plan_path)
    copied_plan_descriptor = file_descriptor(packet / "custody" / "plan.json", relative_to=packet)
    if any(copied_plan_descriptor[key] != plan_descriptor[key] for key in ("bytes", "sha256")):
        fail("copied packet plan differs from the external authorized plan")
    runner_binding = load_json(packet / "custody" / "runner.json")
    _same_descriptor(runner_binding, plan.get("runner", {}), "runner.json")
    _same_descriptor(capture_start.get("plan"), plan_descriptor, "capture-start.plan")
    _same_descriptor(capture_start.get("copied_plan"), copied_plan_descriptor, "capture-start.copied_plan")
    runner_descriptor = require_dict(plan.get("runner"), "plan.runner")
    _same_descriptor(capture_start.get("runner"), runner_descriptor, "capture-start.runner")
    if capture_start.get("children") != children:
        fail("capture-start child schedule differs from the authorized plan")
    state = load_json(packet / "capture-state.json")
    if state.get("status") != "completed" or state.get("children_completed") != len(children):
        fail("capture state is not a completed 27-child packet")
    outer = load_json(packet / "custody" / "outer-terminal.json")
    if outer.get("status") != "completed" or outer.get("children_expected") != len(children) or outer.get("children_completed") != len(children):
        fail("outer capture terminal is not completed")
    _same_descriptor(outer.get("plan"), plan_descriptor, "outer-terminal.plan")
    _same_descriptor(outer.get("copied_plan"), copied_plan_descriptor, "outer-terminal.copied_plan")
    _same_descriptor(outer.get("runner"), runner_descriptor, "outer-terminal.runner")
    old_root = str(manifest.get("packet_root", packet.resolve()))
    reports: dict[str, dict[str, Any]] = {}
    for child in children:
        reports[child["child_id"]] = validate_child(packet, plan, child, old_root)
    by_case_width: dict[tuple[str, int], list[tuple[int, dict[str, Any]]]] = {}
    for child in children:
        by_case_width.setdefault((child["case_id"], child["workers"]), []).append((child["repeat"], reports[child["child_id"]]))
    configs: list[dict[str, Any]] = []
    for case in plan["cases"]:
        case_id = case["case_id"]
        for width in (1, 2, 4):
            repeats = sorted(by_case_width[(case_id, width)], key=lambda pair: pair[0])
            config: dict[str, Any] = {"case_id": case_id, "workers": width, "repeat_count": len(repeats), "repeats": []}
            for repeat, report in repeats:
                phases: dict[str, Any] = {}
                for phase in SUMMARY_PHASES:
                    phases[phase] = {metric: metric_summary(report, phase, metric) for metric in SUMMARY_METRICS}
                derived = {field: derived_summary(report, field) for field in DERIVED_FIELDS}
                config["repeats"].append({"repeat": repeat, "phases": phases, "derived_peak_fields": derived})
            config["primary_core_combined_with_drain"] = {
                "phase": "core_combined_with_drain",
                "metrics": ["allocation_calls", "allocated_bytes"],
                "derived_peak_fields": list(DERIVED_FIELDS),
            }
            configs.append(config)
    comparisons: list[dict[str, Any]] = []
    for case in plan["cases"]:
        case_id = case["case_id"]
        reference = {item["repeat"]: item for item in next(c for c in configs if c["case_id"] == case_id and c["workers"] == 1)["repeats"]}
        for width in (2, 4):
            candidate = {item["repeat"]: item for item in next(c for c in configs if c["case_id"] == case_id and c["workers"] == width)["repeats"]}
            for repeat in sorted(reference):
                per_phase: dict[str, Any] = {}
                for phase in SUMMARY_PHASES:
                    per_phase[phase] = {}
                    for metric in SUMMARY_METRICS:
                        base = reference[repeat]["phases"][phase][metric]["mean"]
                        current = candidate[repeat]["phases"][phase][metric]["mean"]
                        per_phase[phase][metric] = ratio(current, base)
                derived_ratios = {}
                for field in DERIVED_FIELDS:
                    base = reference[repeat]["derived_peak_fields"][field]["mean"]
                    current = candidate[repeat]["derived_peak_fields"][field]["mean"]
                    derived_ratios[field] = ratio(current, base)
                comparisons.append({
                    "case_id": case_id,
                    "workers": width,
                    "reference_workers": 1,
                    "repeat": repeat,
                    "ratios": per_phase,
                    "derived_peak_ratios": derived_ratios,
                })
    return {
        "schema": "litchi.performance.0957.allocator-reduction.v2",
        "status": "completed",
        "claim_scope": "descriptive allocator counters and observer peaks only; no elapsed-time, bootstrap, RSS, throughput, speedup, or adoption claim",
        "plan": file_descriptor(plan_path),
        "packet_manifest": file_descriptor(packet / "packet-manifest.json"),
        "children": len(children),
        "configs": configs,
        "width_ratios_to_worker1": comparisons,
        "regression_threshold": 1.05,
        "source_binary_report_validation": "all 27 reports passed the frozen contract and counter conservation checks",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = reduce_packet(args.packet.resolve(), args.plan.resolve(), args.output.resolve())
        exclusive_json(args.output.resolve(), result)
        return 0
    except Exception as exc:
        print(f"reduce_0957_v2: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
