#!/usr/bin/env python3
"""Bounded offline audit of the failed 0975 instruction-profile capture.

This tool only reads retained JSON/decoder outputs.  It does not run perf,
open /proc, decode new raw data, launch the harness, or reduce the cohort.
The child-02 script streams are parsed descriptively even though raw loss
makes that child inadmissible.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import re
import sys
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

ROOT = Path("/home/zhuhe/litchi-performance-0975")
CAPTURE = ROOT / "capture/instruction-profile-v1"
PLAN = ROOT / "plans/plan-0975-pre-capture-v3.json"
DESIGN = ROOT / "design-0975-v1.md"
PARSER_PATH = Path("/home/zhuhe/litchi-performance-0973/tools-v2/protocol_common_0973_v2.py")
ADAPTER_V5 = Path("/home/zhuhe/litchi-performance-0974/tools/protocol_adapter_0974_v5.py")
RUNNER_V5 = Path("/home/zhuhe/litchi-performance-0974/tools/run_0974_v5.py")
ANALYSIS = ROOT / "analysis"
RECEIPT = ROOT / "receipts/failed-capture-analysis-0975-v3.json"

EXPECTED_PERIOD = 1_000_000
EXPECTED_CPU = 31
MODES = ("stock", "noop", "corrected")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    path = path.absolute()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError(f"expected object: {path}")
    return value


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def tail_lines(path: Path, count: int = 4) -> list[str]:
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    return lines[-count:]


def top_period(mapping: Any, limit: int = 10) -> list[dict[str, Any]]:
    if not isinstance(mapping, dict):
        return []
    rows = sorted(
        ((str(key), int(value)) for key, value in mapping.items()),
        key=lambda row: (-row[1], row[0]),
    )
    return [{"owner": key, "period": period} for key, period in rows[:limit]]


def metric_row(parsed: dict[str, Any]) -> dict[str, Any]:
    fields = (
        "sample_count", "total_period", "lost_events", "empty_callchain_count",
        "explicit_stack_truncated_count", "stack_truncated", "unresolved_frames",
        "marker_sample_count", "marker_period", "marker_period_share",
        "marker_ancestry", "marker_unknown_physical_leaf_samples",
        "marker_no_resolved_frame_samples", "physical_leaf_unknown_samples",
        "physical_leaf_unknown_period", "physical_leaf_classified_period",
    )
    row = {field: parsed.get(field) for field in fields}
    physical = parsed.get("physical_leaf_owner_period") or {}
    proxy = parsed.get("deepest_resolved_proxy_period") or {}
    inclusive = parsed.get("inclusive_marker_ancestor_period") or {}
    row["physical_leaf_owner_count"] = len(physical)
    row["physical_leaf_top10"] = top_period(physical)
    row["deepest_resolved_proxy_owner_count"] = len(proxy)
    row["deepest_resolved_proxy_top10"] = top_period(proxy)
    row["inclusive_marker_ancestor_owner_count"] = len(inclusive)
    # This is deliberately retained as an overlapping coverage sum.  It is
    # never interpreted as a partition or added to another owner table.
    row["inclusive_marker_ancestor_coverage_period_sum"] = sum(int(v) for v in inclusive.values())
    row["inclusive_marker_ancestor_top10"] = top_period(inclusive)
    row["inclusive_marker_ancestor_note"] = "per-owner marker coverage; owners overlap and must not be summed as attribution"
    return row


def report_summary(report: dict[str, Any], report_path: Path) -> dict[str, Any]:
    elapsed = report.get("elapsed_ns")
    if not isinstance(elapsed, list) or not elapsed:
        raise RuntimeError(f"report elapsed_ns missing: {report_path}")
    correctness = report.get("correctness") or {}
    false_keys = sorted(key for key, value in correctness.items() if value is not True)
    route_rows = []
    for route in report.get("route_outputs") or []:
        route_rows.append({
            key: route.get(key)
            for key in (
                "retention", "retention_ceiling_bytes", "output_bytes", "output_capacity",
                "semantic_change_count", "final_retained_before_commit",
            )
        })
    return {
        "binding": binding(report_path),
        "schema": report.get("schema"),
        "case": report.get("case"),
        "edits": report.get("edits"),
        "retention": report.get("retention"),
        "retention_ceiling_bytes": report.get("retention_ceiling_bytes"),
        "warmups": report.get("warmups"),
        "samples_requested": report.get("samples_requested"),
        "timing_claim": report.get("timing_claim"),
        "allocator_instrumented": report.get("allocator_instrumented"),
        "allocator_timing_claim": report.get("allocator_timing_claim"),
        "source": {
            "path": report.get("input"),
            "bytes": report.get("source_bytes"),
            "sha256": report.get("source_sha256"),
        },
        "output": {
            "bytes": report.get("output_bytes"),
            "capacity": report.get("output_capacity"),
            "sha256": report.get("output_sha256"),
            "exact_output_equal_reference": report.get("exact_output_equal_reference"),
        },
        "elapsed_ns": {
            "count": len(elapsed),
            "min": min(elapsed),
            "max": max(elapsed),
            "sum": sum(elapsed),
        },
        "routes": route_rows,
        "correctness": {
            "gate_count": len(correctness),
            "true_count": sum(value is True for value in correctness.values()),
            "false_keys": false_keys,
        },
    }


def command_summary(command: dict[str, Any], launch: dict[str, Any], command_path: Path, launch_path: Path) -> dict[str, Any]:
    argv = command.get("argv")
    return {
        "command_binding": binding(command_path),
        "launch_binding": binding(launch_path),
        "argv": argv,
        "environment": command.get("environment"),
        "working_directory": command.get("cwd", command.get("working_directory")),
        "cwd_policy": command.get("cwd_policy"),
        "launch": {
            "execve": launch.get("execve"),
            "exec_path": launch.get("exec_path"),
            "pid": launch.get("pid"),
            "exec_pid": launch.get("exec_pid"),
            "requested_affinity": launch.get("requested_affinity"),
            "actual_affinity": launch.get("actual_affinity"),
        },
    }


def stack_summary(raw: dict[str, Any], raw_path: Path) -> dict[str, Any]:
    task = raw.get("task_stream") or {}
    scope = raw.get("scope") or {}
    return {
        "binding": binding(raw_path),
        "schema": raw.get("schema"),
        "status": raw.get("status"),
        "period": raw.get("period"),
        "stack_cap": raw.get("stack_cap"),
        "sample_type": raw.get("sample_type"),
        "sample_count": raw.get("sample_count"),
        "sample_line_count": raw.get("sample_line_count"),
        "raw_event_count": raw.get("raw_event_count"),
        "record_type_counts": raw.get("record_type_counts"),
        "raw_lost_record_count": raw.get("raw_lost_record_count"),
        "raw_lost_sum": raw.get("raw_lost_sum"),
        "raw_lost_scope": raw.get("raw_lost_scope"),
        "unconsumed_raw_count": raw.get("unconsumed_raw_count"),
        "bounds_valid": raw.get("bounds_valid"),
        "pid_tid_cpu_valid": raw.get("pid_tid_cpu_valid"),
        "period_valid": raw.get("period_valid"),
        "violations": raw.get("violations"),
        "pid_tid_cpu_counts": raw.get("pid_tid_cpu_counts"),
        "scope": scope,
        "task_stream": {
            "status": task.get("status"),
            "expected_pid": task.get("expected_pid"),
            "event_counts": task.get("event_counts"),
            "thread_maps": task.get("thread_maps"),
            "cpu_maps": task.get("cpu_maps"),
            "comm_records": task.get("comm_records"),
            "exit_records": task.get("exit_records"),
            "errors": task.get("errors"),
            "missing_required_comm_names": task.get("missing_required_comm_names"),
            "scope": task.get("scope"),
            "authority": task.get("authority"),
        },
        "task_scope_interpretation": (
            {
                "kind": "complete_lifetime_task_stream_for_admission",
                "complete_lifetime_proof": True,
                "reason": "zero raw PERF_RECORD_LOST sum permits the retained raw task stream to serve as lifetime authority; /proc snapshots remain bounded observations",
            }
            if raw.get("raw_lost_sum") == 0 else {
                "kind": "parser_of_retained_records_only",
                "complete_lifetime_proof": False,
                "reason": "nonzero raw PERF_RECORD_LOST means task/lifecycle records may be missing; parser pass is only an observation of retained records",
            }
        ),
    }


def terminal_summary(terminal: dict[str, Any], terminal_path: Path, stderr_path: Path) -> dict[str, Any]:
    return {
        "binding": binding(terminal_path),
        "schema": terminal.get("schema"),
        "returncode": terminal.get("returncode"),
        "timed_out": terminal.get("timed_out"),
        "launch_error": terminal.get("launch_error"),
        "process_pid": terminal.get("process_pid"),
        "started_utc": terminal.get("started_utc"),
        "finished_utc": terminal.get("finished_utc"),
        "native_ld_preload": terminal.get("native_ld_preload"),
        # Preserve the exact bounded terminal trailer object rather than
        # combining terminal LOST_SAMPLES with raw PERF_RECORD_LOST.
        "lost_samples_summary": terminal.get("lost_samples_summary"),
        "perf_record_stderr": {
            "binding": binding(stderr_path),
            "tail_lines": tail_lines(stderr_path),
        },
    }


def parser_coverage_evidence(parser_path: Path) -> dict[str, Any]:
    lines = parser_path.read_text(encoding="utf-8").splitlines()
    empty_definition = [index for index, line in enumerate(lines, 1) if "empty_callchains = sum" in line]
    empty_assignment = [index for index, line in enumerate(lines, 1) if "parsed[\"empty_callchain_count\"]" in line]
    truncation_definition = [index for index, line in enumerate(lines, 1) if "explicit_truncation = sum" in line]
    return {
        "parser_binding": binding(parser_path),
        "empty_callchain_definition_lines": empty_definition,
        "empty_callchain_assignment_lines": empty_assignment,
        "explicit_truncation_definition_lines": truncation_definition,
        "meaning": "the count is made from decoded samples whose parser frame list is empty; it does not identify why frames were absent",
    }


def source_order_evidence() -> dict[str, Any]:
    """Bind the retained write-admission ordering used to explain no packet."""
    source = ADAPTER_V5.read_text(encoding="utf-8")
    lines = source.splitlines()
    require_lines = [index for index, line in enumerate(lines, 1) if "V1.require_zero_loss(raw_summary" in line]
    packet_lines = [index for index, line in enumerate(lines, 1) if "endswith(\".packet.json\")" in line]
    raw_lines = [index for index, line in enumerate(lines, 1) if "raw_summary = load_json" in line]
    return {
        "adapter_binding": binding(ADAPTER_V5),
        "runner_binding": binding(RUNNER_V5),
        "packet_write_hook_line": packet_lines,
        "raw_summary_load_line": raw_lines,
        "require_zero_loss_line": require_lines,
        "interpretation": "retained adapter checks raw/task artifacts and calls require_zero_loss while admitting a complete packet; a nonzero raw sum therefore raises after decoder artifacts exist and before a packet is written",
    }


def parse_child(child: int, parser: Any, manifest: dict[str, Any]) -> dict[str, Any]:
    child_dir = CAPTURE / f"child-{child:02d}"
    stem = f"child-{child:02d}"
    paths = {
        "command": child_dir / f"{stem}.command.json",
        "launch": child_dir / f"{stem}.launch.json",
        "terminal": child_dir / f"{stem}.terminal.json",
        "report": child_dir / f"{stem}.report.json",
        "raw_stack": child_dir / f"{stem}.raw-stack.json",
        "target_layout": child_dir / f"{stem}.target-layout.json",
        "header": child_dir / f"{stem}.perf-header.txt",
        "buildid": child_dir / f"{stem}.perf-buildid.txt",
        "perf_data": child_dir / f"{stem}.perf.data",
        "mmap_dump": child_dir / f"{stem}.perf-mmap-dump.txt",
        "stderr": child_dir / f"{stem}.perf.stderr",
        "inline": child_dir / f"{stem}.perf-script-inline.txt",
        "stock_script": child_dir / f"{stem}.stock.perf-script-no-inline.txt",
        "stock_stderr": child_dir / f"{stem}.stock.perf.stderr",
        "stock_no_log": child_dir / f"{stem}.stock-no-log",
        "noop_script": child_dir / f"{stem}.noop.perf-script-no-inline.txt",
        "noop_stderr": child_dir / f"{stem}.noop.perf.stderr",
        "noop_log": child_dir / f"{stem}.noop.dwfl.jsonl",
        "corrected_script": child_dir / f"{stem}.corrected.perf-script-no-inline.txt",
        "corrected_stderr": child_dir / f"{stem}.corrected.perf.stderr",
        "corrected_log": child_dir / f"{stem}.corrected.dwfl.jsonl",
    }
    packet_path = child_dir / f"{stem}.packet.json"
    if packet_path.is_file():
        paths["packet"] = packet_path

    all_bindings = {name: binding(path) for name, path in paths.items()}
    # Packet artifact naming uses `layout`; the local path table calls the
    # same retained file `target_layout` to avoid confusing it with the
    # terminal/layout JSON roles.
    all_bindings["layout"] = all_bindings["target_layout"]
    raw = load_json(paths["raw_stack"])
    terminal = load_json(paths["terminal"])
    report = load_json(paths["report"])
    command = load_json(paths["command"])
    launch = load_json(paths["launch"])
    layout = load_json(paths["target_layout"])
    packet = load_json(packet_path) if packet_path.is_file() else None
    expected_pid = launch.get("exec_pid")
    if not isinstance(expected_pid, int):
        raise RuntimeError(f"child {child}: launch exec_pid missing")

    parsed: dict[str, dict[str, Any]] = {}
    for mode in MODES:
        text = paths[f"{mode}_script"].read_text(encoding="utf-8", errors="strict")
        parsed[mode] = parser.parse_perf_script_strict(
            text,
            expected_pid=expected_pid,
            expected_cpu=EXPECTED_CPU,
            expected_period=EXPECTED_PERIOD,
        )
    selected_metric_fields = (
        "sample_count", "total_period", "lost_events", "empty_callchain_count",
        "marker_sample_count", "marker_period", "marker_period_share",
        "unresolved_frames", "explicit_stack_truncated_count",
    )
    metric_identity = {
        mode: metric_row(parsed[mode]) for mode in MODES
    }
    stock_noop_equal = all(
        parsed["stock"].get(field) == parsed["noop"].get(field)
        for field in selected_metric_fields
    )
    stock_noop_file_equal = sha256(paths["stock_script"]) == sha256(paths["noop_script"])
    packet_comparison: dict[str, Any] = {"present": packet is not None}
    if packet is not None:
        packet_perf = packet.get("perf") or {}
        mismatches: list[dict[str, Any]] = []
        for mode in MODES:
            packet_mode = packet_perf.get(mode) or {}
            for field in selected_metric_fields:
                if packet_mode.get(field) != parsed[mode].get(field):
                    mismatches.append({"mode": mode, "field": field, "packet": packet_mode.get(field), "reparsed": parsed[mode].get(field)})
        artifact_mismatches: list[str] = []
        for name, row in (packet.get("artifacts") or {}).items():
            actual = all_bindings.get(name)
            if actual is None or any(row.get(key) != actual.get(key) for key in ("path", "bytes", "sha256")):
                artifact_mismatches.append(name)
        packet_comparison.update({
            "schema": packet.get("schema"),
            "status": packet.get("status"),
            "gate": packet.get("gate"),
            "artifact_binding_mismatches": artifact_mismatches,
            "artifact_bindings_match": not artifact_mismatches,
            "reparse_metric_mismatches": mismatches,
            "reparse_metrics_match": not mismatches,
            "task_audit_status": (packet.get("task_audit") or {}).get("status"),
        })

    failure = manifest.get("failure") if child == 2 else None
    raw_loss_sum = raw.get("raw_lost_sum")
    terminal_loss = (terminal.get("lost_samples_summary") or {}).get("lost_samples")
    return {
        "child": child,
        "packet_present": packet is not None,
        "packet_admission": "admitted" if packet is not None else "absent_after_decoder_outputs",
        "failure_record": failure,
        "artifact_bindings": all_bindings,
        "command": command_summary(command, launch, paths["command"], paths["launch"]),
        "terminal": terminal_summary(terminal, paths["terminal"], paths["stderr"]),
        "raw_stack": stack_summary(raw, paths["raw_stack"]),
        "target_layout": {
            "binding": binding(paths["target_layout"]),
            "schema": layout.get("schema"),
            "binary": layout.get("binary"),
            "corrected_base": layout.get("corrected_base"),
            "perf_expected_start": layout.get("perf_expected_start"),
            "process_tree": layout.get("process_tree"),
            "executable_mapping": layout.get("executable_mapping"),
        },
        "report": report_summary(report, paths["report"]),
        "decoder_modes": metric_identity,
        "decoder_controls": {
            "stock_noop_file_byte_identical": stock_noop_file_equal,
            "stock_noop_metric_identity": stock_noop_equal,
            "stock_noop_control_stdout_sha256": [parsed["stock"].get("control_stdout_sha256"), parsed["noop"].get("control_stdout_sha256")],
            "corrected_has_marker_ancestry": parsed["corrected"].get("marker_sample_count", 0) > 0,
            "all_mode_sample_count_equal_raw": all(parsed[mode].get("sample_count") == raw.get("sample_count") for mode in MODES),
            "all_mode_period_sum_equal_raw": all(parsed[mode].get("total_period") == raw.get("sample_count", 0) * EXPECTED_PERIOD for mode in MODES),
            "raw_vs_terminal_loss": {
                "raw_perf_record_lost_sum": raw_loss_sum,
                "terminal_perf_record_lost_samples": terminal_loss,
                "kept_as_separate_gates": True,
            },
        },
        "packet_reconciliation": packet_comparison,
        "empty_chain_coverage": {
            "definition": "empty_callchain_count is the parser count of decoded sample records with no parsed frame lines; explicit stack truncation is counted separately",
            "raw_sample_denominator": raw.get("sample_count"),
            "stock_empty": parsed["stock"].get("empty_callchain_count"),
            "noop_empty": parsed["noop"].get("empty_callchain_count"),
            "corrected_empty": parsed["corrected"].get("empty_callchain_count"),
            "stock_noop_scripts_byte_identical": stock_noop_file_equal,
            "coverage_limit": "These are exported/decoded callchain coverage counts, not a causal diagnosis of perf capture, unwinding, or symbolization loss.",
            "not_established": ["why individual samples have no frame lines", "whether missing frames arose during capture, unwind, export, or resolution"],
        },
    }


def main() -> int:
    parser = load_module(PARSER_PATH, "protocol_common_0973_v2_analysis_0975")
    manifest_path = CAPTURE / "capture-manifest-v2.json"
    manifest = load_json(manifest_path)
    plan = load_json(PLAN)
    children = [parse_child(child, parser, manifest) for child in (1, 2)]
    root_terminal_path = ROOT / "receipts/root-terminal-0975-v1.json"
    root_go_path = ROOT / "receipts/root-go-0975-v1.json"
    runner_log_path = ROOT / "runner-0975-v1.log"
    parser_source = binding(PARSER_PATH)
    adapter_evidence = source_order_evidence()
    parser_evidence = parser_coverage_evidence(PARSER_PATH)
    root_terminal = load_json(root_terminal_path)
    root_go = load_json(root_go_path)
    runner_log = runner_log_path.read_text(encoding="utf-8", errors="replace").splitlines()

    raw_sums = [child["raw_stack"].get("raw_lost_sum") for child in children]
    raw_records = [child["raw_stack"].get("raw_lost_record_count") for child in children]
    terminal_sums = [child["terminal"].get("lost_samples_summary", {}).get("lost_samples") for child in children]
    packet_children = [child["child"] for child in children if child["packet_present"]]
    output = {
        "schema": "litchi.performance.0975.failed_capture_analysis.v1",
        "status": "complete_descriptive_only",
        "created_utc": "2026-10-08T00:00:00Z",
        "scope": {
            "capture_root": str(CAPTURE),
            "children_examined": [1, 2],
            "children_requested": manifest.get("children_requested"),
            "children_not_launched": root_terminal.get("children_not_launched"),
            "all_12_reduction_performed": False,
            "native_or_build_invoked_by_this_tool": False,
            "performance_or_adoption_claim": False,
        },
        "custody": {
            "plan": binding(PLAN),
            "design": binding(DESIGN),
            "manifest": binding(manifest_path),
            "root_go": binding(root_go_path),
            "root_terminal": binding(root_terminal_path),
            "runner_log": binding(runner_log_path),
            "decoder_parser": parser_source,
            "decoder_empty_chain_definition": parser_evidence,
            "write_admission_order": adapter_evidence,
        },
        "capture_outcome": {
            "manifest_status": manifest.get("status"),
            "manifest_failure": manifest.get("failure"),
            "root_terminal_status": root_terminal.get("status"),
            "runner_returncode": root_terminal.get("runner_returncode"),
            "native_children_launched": root_terminal.get("native_children_launched"),
            "no_retry": root_terminal.get("no_retry"),
            "runner_log_lines": runner_log,
            "packet_children": packet_children,
        },
        "raw_loss_audit": {
            "children": [
                {
                    "child": child["child"],
                    "raw_perf_record_lost_record_count": child["raw_stack"].get("raw_lost_record_count"),
                    "raw_perf_record_lost_sum": child["raw_stack"].get("raw_lost_sum"),
                    "terminal_perf_record_lost_samples": child["terminal"].get("lost_samples_summary", {}).get("lost_samples"),
                    "terminal_trailer": child["terminal"].get("lost_samples_summary"),
                    "raw_and_terminal_are_separate": True,
                    "raw_gate_passed": child["raw_stack"].get("raw_lost_sum") == 0,
                    "terminal_gate_passed": child["terminal"].get("lost_samples_summary", {}).get("lost_samples") == 0,
                }
                for child in children
            ],
            "raw_record_count_vector": raw_records,
            "raw_sum_vector": raw_sums,
            "terminal_lost_samples_vector": terminal_sums,
            "child_01_zero_loss": raw_sums[0] == 0 and terminal_sums[0] == 0,
            "child_02_zero_loss": raw_sums[1] == 0 and terminal_sums[1] == 0,
            "child_02_packet_absence_reason": "raw PERF_RECORD_LOST sum is nonzero; complete decoder outputs exist, then packet admission require_zero_loss raised before packet write",
        },
        "children": children,
        "interpretation": {
            "child_01": "one admitted child with descriptive decoder coverage; no cohort or speed claim",
            "child_02": "decoder outputs are retained for audit only; raw loss 184 makes the child inadmissible and no packet was written",
            "empty_callchains": "remaining empty counts quantify exported/decoded frame coverage only; no cause is assigned",
            "owners": "corrected physical/deepest tables are bounded descriptive observations; inclusive owner rows overlap and are not additive",
        },
    }
    ANALYSIS.mkdir(parents=True, exist_ok=True)
    # Avoid circular self-binding: write once, then bind this immutable tool in
    # the receipt and write the receipt exclusively.
    tool_binding = binding(Path(__file__))
    output["custody"]["analysis_helper"] = tool_binding
    output["created_utc"] = __import__("datetime").datetime.now(__import__("datetime").timezone.utc).isoformat()
    if RECEIPT.exists():
        raise RuntimeError(f"refusing to overwrite existing receipt: {RECEIPT}")
    with RECEIPT.open("x", encoding="utf-8") as handle:
        json.dump(output, handle, indent=2, sort_keys=True, allow_nan=False)
        handle.write("\n")
    print(json.dumps({
        "status": output["status"],
        "receipt": str(RECEIPT),
        "receipt_sha256": sha256(RECEIPT),
        "helper": tool_binding,
        "children": [{"child": row["child"], "packet_present": row["packet_present"], "raw_loss": row["raw_stack"]["raw_lost_sum"], "terminal_loss": row["terminal"]["lost_samples_summary"]["lost_samples"]} for row in children],
    }, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
