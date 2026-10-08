#!/usr/bin/env python3
"""Fail-closed offline reducer for the 0977 corrected-only cohort.

The capture runner hands this reducer one completed packet per child.  The
decoder packet contains only bounded summaries; this reducer reparses every
corrected ``perf script`` stream and then passes the canonical parsed rows to
the retained 0973 arithmetic reducer.  No stock/no-op replay is required in
0977: the final 0976 decoder controls are bound by the plan and are checked
through their final log schema here.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Iterable


ROOT = Path("/home/zhuhe/litchi-performance-0977")
BASE_TOOLS = Path("/home/zhuhe/litchi-performance-0973/tools-v2")
sys.path.insert(0, str(BASE_TOOLS))
import protocol_common_0973_v2 as p  # noqa: E402


def _load_base_reducer() -> Any:
    path = BASE_TOOLS / "reduce_0973_v2.py"
    spec = importlib.util.spec_from_file_location("litchi_reduce_0973_v2_for_0977", path)
    if spec is None or spec.loader is None:
        raise p.ProtocolError(f"cannot load retained reducer: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


BASE = _load_base_reducer()

SCHEMA = "litchi.performance.0977.instruction_profile_analysis.v1"
PACKET_SCHEMA = "litchi.performance.0977.child_packet.v1"
MANIFEST_SCHEMA = "litchi.performance.0977.capture_manifest.v1"
FINAL_REPORT_SCHEMA = "litchi.performance.0976.dwfl_report.final"
FINAL_INFO_SCHEMA = "litchi.performance.0976.dwfl_info.final"
CHILDREN = 12
CPU = 31
PERIOD = 1_000_000
WARMUPS = 5
MEASURED = 80
CALLS = WARMUPS + MEASURED
STACK_CAP = 32_768

EXPECTED_SOURCE_MANIFEST = {
    "bytes": 14960,
    "path": "/home/zhuhe/litchi-performance-0969/manifests/source-manifest-v1.json",
    "sha256": "f1818af5c306fb6ea01b59ca8df29710e773386b246071202bd4aa896f4e8e2b",
}
EXPECTED_MARKER = {
    "build_id": "e5c457384eeebd057932c444efda77618f137399",
    "calls": 85,
    "dso_path": "/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile",
    "marker_scope": "5 warmups + 80 measured calls combined",
    "owner_scope": "all 12 children; missing owner membership unknown",
    "p50_scope": "80 measured samples only",
    "symbol": "doc_retention_probe_0942_cfb_profile::profiled_measured_case",
}
EXPECTED_PILOT_GATE = {
    "child": 1,
    "lost_events": 0,
    "marker_period_share_positive": True,
    "min_marker_samples": 100,
    "on_failure": "stop_without_retry",
}

REQUIRED_SUMMARY_KEYS = {
    "sample_count",
    "total_period",
    "marker",
    "marker_sample_count",
    "marker_period",
    "marker_period_share",
    "lost_events",
    "unresolved_frames",
    "explicit_stack_truncated_count",
    "empty_callchain_count",
    "physical_leaf_owner_period",
    "physical_leaf_unknown_period",
    "physical_leaf_unknown_samples",
    "marker_unknown_physical_leaf_samples",
    "marker_no_resolved_frame_samples",
    "physical_leaf_classified_period",
    "physical_leaf_conservation",
    "physical_leaf_scope",
    "deepest_resolved_proxy_period",
    "inclusive_marker_ancestor_period",
}
ALLOWED_SUMMARY_EXTRAS = {
    "raw_lost_sum",
    "terminal_lost_samples",
    "decoder_schema",
    "decoder_version",
    "corrected_script_sha256",
    "corrected_log_sha256",
    "summary_scope",
}


def _error(message: str) -> p.ProtocolError:
    return p.ProtocolError(message)


def _load(path: Path) -> dict[str, Any]:
    return p.load_json(path)


def _check_binding(row: Any, label: str) -> Path:
    return p.check_binding(row, label)


def _path_from_binding(row: Any, label: str) -> Path:
    return _check_binding(row, label)


def _same_json(left: Any, right: Any) -> bool:
    # JSON values in these summaries are all scalar/list/dict values.  Keeping
    # the comparison structural catches a repaired decoder count copied into a
    # packet without making floating-point tolerances hide a mismatch.
    return left == right


def _require_plan_contract(plan: dict[str, Any], plan_path: Path) -> None:
    if plan.get("status") != "pre_capture":
        raise _error("0977 plan is not the immutable pre-capture plan")
    raw_plan_path = plan.get("plan_path")
    if raw_plan_path != str(plan_path.absolute()):
        raise _error("0977 plan self-path binding mismatch")
    if plan.get("source_manifest") != EXPECTED_SOURCE_MANIFEST:
        raise _error("source_manifest is not the exact retained 0975 object")
    if plan.get("marker") != EXPECTED_MARKER:
        raise _error("marker contract is not the exact retained 0975 object")
    if plan.get("pilot_gate") != EXPECTED_PILOT_GATE:
        raise _error("pilot_gate is not the exact retained 0975 object")
    if plan.get("child_gate") != {
        "all_children": True, "raw_lost_sum": 0, "terminal_lost_samples": 0,
        "min_marker_samples": 100, "marker_period_share_positive": True,
        "empty_callchain_count": 0, "on_failure": "stop_without_retry",
    }:
        raise _error("child_gate is not the exact all-child stop gate")
    batch = plan.get("batch") or {}
    for key, value in {
        "children": CHILDREN,
        "serial": True,
        "cpu": CPU,
        "period": PERIOD,
        "stack_cap": STACK_CAP,
        "warmups": WARMUPS,
        "samples": MEASURED,
        "minimum_marker_samples": 100,
        "marker_period_positive": True,
        "raw_lost_sum": 0,
        "terminal_lost_samples": 0,
        "empty_callchain_count": 0,
        "stop_on_first_failure": True,
        "retry": False,
    }.items():
        if batch.get(key) != value:
            raise _error(f"batch.{key} drifted: {batch.get(key)!r}")
    perf = plan.get("perf") or {}
    if perf.get("event") != "instructions:u" or perf.get("period") != PERIOD:
        raise _error("perf event/period contract drift")
    if perf.get("callgraph") != "dwarf,32768":
        raise _error("perf callgraph contract drift")
    marker = plan["marker"]
    if marker["dso_path"] != plan.get("binary", {}).get("path"):
        raise _error("marker DSO is not the bound binary")
    decoder = plan.get("decoder_interface") or {}
    if decoder.get("offline_only") is not True:
        raise _error("decoder is not declared offline-only")
    if decoder.get("native_ld_preload") != "absent":
        raise _error("native LD_PRELOAD contract drift")
    if plan.get("terminal_contract") != {
        "field": "PERF_RECORD_LOST_SAMPLES", "required_value": 0,
        "source": "decoder-owned offline perf report --stdio",
        "runner_precomputes": False, "raw_loss_not_added": True,
    }:
        raise _error("terminal report ownership contract drift")


def _packet_artifacts(packet: dict[str, Any], packet_path: Path) -> dict[str, dict[str, Any]]:
    """Merge decoder and runner bindings while rejecting path drift."""
    merged: dict[str, dict[str, Any]] = {}
    for source_name in ("artifacts", "runner_bindings"):
        source = packet.get(source_name)
        if source is None:
            continue
        if not isinstance(source, dict):
            raise _error(f"{packet.get('packet_id')}: {source_name} is not an object")
        for key, row in source.items():
            if not isinstance(row, dict):
                raise _error(f"{packet.get('packet_id')}: artifact {key} is not a binding")
            # Validate every supplied binding, including auxiliary decoder
            # files.  This prevents the reducer from silently using an output
            # different from the one named by the packet.
            _check_binding(row, f"{packet.get('packet_id')}.{source_name}.{key}")
            previous = merged.get(key)
            if previous is not None and previous != row:
                raise _error(f"{packet.get('packet_id')}: artifact binding conflict for {key}")
            merged[key] = row
    if not merged:
        raise _error(f"{packet.get('packet_id')}: packet has no artifact bindings")
    # The runner's immutable files may be named with hyphens while the decoder
    # interface uses concise snake-case names.  Both forms are accepted only
    # as aliases for the same binding.
    for left, right in (
        ("raw_task", "raw-task.txt"),
        ("terminal_summary", "terminal-summary.json"),
        ("layout", "target_layout"),
        ("layout", "target-layout.json"),
        ("corrected_script", "corrected.perf-script-no-inline.txt"),
        ("corrected_log", "corrected.dwfl.jsonl"),
    ):
        if left in merged and right in merged and merged[left] != merged[right]:
            raise _error(f"{packet.get('packet_id')}: conflicting artifact aliases {left}/{right}")
        if left not in merged and right in merged:
            merged[left] = merged[right]
        if right not in merged and left in merged:
            merged[right] = merged[left]
    required = ("corrected_script", "corrected_log", "report", "terminal", "launch", "layout")
    missing = [name for name in required if name not in merged]
    if missing:
        raise _error(f"{packet.get('packet_id')}: missing artifact bindings {missing}")
    return merged


def _artifact_path(artifacts: dict[str, dict[str, Any]], *names: str) -> Path:
    for name in names:
        row = artifacts.get(name)
        if row is not None:
            return _path_from_binding(row, f"artifact.{name}")
    raise _error(f"missing artifact aliases {names}")


def _raw_lost_sum(text: str) -> int:
    total = 0
    for line in text.splitlines():
        # perf script -D can print both the raw PERF_RECORD_LOST records and
        # the terminal PERF_RECORD_LOST_SAMPLES trailer.  The latter is a
        # separate gate and must never be added to the raw event count.
        if "PERF_RECORD_LOST_SAMPLES" in line:
            continue
        if re.search(r"PERF_RECORD_LOST(?:\s|:|$)", line) is None:
            continue
        values = re.findall(r"\b(?:lost|events?)\s*[:=]?\s*(\d+)\b", line, flags=re.IGNORECASE)
        if not values:
            values = re.findall(r"\((\d+)\s+events?\)", line, flags=re.IGNORECASE)
        if not values:
            raise _error(f"unparseable PERF_RECORD_LOST line: {line[:240]}")
        total += int(values[-1])
    return total


def _check_loss(packet: dict[str, Any], artifacts: dict[str, dict[str, Any]]) -> dict[str, int]:
    raw_path = _artifact_path(artifacts, "raw_task", "raw-task.txt")
    raw_sum = _raw_lost_sum(raw_path.read_text(encoding="utf-8", errors="strict"))
    terminal_summary_path = _artifact_path(artifacts, "terminal_summary", "terminal-summary.json")
    terminal_summary = _load(terminal_summary_path)
    terminal_lost = terminal_summary.get("lost_samples")
    if type(terminal_lost) is not int:
        raise _error(f"{packet['packet_id']}: terminal lost_samples is missing/non-integer")
    if raw_sum != 0 or terminal_lost != 0:
        raise _error(f"{packet['packet_id']}: loss gate failed ({raw_sum}, {terminal_lost})")
    loss_row = packet.get("loss")
    loss_path = artifacts.get("loss.json") or artifacts.get("loss")
    if loss_path is not None:
        recorded = _load(_path_from_binding(loss_path, f"{packet['packet_id']}.loss"))
        if recorded.get("raw_lost_sum") != raw_sum or recorded.get("terminal_lost_samples") != terminal_lost:
            raise _error(f"{packet['packet_id']}: loss receipt disagrees with raw/terminal evidence")
    if isinstance(loss_row, dict):
        if loss_row.get("raw_lost_sum") != raw_sum or loss_row.get("terminal_lost_samples") != terminal_lost:
            raise _error(f"{packet['packet_id']}: packet loss summary disagrees with raw/terminal evidence")
    return {"raw_lost_sum": raw_sum, "terminal_lost_samples": terminal_lost}


def _check_report(packet: dict[str, Any], artifacts: dict[str, dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    report_path = _artifact_path(artifacts, "report")
    report_file = _load(report_path)
    report = packet.get("report")
    if report != report_file:
        raise _error(f"{packet['packet_id']}: report packet/artifact mismatch")
    elapsed = report.get("elapsed_ns")
    if not isinstance(elapsed, list) or len(elapsed) != MEASURED:
        raise _error(f"{packet['packet_id']}: report must contain exactly 80 measured elapsed values")
    if any(isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or float(value) < 0 for value in elapsed):
        raise _error(f"{packet['packet_id']}: elapsed_ns contains an invalid value")
    contract = plan.get("report_contract") or {}
    if report.get("samples_requested", MEASURED) != MEASURED:
        raise _error(f"{packet['packet_id']}: measured sample count drift")
    if report.get("warmups", WARMUPS) != WARMUPS:
        raise _error(f"{packet['packet_id']}: warmup count drift")
    if report.get("marker_calls", CALLS) != CALLS:
        raise _error(f"{packet['packet_id']}: marker call count drift")
    if contract.get("case") is not None and report.get("case") != contract["case"]:
        raise _error(f"{packet['packet_id']}: report case drift")
    # The retained harness exposes per-sample rows.  Check their identity when
    # present, but do not manufacture those rows if the analyzer intentionally
    # returns only the bounded elapsed vector.
    rows = report.get("samples")
    if rows is not None:
        if not isinstance(rows, list) or len(rows) != MEASURED:
            raise _error(f"{packet['packet_id']}: report samples are not 80 rows")
        for index, row in enumerate(rows):
            if not isinstance(row, dict) or row.get("index") != index or row.get("elapsed_ns") != elapsed[index]:
                raise _error(f"{packet['packet_id']}: report sample identity drift at {index}")
    return report


def _expected_native_record_argv(plan: dict[str, Any], child_dir: Path) -> list[str]:
    """Reconstruct the sealed native argv for an independent custody check."""
    binary = str(Path(plan["binary"]["path"]).resolve())
    fixture = str(Path(plan["fixture"]["path"]).resolve())
    python = str(Path(plan["python"]["path"]).resolve())
    launcher = str(Path(plan["launcher"]["path"]).resolve())
    bridge = str(Path(plan["helpers"]["bridge"]["path"]).resolve())
    child = child_dir.name
    data = str((child_dir / f"{child}.perf.data").resolve())
    bridge_receipt = str((child_dir / f"{child}.bridge.json").resolve())
    launch = str((child_dir / f"{child}.launch.json").resolve())
    target = [python, launcher, "--metadata", launch, "--requested-affinity", "[31]", "--", binary,
              "--case", "docfloat", "--input", fixture, "--edits", "1", "--retention", "default",
              "--warmup", "5", "--samples", "80"]
    bridge_args = [python, bridge, "--expected-perf-exe", str(Path(plan["perf"]["path"]).resolve()),
                   "--expected-perf-sha256", plan["perf"]["sha256"], "--expected-target-exe", python,
                   "--expected-target-sha256", plan["python"]["sha256"], "--expected-target-script", launcher,
                   "--expected-target-script-sha256", plan["launcher"]["sha256"], "--bridge-receipt", bridge_receipt,
                   "--target-cpu", "31", "--recorder-cpu", "30", "--", *target]
    return [str(Path(plan["taskset"]["path"]).resolve()), "--cpu-list", "31",
            str(Path(plan["perf"]["path"]).resolve()), "record", "--per-thread", "-e", "instructions:u",
            "-c", "1000000", "-m", "16M", "--call-graph", "dwarf,32768", "--period", "--timestamp",
            "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--user-callchains", "-o", data,
            "--", *bridge_args]


def _check_runner_custody(packet: dict[str, Any], artifacts: dict[str, dict[str, Any]],
                          terminal: dict[str, Any], plan: dict[str, Any], packet_path: Path) -> None:
    runner_gate = packet.get("runner_gate")
    if not isinstance(runner_gate, dict) or runner_gate.get("passed") is not True:
        raise _error(f"{packet['packet_id']}: runner gate is not passing")
    affinity_gate = runner_gate.get("affinity")
    if not isinstance(affinity_gate, dict) or affinity_gate.get("passed") is not True:
        raise _error(f"{packet['packet_id']}: sampled affinity gate is not passing")
    cleanup = runner_gate.get("process_group_cleanup")
    if not isinstance(cleanup, dict) or cleanup.get("passed") is not True:
        raise _error(f"{packet['packet_id']}: process-group cleanup gate is not passing")
    for key in ("affinity_gate", "process_group_cleanup"):
        if packet.get("terminal", {}).get(key) != terminal.get(key):
            raise _error(f"{packet['packet_id']}: terminal {key} changed from immutable artifact")
    command = _load(_artifact_path(artifacts, "command.json", "command"))
    expected = _expected_native_record_argv(plan, packet_path.parent)
    if command.get("argv") != expected or terminal.get("command") != expected:
        raise _error(f"{packet['packet_id']}: native command differs from sealed plan argv")
    if command.get("cwd") != (plan.get("command_receipt_contract") or {}).get("cwd"):
        raise _error(f"{packet['packet_id']}: command cwd differs from bound plan cwd")
    observation_row = next((row for name, row in artifacts.items() if name.endswith(".affinity-observations.json")), None)
    if observation_row is None:
        raise _error(f"{packet['packet_id']}: affinity observations are not bound")
    observations = _load(_path_from_binding(observation_row, f"{packet['packet_id']}.affinity-observations"))
    rows = observations.get("rows")
    if not isinstance(rows, list) or not any(isinstance(row, dict) and row.get("bridge_handoff_ready") is True for row in rows):
        raise _error(f"{packet['packet_id']}: no post-handoff observation row is retained")


def _check_launch_terminal(packet: dict[str, Any], artifacts: dict[str, dict[str, Any]], plan: dict[str, Any], packet_path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    terminal = _load(_artifact_path(artifacts, "terminal"))
    packet_terminal = packet.get("terminal") or {}
    for key in ("command", "returncode", "timed_out", "native_ld_preload"):
        if packet_terminal.get(key) != terminal.get(key):
            raise _error(f"{packet['packet_id']}: terminal {key} changed from immutable artifact")
    if terminal.get("returncode") != 0 or terminal.get("timed_out") is not False or terminal.get("native_ld_preload") != "absent":
        raise _error(f"{packet['packet_id']}: terminal execution gate failed")
    _check_runner_custody(packet, artifacts, terminal, plan, packet_path)
    launch = _load(_artifact_path(artifacts, "launch"))
    if packet.get("launch") != launch:
        raise _error(f"{packet['packet_id']}: launch packet/artifact mismatch")
    binary = plan.get("binary") or {}
    if launch.get("execve") is not True or launch.get("exec_path") != str(Path(binary["path"]).resolve()):
        raise _error(f"{packet['packet_id']}: exact native exec binding failed")
    if launch.get("pid") != launch.get("exec_pid") or launch.get("actual_affinity") != [CPU] or launch.get("requested_affinity") != [CPU]:
        raise _error(f"{packet['packet_id']}: native PID/affinity binding failed")
    command = _load(_artifact_path(artifacts, "command.json", "command"))
    if command.get("argv") != terminal.get("command"):
        raise _error(f"{packet['packet_id']}: terminal command differs from command receipt")
    argv = command.get("argv")
    if not isinstance(argv, list):
        raise _error(f"{packet['packet_id']}: command argv is not a list")
    required_tokens = ("--per-thread", "instructions:u", "-m", "16M", "--call-graph", "dwarf,32768")
    if any(token not in argv for token in required_tokens):
        raise _error(f"{packet['packet_id']}: command sampling tokens are incomplete")
    if any(token in argv for token in ("-a", "--all-cpus", "--inherit")):
        raise _error(f"{packet['packet_id']}: forbidden system-wide/inherit perf token")
    return launch, terminal


def _check_layout(packet: dict[str, Any], artifacts: dict[str, dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    layout = _load(_artifact_path(artifacts, "layout"))
    if packet.get("layout") != layout:
        raise _error(f"{packet['packet_id']}: layout packet/artifact mismatch")
    binary = plan.get("binary") or {}
    layout_binary = layout.get("binary") or {}
    if layout_binary.get("path") != binary.get("path") or layout_binary.get("sha256") != binary.get("sha256"):
        raise _error(f"{packet['packet_id']}: layout binary custody mismatch")
    if not isinstance(layout.get("corrected_base"), str) or not isinstance(layout.get("perf_expected_start"), str):
        raise _error(f"{packet['packet_id']}: layout bases are missing")
    return layout


def _check_final_log(path: Path, layout: dict[str, Any], binary_path: str) -> dict[str, Any]:
    counts = Counter()
    reports: list[dict[str, Any]] = []
    infos: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8", errors="strict").splitlines(), 1):
        if not line.strip():
            continue
        try:
            row = json.loads(line)
        except json.JSONDecodeError as exc:
            raise _error(f"final decoder log line {line_number} is not JSON") from exc
        schema = row.get("schema")
        if schema == FINAL_REPORT_SCHEMA:
            reports.append(row)
        elif schema == FINAL_INFO_SCHEMA:
            infos.append(row)
        else:
            raise _error(f"unexpected final decoder log schema: {schema!r}")
        counts[str(schema)] += 1
        if row.get("mode") != "corrected" or row.get("config_valid") is not True:
            raise _error("final decoder log has a non-corrected/config-invalid row")
    exact_reports = [row for row in reports if row.get("path_match") is True and row.get("name_match") is True]
    if not exact_reports:
        raise _error("final decoder log has no exact target report row")
    expected_base = layout["corrected_base"]
    expected_start = layout["perf_expected_start"]
    for row in exact_reports:
        if row.get("file_name") != binary_path:
            raise _error("final target report path differs from bound binary")
        if row.get("admission") is not True or row.get("identity_valid") is not True:
            raise _error("final target report was not admitted with verified identity")
        if row.get("caller_is_perf") is not True or row.get("caller_file") != "/usr/bin/perf":
            raise _error("final target report caller is not perf")
        if row.get("forwarded_base") != expected_base or row.get("original_base") != expected_start:
            raise _error("final target report base correction differs from layout")
        if str(row.get("result_module")) in {"(nil)", "None", "null"}:
            raise _error("final target report has a null result module")
    stored = [row for row in exact_reports if row.get("cache_stored") is True]
    reused = [row for row in exact_reports if row.get("reused_existing") is True]
    initial = [row for row in exact_reports if row.get("initial_rewrite") is True]
    # ``initial_rewrite`` is a per-call diagnostic in the final shim and is
    # true on every exact row in the retained qualified replay.  The cache
    # custody gates are the stable ones: exactly one store and every remaining
    # exact row reuses the verified entry.
    if len(stored) != 1 or len(reused) != len(exact_reports) - 1:
        raise _error("final target report cache admission/reuse counts are invalid")
    target_infos = [row for row in infos if row.get("target_match") is True]
    if not target_infos:
        raise _error("final decoder log has no target info row")
    for row in target_infos:
        if row.get("caller_is_perf") is not True or row.get("caller_file") != "/usr/bin/perf":
            raise _error("final target info caller is not perf")
        if row.get("current_range_valid") is not True:
            raise _error("final target info range is not valid")
    return {
        "schema": "litchi.performance.0976.final-v2",
        "file": p.file_binding(path),
        "rows": sum(counts.values()),
        "report_rows": len(reports),
        "info_rows": len(infos),
        "exact_target_reports": len(exact_reports),
        "exact_target_infos": len(target_infos),
        "cache_stored_rows": len(stored),
        "initial_rewrite_rows": len(initial),
        "reused_existing_rows": len(reused),
        "non_target_report_rows": len(reports) - len(exact_reports),
        "checks": {
            "only_final_schemas": True,
            "all_corrected_config_valid": True,
            "exact_target_admitted": True,
            "exact_target_identity_valid": True,
            "one_cache_store": True,
            "remaining_rows_reused": True,
            "target_info_present": True,
        },
    }


def _compare_decoder_summary(packet: dict[str, Any], parsed: dict[str, Any], losses: dict[str, int]) -> dict[str, Any]:
    perf = packet.get("perf")
    if not isinstance(perf, dict) or set(perf) != {"corrected"}:
        raise _error(f"{packet['packet_id']}: 0977 packet must contain corrected-only perf summary")
    summary = perf.get("corrected")
    if not isinstance(summary, dict) or "samples" in summary:
        raise _error(f"{packet['packet_id']}: decoder summary is missing/bloated with sample rows")
    missing = sorted(REQUIRED_SUMMARY_KEYS - set(summary))
    if missing:
        raise _error(f"{packet['packet_id']}: decoder summary missing {missing}")
    for key, value in summary.items():
        if key in ALLOWED_SUMMARY_EXTRAS:
            continue
        if key not in parsed:
            raise _error(f"{packet['packet_id']}: unexpected decoder summary key {key}")
        if not _same_json(value, parsed[key]):
            raise _error(f"{packet['packet_id']}: repaired decoder count differs at {key}")
    for key in REQUIRED_SUMMARY_KEYS:
        if not _same_json(summary[key], parsed[key]):
            raise _error(f"{packet['packet_id']}: summary/parser mismatch at {key}")
    if summary.get("raw_lost_sum", losses["raw_lost_sum"]) != losses["raw_lost_sum"]:
        raise _error(f"{packet['packet_id']}: decoder raw loss differs from raw task stream")
    if summary.get("terminal_lost_samples", losses["terminal_lost_samples"]) != losses["terminal_lost_samples"]:
        raise _error(f"{packet['packet_id']}: decoder terminal loss differs from terminal summary")
    if parsed.get("lost_events") != 0:
        raise _error(f"{packet['packet_id']}: corrected parser reports nonzero loss")
    return summary


def _decode_packet(packet: dict[str, Any], packet_path: Path, plan: dict[str, Any]) -> dict[str, Any]:
    child_id = packet.get("packet_id")
    artifacts = _packet_artifacts(packet, packet_path)
    launch, terminal = _check_launch_terminal(packet, artifacts, plan, packet_path)
    _check_report(packet, artifacts, plan)
    layout = _check_layout(packet, artifacts, plan)
    losses = _check_loss(packet, artifacts)
    gate = packet.get("gate")
    if not isinstance(gate, dict) or gate.get("passed") is not True or gate.get("reasons") not in ([], None):
        raise _error(f"{child_id}: packet gate is not passed")
    script_path = _artifact_path(artifacts, "corrected_script", "corrected.perf-script-no-inline.txt")
    log_path = _artifact_path(artifacts, "corrected_log", "corrected.dwfl.jsonl")
    parsed = p.parse_perf_script_strict(
        script_path.read_text(encoding="utf-8", errors="strict"),
        marker=plan["marker"]["symbol"], marker_dso=plan["binary"]["path"],
        expected_pid=int(launch["exec_pid"]), expected_cpu=CPU, expected_period=PERIOD,
    )
    summary = _compare_decoder_summary(packet, parsed, losses)
    log_stats = _check_final_log(log_path, layout, plan["binary"]["path"])
    if summary.get("corrected_script_sha256") is not None and summary["corrected_script_sha256"] != p.sha256(script_path):
        raise _error(f"{child_id}: corrected script digest in summary is wrong")
    if summary.get("corrected_log_sha256") is not None and summary["corrected_log_sha256"] != p.sha256(log_path):
        raise _error(f"{child_id}: corrected log digest in summary is wrong")
    raw_stack: dict[str, Any]
    raw_stack_path = artifacts.get("raw_stack") or artifacts.get("raw-stack.json")
    if raw_stack_path is not None:
        raw_stack = _load(_path_from_binding(raw_stack_path, f"{child_id}.raw_stack"))
        if packet.get("raw_stack") is not None and packet.get("raw_stack") != raw_stack:
            raise _error(f"{child_id}: raw stack packet/artifact mismatch")
    elif isinstance(packet.get("raw_stack"), dict):
        raw_stack = packet["raw_stack"]
    else:
        raw_stack = {
            "schema": "litchi.performance.0977.raw_stack_summary.v1",
            "status": "not_returned_by_decoder",
            "scope": "corrected parser metrics only; raw STACK_USER distributions unavailable in packet",
        }
    normalized = dict(packet)
    normalized["artifacts"] = artifacts
    normalized["_packet_path"] = str(packet_path.absolute())
    normalized["_artifacts"] = {key: _path_from_binding(row, f"{child_id}.{key}") for key, row in artifacts.items()}
    normalized["_parsed"] = {"corrected": parsed}
    normalized["_raw_stack"] = raw_stack
    normalized["_log_stats"] = {"corrected": log_stats}
    normalized["_loss"] = losses
    normalized["_decoder_summary"] = summary
    normalized["_terminal"] = terminal
    normalized["_layout"] = layout
    return normalized


def _load_packets(plan_path: Path, manifest_path: Path, plan: dict[str, Any]) -> list[dict[str, Any]]:
    manifest = _load(manifest_path)
    if manifest.get("schema") != MANIFEST_SCHEMA or manifest.get("status") != "complete" or manifest.get("no_retry") is not True:
        raise _error("0977 manifest is not complete/no-retry")
    if manifest.get("serial") is not True or manifest.get("failure") is not None:
        raise _error("0977 manifest has a failure or nonserial capture")
    if manifest.get("plan") != p.file_binding(plan_path):
        raise _error("0977 manifest plan binding mismatch")
    paths = manifest.get("packets")
    if manifest.get("children_requested") != CHILDREN or manifest.get("children_completed") != CHILDREN or not isinstance(paths, list) or len(paths) != CHILDREN or len(set(paths)) != CHILDREN:
        raise _error("0977 manifest does not contain exactly 12 unique packets")
    packets: list[dict[str, Any]] = []
    for index, raw_path in enumerate(paths, 1):
        packet_path = Path(raw_path).absolute()
        packet = _load(packet_path)
        child_id = f"child-{index:02d}"
        if packet.get("schema") != PACKET_SCHEMA or packet.get("status") != "complete" or packet.get("packet_id") != child_id:
            raise _error(f"{child_id}: packet schema/status/identity mismatch")
        if packet.get("gate", {}).get("passed") is not True:
            raise _error(f"{child_id}: packet gate failed")
        packets.append(_decode_packet(packet, packet_path, plan))
    return packets


def _correlations(values: list[float], p50s: list[float], p95s: list[float]) -> dict[str, Any]:
    return {
        "values_all_12": values,
        "vs_T50": {"pearson": p.pearson(values, p50s), "spearman": p.spearman(values, p50s)},
        "vs_T95": {"pearson": p.pearson(values, p95s), "spearman": p.spearman(values, p95s)},
    }


def _augment_proxy_associations(result: dict[str, Any], packets: list[dict[str, Any]]) -> None:
    """Add all-12 deepest-proxy vectors, shares, and timing associations.

    The retained reducer already reports proxy totals.  This narrow addition
    makes its child-spread output as explicit as the physical-leaf and
    inclusive views without treating the proxy as an exact self-cost owner.
    """
    rows = result.get("children") or []
    p50s = [float(row["p50_ns"]) for row in rows]
    p95s = [float(row["p95_ns"]) for row in rows]
    m_periods = [float(row["marker_period"]) for row in rows]
    names = sorted({owner for packet in packets for owner in packet["_parsed"]["corrected"].get("deepest_resolved_proxy_period", {})})
    associations: dict[str, Any] = {}
    for owner in names:
        periods = [float(packet["_parsed"]["corrected"].get("deepest_resolved_proxy_period", {}).get(owner, 0.0)) for packet in packets]
        shares = [period / marker if marker > 0 else 0.0 for period, marker in zip(periods, m_periods)]
        present = [period > 0 for period in periods]
        associations[owner] = {
            "periods_all_12": periods,
            "shares_all_12": shares,
            "present_all_12": present,
            "presence_count": sum(present),
            "missing_children": [row["packet_id"] for row, present in zip(rows, present) if not present],
            "period_association": {"T50": _correlations(periods, p50s, p95s), "T95": _correlations(periods, p50s, p95s)},
            "share_association": {"T50": _correlations(shares, p50s, p95s), "T95": _correlations(shares, p50s, p95s)},
            "p50_vs_period": {"pearson": p.pearson(p50s, periods), "spearman": p.spearman(p50s, periods)},
            "p95_vs_period": {"pearson": p.pearson(p95s, periods), "spearman": p.spearman(p95s, periods)},
            "p50_vs_share": {"pearson": p.pearson(p50s, shares), "spearman": p.spearman(p50s, shares)},
            "p95_vs_share": {"pearson": p.pearson(p95s, shares), "spearman": p.spearman(p95s, shares)},
            "scope": "all 12 children; first resolved frame proxy across all DSOs; not physical frame-0 self-cost; descriptive only",
        }
    summary = result.setdefault("summary", {})
    summary["deepest_resolved_proxy_associations"] = associations
    summary.setdefault("associations_descriptive", {})["deepest_resolved_proxy"] = associations


def _reduce_packets(plan_path: Path, manifest_path: Path, plan: dict[str, Any], packets: list[dict[str, Any]]) -> dict[str, Any]:
    result = BASE._reduce_packets(plan_path, manifest_path, plan, packets)
    _augment_proxy_associations(result, packets)
    result["schema"] = SCHEMA
    result["batch"] = {
        "children": CHILDREN,
        "serial": True,
        "cpu": CPU,
        "period": PERIOD,
        "stack_cap": STACK_CAP,
        "warmups": WARMUPS,
        "measured_samples": MEASURED,
        "marker_calls": CALLS,
        "corrected_only": True,
        "stock_noop_replay": "not required; qualified 0976 final decoder controls are bound by plan",
    }
    result["decoder_contract"] = {
        "mode": "0976 final-v2 corrected-only offline decoder",
        "log_schemas": [FINAL_REPORT_SCHEMA, FINAL_INFO_SCHEMA],
        "script_reparsed_by_reducer": True,
        "packet_summary_has_no_samples": True,
        "counts_source": "fresh strict parse of each corrected script; packet summary must match",
        "aslr_and_noop_controls": "qualified 0976 receipts bound by plan; no per-child stock/noop replay",
    }
    result["attribution_contract"] = {
        "D": "all decoded instructions:u period per child",
        "M": "period from marker-containing samples only",
        "R": "D-M; marker ancestry not observed (may include missing/truncated ancestry or wrapper/setup work)",
        "timing": "T50/T95 are descriptive wall-time statistics over 80 measured samples; no speedup claim",
        "marker_scope": "5 warmups + 80 measured calls; owner denominators are marker period M",
        "inclusive": "resolved marker-chain owners deduplicated per sample; overlap is intentional and never summed",
        "physical_leaf": "frame-0 resolved leaf across all DSOs plus explicit unknown; disjoint buckets conserve M",
        "deepest_proxy": "first resolved frame across all DSOs; separate proxy, not physical leaf/self-cost",
        "unresolved": "empty callchains, explicit truncation, unresolved frames, unknown physical leaf and marker-no-resolved counts remain explicit",
        "scope": "all 12 children; descriptive diagnostic; no exact timed-region or causal owner claim",
    }
    result["owner_selection"] = {
        "status": "not_automatic",
        "required_views": ["inclusive_owner", "physical_leaf_frame0", "physical_leaf_unknown", "deepest_resolved_proxy"],
        "required_child_vectors": "all 12; no fast/slow subset",
        "required_quality_denominators": ["marker_period", "marker_sample_count", "unresolved_frames", "explicit_stack_truncated_count", "empty_callchain_count", "physical_leaf_unknown_samples"],
        "decision_rule": "A reviewer may nominate a source owner only after checking presence/spread and direction across all 12 plus agreement of applicable views; reducer does not nominate one.",
        "claim_limit": "instruction-period association is descriptive and does not imply wall-time speedup or causality",
    }
    result["interpretation"] = result.get("interpretation", "") + " 0977 reparses corrected-only final-v2 streams and retains explicit unresolved denominators; no source owner or speedup is selected automatically."
    return result


def _synthetic_final_log_test() -> None:
    import tempfile

    with tempfile.TemporaryDirectory(prefix="litchi-0977-reducer-") as temp:
        root = Path(temp)
        binary = "/tmp/doc_retention_probe_0942_cfb_profile"
        layout = {"corrected_base": "0x1000", "perf_expected_start": "0x2000"}
        rows = [
            {"schema": FINAL_REPORT_SCHEMA, "mode": "corrected", "path_match": True, "name_match": True, "config_valid": True, "file_name": binary, "admission": True, "identity_valid": True, "caller_is_perf": True, "caller_file": "/usr/bin/perf", "forwarded_base": "0x1000", "original_base": "0x2000", "result_module": "0x1", "initial_rewrite": True, "cache_stored": True, "reused_existing": False},
            {"schema": FINAL_REPORT_SCHEMA, "mode": "corrected", "path_match": True, "name_match": True, "config_valid": True, "file_name": binary, "admission": True, "identity_valid": True, "caller_is_perf": True, "caller_file": "/usr/bin/perf", "forwarded_base": "0x1000", "original_base": "0x2000", "result_module": "0x1", "initial_rewrite": False, "cache_stored": False, "reused_existing": True},
            {"schema": FINAL_INFO_SCHEMA, "mode": "corrected", "target_match": True, "config_valid": True, "caller_is_perf": True, "caller_file": "/usr/bin/perf", "current_range_valid": True},
        ]
        path = root / "final.jsonl"
        path.write_text("".join(json.dumps(row) + "\n" for row in rows), encoding="utf-8")
        stats = _check_final_log(path, layout, binary)
        if stats["exact_target_reports"] != 2 or stats["exact_target_infos"] != 1:
            raise _error("final log self-test counts are wrong")


def self_test() -> int:
    values = [1.0, 2.0, 3.0, 4.0]
    if p.percentile(values, 0.5) != 2.5 or p.pearson(values, values) != 1.0 or p.spearman(values, list(reversed(values))) != -1.0:
        raise _error("retained arithmetic self-test failed")
    if _raw_lost_sum("x: PERF_RECORD_LOST lost 4\ny: PERF_RECORD_LOST_SAMPLES lost samples :8\n") != 4:
        raise _error("raw/terminal loss separation self-test failed")
    _synthetic_final_log_test()
    print(json.dumps({
        "schema": "litchi.performance.0977.reducer_self_test.v1",
        "status": "pass",
        "workload_started": False,
        "corrected_only": True,
        "final_log_schemas": [FINAL_REPORT_SCHEMA, FINAL_INFO_SCHEMA],
        "children": CHILDREN,
    }, indent=2, sort_keys=True))
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--plan", type=Path)
    ap.add_argument("--manifest", type=Path)
    ap.add_argument("--output", type=Path)
    args = ap.parse_args()
    try:
        if args.self_test:
            if args.plan or args.manifest or args.output:
                raise _error("--self-test is standalone")
            return self_test()
        if args.plan is None or args.manifest is None or args.output is None:
            raise _error("--plan, --manifest, and --output are required")
        plan_path = args.plan.resolve()
        manifest_path = args.manifest.resolve()
        plan = _load(plan_path)
        _require_plan_contract(plan, plan_path)
        packets = _load_packets(plan_path, manifest_path, plan)
        result = _reduce_packets(plan_path, manifest_path, plan, packets)
        p.write_json_exclusive(args.output.resolve(), result)
        print(json.dumps({"status": "pass", "output": str(args.output.resolve()), "sha256": p.sha256(args.output.resolve())}, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0977 reducer rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
