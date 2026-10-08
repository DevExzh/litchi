#!/usr/bin/env python3
"""Small 0974 adapter around the sealed 0973 decoder protocol.

The adapter deliberately keeps the 0973 packet/decoder semantics and changes
only the native perf record command's memory selector.  It also gives the
0974 run/reduce layers one loss field for the raw PERF_RECORD_LOST sum and
one terminal LOST_SAMPLES summary; those two representations are checked
independently and are never added together.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import importlib.util
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0974")
RETAINED_ROOT = Path("/home/zhuhe/litchi-performance-0973")
RETAINED_TOOLS = RETAINED_ROOT / "tools-v2"
RETAINED_HELPER_SEAL = RETAINED_ROOT / "receipts/helper-seal-0973-final-v2.json"
RETAINED_PLAN = RETAINED_ROOT / "plans/plan-0973-final-v2.json"

RETAINED_FILES = {
    "protocol": RETAINED_TOOLS / "protocol_common_0973_v2.py",
    "runner": RETAINED_TOOLS / "run_0973_v2.py",
    "reducer": RETAINED_TOOLS / "reduce_0973_v2.py",
    "raw_stack": RETAINED_TOOLS / "raw_stack_0973_v2.py",
    "plan_generator": RETAINED_TOOLS / "prepare_plan_0973_v2.py",
}
RETAINED_SHA256 = {
    "protocol": "56b509fde01bbd66702ac8cd835da0d09d74d332935926078c6545093ed52bf4",
    "runner": "f4f7af4ec8b8bb1818f3ce25dec046c8b8ad7c23d0a2611c76e24bf10b6aef5d",
    "reducer": "f33f14724f6264d87b2c3a0082c16bca711787140d06f3c7cfdd6e78b6f4c63b",
    "raw_stack": "3a8e36ebc5f21f4f0b8293cafa675d722f2264fb1cfd6c7f48271b4f0a5056d2",
    "plan_generator": "c3beab19960e0a7b4d73f252aa015c6b398508f2bf4e33cb6ace4b271129f65f",
}

class ProtocolError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_binding(path: str | Path) -> dict[str, Any]:
    # Preserve the custody path spelling from the retained plan.  Several
    # retained 0969 inputs are symlinks into the checkout; resolving them
    # would silently relabel the historical binding while the bytes stay the
    # same.
    resolved = Path(path).absolute()
    if not resolved.is_file():
        raise ProtocolError(f"binding path is not a file: {resolved}")
    return {"path": str(resolved), "sha256": sha256(resolved), "bytes": resolved.stat().st_size}


def load_json(path: str | Path) -> dict[str, Any]:
    resolved = Path(path).resolve()
    try:
        value = json.loads(resolved.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"cannot load JSON {resolved}: {exc}") from exc
    if not isinstance(value, dict):
        raise ProtocolError(f"JSON root is not an object: {resolved}")
    return value


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ProtocolError(f"cannot load retained module {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def verify_retained_files() -> None:
    for name, path in RETAINED_FILES.items():
        actual = sha256(path)
        if actual != RETAINED_SHA256[name]:
            raise ProtocolError(f"retained 0973 {name} changed: {actual} != {RETAINED_SHA256[name]}")
    if not RETAINED_HELPER_SEAL.is_file() or not RETAINED_PLAN.is_file():
        raise ProtocolError("retained 0973 seal/plan is missing")


verify_retained_files()
BASE = load_module("protocol_common_0973_v2", RETAINED_FILES["protocol"])
BASE_RAW_STACK = load_module("raw_stack_0973_v2", RETAINED_FILES["raw_stack"])
BASE_RUNNER = load_module("run_0973_v2", RETAINED_FILES["runner"])
BASE_REDUCER = load_module("reduce_0973_v2", RETAINED_FILES["reducer"])

PLAN_SCHEMA = "litchi.performance.0974.plan.v1"
GO_SCHEMA = "litchi.performance.0974.root_go.v1"
HELPER_SCHEMA = "litchi.performance.0974.helper_seal.v1"
AUTHOR_SCHEMA = "litchi.performance.0974.author_seal.v1"
MEMORY_TOKENS = ["-m", "16M"]
MEMORY_BYTES = 16 * 1024 * 1024
RAW_LOST_FIELD = "raw_lost_sum"
TERMINAL_LOST_FIELD = "lost_samples_summary"
TERMINAL_SCHEMA = "litchi.performance.0973.terminal.v2"
RAW_SCHEMA = BASE.RAW_STACK_SCHEMA


def binding_matches(row: Any, expected: dict[str, Any]) -> bool:
    return isinstance(row, dict) and row == expected


def retained_binding(name: str) -> dict[str, Any]:
    path = RETAINED_FILES[name]
    row = file_binding(path)
    if row["sha256"] != RETAINED_SHA256[name]:
        raise ProtocolError(f"retained helper drift: {name}")
    return row


def native_record_argv(plan: dict[str, Any], child_dir: Path) -> list[str]:
    """Return the sealed 0973 argv with exactly one 0974 insertion."""
    base_argv = list(BASE_RUNNER.child_record_argv(plan, child_dir))
    if base_argv.count("--call-graph") != 1:
        raise ProtocolError("retained native command has no unique --call-graph")
    if MEMORY_TOKENS != ["-m", "16M"]:
        raise ProtocolError("0974 memory delta changed")
    index = base_argv.index("--call-graph")
    result = base_argv[:index] + MEMORY_TOKENS + base_argv[index:]
    if result[index:index + 2] != MEMORY_TOKENS or result[index + 2] != "--call-graph":
        raise ProtocolError("0974 memory delta was not inserted before --call-graph")
    if result.count("-m") != 1 or result.count("16M") != 1:
        raise ProtocolError("0974 memory selector is duplicated")
    return result


_RAW_LOST = re.compile(r"PERF_RECORD_LOST:\s+.*?\blost:(\d+)\b")


def parse_raw_lost_sum(text: str) -> dict[str, Any]:
    """Account only PERF_RECORD_LOST values; do not add LOST_SAMPLES."""
    values = [int(value) for value in _RAW_LOST.findall(text)]
    return {
        RAW_LOST_FIELD: sum(values),
        "raw_lost_record_count": len(values),
        "raw_lost_scope": "sum of PERF_RECORD_LOST lost fields; PERF_RECORD_LOST_SAMPLES is not added",
    }


_TOTAL_LOST = re.compile(r"^#\s*Total Lost Samples:\s*(\d+)\s*$", re.MULTILINE)


def collect_terminal_lost_samples(perf_data: Path) -> dict[str, Any]:
    """Read perf's single terminal LOST_SAMPLES summary for one data file."""
    argv = [BASE.PERF_PATH, "report", "--stdio", "--percent-limit", "100", "-i", str(perf_data)]
    env = dict(BASE.ENVIRONMENT)
    cp = subprocess.run(argv, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE, text=True, env=env, timeout=300,
                        check=False)
    if cp.returncode != 0:
        raise ProtocolError(f"terminal perf report failed: {cp.stderr[-500:]}")
    matches = _TOTAL_LOST.findall(cp.stdout)
    if len(matches) != 1:
        raise ProtocolError("terminal LOST_SAMPLES summary is missing or duplicated")
    return {
        "record_type": "PERF_RECORD_LOST_SAMPLES",
        "lost_samples": int(matches[0]),
        "source": "perf report --stdio terminal summary",
        "argv": argv,
    }


def raw_stack_with_loss(*args: Any, **kwargs: Any) -> dict[str, Any]:
    result = BASE_RAW_STACK.scan_dump(*args, **kwargs)
    dump_path = Path(args[0]) if args else Path(kwargs["dump_path"])
    result.update(parse_raw_lost_sum(dump_path.read_text(encoding="utf-8", errors="strict")))
    return result


def require_zero_loss(raw_stack: dict[str, Any], terminal: dict[str, Any]) -> None:
    if raw_stack.get(RAW_LOST_FIELD) != 0:
        raise ProtocolError(f"raw PERF_RECORD_LOST sum is nonzero: {raw_stack.get(RAW_LOST_FIELD)!r}")
    summary = terminal.get(TERMINAL_LOST_FIELD)
    if not isinstance(summary, dict) or summary.get("record_type") != "PERF_RECORD_LOST_SAMPLES":
        raise ProtocolError("terminal LOST_SAMPLES summary is missing")
    if summary.get("lost_samples") != 0:
        raise ProtocolError(f"terminal LOST_SAMPLES summary is nonzero: {summary.get('lost_samples')!r}")


def _require_binding(row: Any, name: str) -> None:
    if not isinstance(row, dict):
        raise ProtocolError(f"missing plan binding: {name}")
    path = row.get("path")
    if not isinstance(path, str) or file_binding(path) != {key: row[key] for key in ("path", "sha256", "bytes") if key in row}:
        # Preserve optional metadata (for example version/build ID), but the
        # path/hash/size tuple must be the current file tuple.
        actual = file_binding(path) if isinstance(path, str) and Path(path).is_file() else None
        if actual is None or any(row.get(key) != actual[key] for key in actual):
            raise ProtocolError(f"plan binding changed: {name}")


def _check_reused_control(row: dict[str, Any], path: Path, name: str) -> None:
    expected = file_binding(path)
    for key in ("path", "sha256", "bytes"):
        if row.get(key) != expected[key]:
            raise ProtocolError(f"retained {name} binding drift")


def _check_helper_seal(plan: dict[str, Any]) -> None:
    seal_row = plan.get("helper_seal")
    if not isinstance(seal_row, dict):
        raise ProtocolError("missing 0974 helper seal")
    _require_binding(seal_row, "helper_seal")
    seal = load_json(seal_row["path"])
    if seal.get("schema") != HELPER_SCHEMA or seal.get("status") != "sealed":
        raise ProtocolError("0974 helper seal is not sealed")
    if plan.get("helpers_seal_sha256") != seal_row.get("sha256"):
        raise ProtocolError("0974 helper seal checksum drift")
    helpers = plan.get("helpers")
    if not isinstance(helpers, dict) or set(helpers) != {"protocol", "runner", "reducer", "plan_generator", "raw_stack", "shim_source", "shim_map", "shim_so"}:
        raise ProtocolError("0974 helper set drift")
    if seal.get("helpers") != helpers:
        raise ProtocolError("0974 helper seal/plan mismatch")
    for name, row in helpers.items():
        _require_binding(row, f"helper {name}")
    retained = seal.get("retained_0973")
    if not isinstance(retained, dict):
        raise ProtocolError("0973 retained helper custody is missing")
    for name in RETAINED_FILES:
        expected = retained_binding(name)
        if retained.get(name) != expected:
            raise ProtocolError(f"retained 0973 helper custody drift: {name}")
    old_seal = seal.get("retained_0973_helper_seal")
    if old_seal != file_binding(RETAINED_HELPER_SEAL):
        raise ProtocolError("retained 0973 helper seal custody drift")


def check_plan_bindings(plan: dict[str, Any], plan_path: Path | None = None) -> None:
    """Validate the fresh plan while leaving 0973's decoder contracts intact."""
    if plan.get("schema") != PLAN_SCHEMA or plan.get("status") != "pre_capture":
        raise ProtocolError("0974 plan schema/status mismatch")
    if plan.get("no_build_or_workload") is not True:
        raise ProtocolError("0974 plan permits build/workload before GO")
    if plan.get("execution") != {"serial": True, "retry": False, "timeout_seconds": 900, "children": BASE.CHILDREN}:
        raise ProtocolError("0974 execution contract drift")
    matrix = plan.get("matrix") or {}
    if matrix != {"children": BASE.CHILDREN, "warmups": BASE.WARMUPS, "samples": BASE.SAMPLES, "cpu": BASE.CPU, "cells": [{"case": "docfloat", "edits": 1, "retention": "default"}], "serial_order": list(range(1, BASE.CHILDREN + 1))}:
        raise ProtocolError("0974 matrix drift")
    perf = plan.get("perf") or {}
    if perf.get("path") != BASE.PERF_PATH or perf.get("sha256") != BASE.PERF_SHA or perf.get("event") != "instructions:u" or perf.get("period") != BASE.PERIOD or perf.get("callgraph") != BASE.CALLGRAPH or perf.get("stack_cap") != BASE.STACK_CAP:
        raise ProtocolError("0974 perf contract drift")
    if perf.get("record_memory") != {"tokens": MEMORY_TOKENS, "bytes": MEMORY_BYTES, "position": "immediately before --call-graph"}:
        raise ProtocolError("0974 perf memory selector drift")
    if perf.get("sample_flags") != ["--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--user-callchains"]:
        raise ProtocolError("0974 perf sample flags drift")
    marker = plan.get("marker") or {}
    expected_marker = {"symbol": BASE.MARKER, "dso_path": BASE.BINARY_PATH, "build_id": BASE.BINARY_BUILD_ID, "calls": BASE.CALLS, "p50_scope": "80 measured samples only", "marker_scope": "5 warmups + 80 measured calls combined", "owner_scope": "all 12 children; missing owner membership unknown"}
    if marker != expected_marker:
        raise ProtocolError("0974 marker scope drift")
    expected_gate = {"all_children": True, "min_marker_samples": 100, "lost_events": 0, "marker_period_share_positive": True, "on_failure": "stop_without_retry"}
    expected_pilot = {"child": 1, "min_marker_samples": 100, "lost_events": 0, "marker_period_share_positive": True, "on_failure": "stop_without_retry"}
    if plan.get("pilot_gate") != expected_pilot or plan.get("child_gate") != expected_gate:
        raise ProtocolError("0974 marker gate drift")
    if plan.get("environment") != BASE.ENVIRONMENT:
        raise ProtocolError("0974 environment drift")
    if plan.get("native_environment") != {"LD_PRELOAD": "absent", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}:
        raise ProtocolError("0974 native environment drift")
    if plan.get("command_receipt_contract") != {"environment": {"LD_PRELOAD": "absent", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}, "cwd": "record inherited runner cwd before launch", "no_secret_environment": True}:
        raise ProtocolError("0974 command receipt contract drift")
    required = ("binary", "fixture", "source_manifest", "source_file_manifest", "source_proof", "proposal", "design", "metric_proposal", "adr_revalidation", "retained_adr_revalidation", "validator", "launcher", "python", "taskset", "perf", "shim_source", "shim_version_script", "helper_seal", "review_context", "review_findings", "libdw", "libelf", "compiler", "header", "aslr_control", "synthetic_e2e", "capability_probe")
    for key in required:
        _require_binding(plan.get(key), key)
    binary = plan["binary"]
    expected_binary = {**file_binding(BASE.BINARY_PATH), "build_id": BASE.BINARY_BUILD_ID}
    if binary != expected_binary:
        raise ProtocolError("0974 binary binding drift")
    if plan["fixture"] != file_binding(BASE.FIXTURE_PATH):
        raise ProtocolError("0974 fixture binding drift")
    if plan["source_manifest"] != file_binding(BASE.SOURCE_MANIFEST_PATH) or plan["source_file_manifest"] != file_binding(BASE.SOURCE_FILE_MANIFEST_PATH) or plan["source_proof"].get("sha256") != BASE.SOURCE_PROOF_SHA:
        raise ProtocolError("0974 source custody drift")
    if plan["adr_revalidation"].get("path") != str((ROOT / "receipts/adr-goal-revalidation-0974-v1.json").resolve()):
        raise ProtocolError("0974 ADR receipt path drift")
    if plan["retained_adr_revalidation"].get("path") != "/home/zhuhe/litchi-performance-0973/receipts/adr-goal-revalidation-0973-v1.json":
        raise ProtocolError("retained ADR receipt path drift")
    adr = load_json(plan["adr_revalidation"]["path"])
    if adr.get("status") != "pass" or adr.get("accepted_count") != 32 or len(adr.get("adr_entries") or []) != 32 or adr.get("goal", {}).get("path") != "/home/zhuhe/code/litchi/docs/GOAL.md":
        raise ProtocolError("0974 ADR/GOAL receipt is not a passing 32-entry revalidation")
    for key, path in (("metric_proposal", BASE.METRIC_PROPOSAL_PATH), ("review_findings", BASE.REVIEW_FINDINGS_PATH), ("aslr_control", BASE.ASLR_CONTROL_PATH), ("synthetic_e2e", BASE.SYNTHETIC_E2E_PATH)):
        _check_reused_control(plan[key], Path(path), key)
    _check_reused_control(plan["design"], ROOT / "design-v1.md", "0974 design")
    aslr = load_json(plan["aslr_control"]["path"])
    if aslr.get("status") != "pass" or aslr.get("source_changed") is not False:
        raise ProtocolError("retained ASLR control is not a passing unchanged-source control")
    e2e = load_json(plan["synthetic_e2e"]["path"])
    if e2e.get("status") != "pass" or e2e.get("workload_started") is not False or e2e.get("build_started") is not False:
        raise ProtocolError("retained E2E control is not an offline pass")
    if plan["review_context"].get("path") != "/home/zhuhe/litchi-performance-0972/receipts/dwfl-replay-author-seal-v6.json":
        raise ProtocolError("retained review context drift")
    capability = load_json(plan["capability_probe"]["path"])
    if capability.get("status") != "setup_failed_no_retry" or capability.get("failure", {}).get("returncode") != 255 or capability.get("failure", {}).get("workload_executed") is not False:
        raise ProtocolError("0974 -m 16M capability result is not the recorded setup failure")
    if capability.get("command", {}).get("argv", [])[0:6] != [BASE.TASKSET_PATH, "--cpu-list", str(BASE.CPU), BASE.PERF_PATH, "record", "-e"]:
        raise ProtocolError("0974 capability probe command drift")
    if plan["proposal"].get("path") != "/home/zhuhe/litchi-performance-0973/diagnostic-proposal-v1.md":
        raise ProtocolError("retained proposal drift")
    if plan["validator"].get("path") != BASE.VALIDATOR_PATH or plan["launcher"].get("path") != BASE.LAUNCHER_PATH or plan["python"].get("path") != BASE.PYTHON_PATH or plan["taskset"].get("path") != BASE.TASKSET_PATH:
        raise ProtocolError("0974 executable/helper path drift")
    if plan["libdw"].get("path") != BASE.LIBDW_PATH or plan["libelf"].get("path") != BASE.LIBELF_PATH or plan["compiler"].get("path") != BASE.COMPILER_PATH or plan["header"].get("path") != BASE.HEADER_PATH:
        raise ProtocolError("0974 decoder dependency path drift")
    if plan.get("report_contract") != {"case": "docfloat", "edits": 1, "retention": "default_8mib", "warmups": BASE.WARMUPS, "samples_requested": BASE.SAMPLES, "sample_indices": list(range(BASE.SAMPLES)), "elapsed_ns_identity": "samples[i].elapsed_ns == elapsed_ns[i]"}:
        raise ProtocolError("0974 report contract drift")
    expected_raw = {"schema": BASE.RAW_STACK_SCHEMA, "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER", "period": BASE.PERIOD, "stack_cap": BASE.STACK_CAP, "bounds": "0 <= dyn_size <= captured_stack_size <= 32768", "missing_or_short": "missing fields reject; bounded short captured_stack_size is retained in its distribution", "raw_scope": "all PERF_RECORD_SAMPLE records without misc filtering", "raw_lost_field": RAW_LOST_FIELD, "raw_lost_scope": "sum of PERF_RECORD_LOST lost fields; PERF_RECORD_LOST_SAMPLES is not added"}
    if plan.get("raw_stack_contract") != expected_raw:
        raise ProtocolError("0974 raw STACK_USER/loss contract drift")
    if plan.get("terminal_contract") != {"field": TERMINAL_LOST_FIELD, "record_type": "PERF_RECORD_LOST_SAMPLES", "required_value": 0, "source": "perf report --stdio terminal summary", "raw_loss_not_added": True}:
        raise ProtocolError("0974 terminal loss contract drift")
    if plan.get("raw_artifacts") != BASE.RAW_ARTIFACTS:
        raise ProtocolError("0974 raw artifact contract drift")
    if plan.get("process_policy") != {"require_sample_pid_tid_equal_exec_pid": True, "reject_any_other_thread_or_descendant": True, "sampling_monitor": False}:
        raise ProtocolError("0974 process policy drift")
    if plan.get("outputs", {}).get("capture_root") != str((ROOT / "capture/instruction-profile-v2").resolve()):
        raise ProtocolError("0974 output custody drift")
    _check_helper_seal(plan)
    if plan_path is not None and plan.get("plan_path") != str(plan_path.resolve()):
        raise ProtocolError("0974 plan self path mismatch")


def load_plan(path: Path) -> dict[str, Any]:
    plan = load_json(path)
    check_plan_bindings(plan, path)
    return plan


def check_go(go: dict[str, Any], plan: dict[str, Any], plan_path: Path, go_path: Path) -> None:
    if go.get("schema") != GO_SCHEMA or go.get("authorized") is not True or go.get("status") != "authorized":
        raise ProtocolError("0974 root GO is not authorized")
    row = go.get("plan") or {}
    expected = file_binding(plan_path)
    if row != expected:
        raise ProtocolError("0974 GO plan binding mismatch")
    if (go.get("quiescence") or {}).get("status") != "pass":
        raise ProtocolError("0974 GO quiescence missing")
    if go.get("helper_seal") != plan.get("helper_seal"):
        raise ProtocolError("0974 GO helper seal mismatch")
    if go.get("binary") != plan.get("binary") or go.get("fixture") != plan.get("fixture"):
        raise ProtocolError("0974 GO input binding mismatch")
    if go.get("record_memory") != plan.get("perf", {}).get("record_memory"):
        raise ProtocolError("0974 GO memory selector mismatch")
    for key in ("libdw", "libelf", "compiler", "header", "shim_source", "shim_version_script"):
        if go.get(key) != plan.get(key):
            raise ProtocolError(f"0974 GO dependency binding mismatch: {key}")
    _check_helper_seal(plan)


def _terminal_for_path(path: Path) -> Path:
    return path.with_name(path.name.replace(".packet.json", ".terminal.json"))


def install_runner_patches() -> None:
    """Patch only the retained runner's command/loss boundaries."""
    original_write = BASE_RUNNER.p.write_json_exclusive
    original_scan = BASE_RUNNER.raw_stack.scan_dump
    terminal_cache: dict[Path, dict[str, Any]] = {}

    def write(path: Path, value: Any) -> None:
        resolved = Path(path).resolve()
        payload = value
        if resolved.name.endswith(".command.json"):
            payload = dict(value)
            # Keep the command receipt's environment intentionally small and
            # record the inherited cwd separately; no secret process
            # environment is serialized.
            payload["cwd"] = str(Path.cwd())
            payload["cwd_policy"] = "inherited runner cwd; recorded before launch"
        elif resolved.name.endswith(".terminal.json"):
            payload = dict(value)
            data = resolved.with_name(resolved.name.replace(".terminal.json", ".perf.data"))
            if data.is_file() and payload.get("returncode") == 0:
                payload[TERMINAL_LOST_FIELD] = collect_terminal_lost_samples(data)
            else:
                payload[TERMINAL_LOST_FIELD] = {
                    "record_type": "PERF_RECORD_LOST_SAMPLES",
                    "lost_samples": None,
                    "source": "perf report --stdio terminal summary unavailable",
                    "argv": [BASE.PERF_PATH, "report", "--stdio", "--percent-limit", "100", "-i", str(data)],
                }
            terminal_cache[resolved] = payload
        elif resolved.name.endswith(".packet.json"):
            payload = dict(value)
            terminal_path = _terminal_for_path(resolved)
            if terminal_path in terminal_cache:
                payload["terminal"] = terminal_cache[terminal_path]
            if payload.get("status") == "complete":
                artifact_row = (payload.get("artifacts") or {}).get("raw_stack")
                if not isinstance(artifact_row, dict) or not isinstance(artifact_row.get("path"), str):
                    raise ProtocolError("complete packet has no raw STACK_USER artifact")
                raw_summary = load_json(artifact_row["path"])
                require_zero_loss(raw_summary, payload.get("terminal") or {})
        original_write(resolved, payload)

    def scan(*args: Any, **kwargs: Any) -> dict[str, Any]:
        result = original_scan(*args, **kwargs)
        dump_path = Path(args[0]) if args else Path(kwargs["dump_path"])
        result.update(parse_raw_lost_sum(dump_path.read_text(encoding="utf-8", errors="strict")))
        return result

    BASE_RUNNER.p.write_json_exclusive = write
    BASE_RUNNER.raw_stack.scan_dump = scan
    BASE_RUNNER.child_record_argv = native_record_argv


def self_test() -> dict[str, Any]:
    plan_stub = {
        "binary": {"path": BASE.BINARY_PATH},
        "fixture": {"path": BASE.FIXTURE_PATH},
    }
    command = native_record_argv(plan_stub, ROOT / "capture" / "child-01")
    index = command.index("--call-graph")
    if command[index - 2:index] != MEMORY_TOKENS:
        raise ProtocolError("self-test memory insertion failed")
    if command[index + 1] != BASE.CALLGRAPH:
        raise ProtocolError("self-test callgraph changed")
    zero_raw = parse_raw_lost_sum("0 0 [0x20]: PERF_RECORD_SAMPLE\n")
    if zero_raw[RAW_LOST_FIELD] != 0:
        raise ProtocolError("zero raw loss self-test failed")
    nonzero_raw = parse_raw_lost_sum("31 0 [0x30]: PERF_RECORD_LOST: id:1: lost:4\n")
    if nonzero_raw[RAW_LOST_FIELD] != 4:
        raise ProtocolError("raw loss mutation self-test failed")
    zero_terminal = {
        TERMINAL_LOST_FIELD: {"record_type": "PERF_RECORD_LOST_SAMPLES", "lost_samples": 0}
    }
    require_zero_loss(zero_raw, zero_terminal)
    try:
        require_zero_loss(nonzero_raw, zero_terminal)
    except ProtocolError:
        pass
    else:
        raise ProtocolError("raw nonzero loss was accepted")
    try:
        require_zero_loss(zero_raw, {TERMINAL_LOST_FIELD: {"record_type": "PERF_RECORD_LOST_SAMPLES", "lost_samples": 1}})
    except ProtocolError:
        pass
    else:
        raise ProtocolError("terminal nonzero loss was accepted")
    return {
        "schema": "litchi.performance.0974.protocol_adapter_self_test.v1",
        "status": "pass",
        "workload_started": False,
        "retained_0973": {name: retained_binding(name) for name in RETAINED_FILES},
        "record_memory_tokens": MEMORY_TOKENS,
        "raw_loss_field": RAW_LOST_FIELD,
        "terminal_loss_field": TERMINAL_LOST_FIELD,
    }


if __name__ == "__main__":
    print(json.dumps(self_test(), indent=2, sort_keys=True))
