#!/usr/bin/env python3
"""Create a fresh, immutable 0978 helper seal and pre-capture plan.

The command only hashes files and writes new JSON with exclusive creation. It
never builds, launches perf, starts Office, or mutates a retained 0977
receipt. The decoder, reducer, and protocol tests are arguments so their
final bytes can be reviewed and bound before a root GO is issued.
"""
from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import plan_binding_0978_v1 as b


ROOT = b.ROOT
TOOLS = ROOT / "tools"
DESIGN = b.PROTOCOL_DESIGN
ADR_RECEIPT = ROOT / "receipts/adr-goal-revalidation-0978-v1.json"
POLICY = b.POLICY_REVIEW


def binding(path: Path) -> dict[str, Any]:
    return b.binding(path)


def dependency(path: Path, *, scope: str | None = None) -> dict[str, Any]:
    row = binding(path)
    if scope is not None:
        row["scope"] = scope
    return row


def load(path: Path) -> dict[str, Any]:
    value = b.strict_load(path)
    if not isinstance(value, dict):
        raise b.BindingError(f"{path}: expected object")
    return value


def _paths(values: Iterable[Path]) -> tuple[Path, ...]:
    rows = tuple(Path(value).resolve() for value in values)
    if not rows:
        raise b.BindingError("at least one protocol test helper is required")
    return rows


def make_helpers(
    decoder: Path = b.FOCUSED_DECODER,
    reducer: Path = b.REDUCER,
    *,
    scanner: Path = b.SCANNER,
    lifecycle: Path = b.LIFECYCLE,
    tests: Iterable[Path] = (b.SCANNER_TESTS, b.LIFECYCLE_TESTS, b.REDUCER_TESTS),
    policy: Path = b.POLICY_REVIEW,
) -> dict[str, Any]:
    return {
        # This is a retained sealed native handoff helper. Binding it here
        # preserves the exact argv/bridge boundary without copying it.
        "bridge": binding(b.BRIDGE),
        "runner": binding(TOOLS / "run_0978_v1.py"),
        "plan_binding": binding(TOOLS / "plan_binding_0978_v1.py"),
        "plan_generator": binding(TOOLS / "prepare_plan_0978_v1.py"),
        "launcher": binding(b.LAUNCHER),
        "decoder": binding(decoder),
        "scanner": binding(scanner),
        "lifecycle": binding(lifecycle),
        "reducer": binding(reducer),
        "tests": [binding(path) for path in _paths(tests)],
        "policy": binding(policy),
    }


def make_seal(
    decoder: Path = b.FOCUSED_DECODER,
    reducer: Path = b.REDUCER,
    *,
    scanner: Path = b.SCANNER,
    lifecycle: Path = b.LIFECYCLE,
    tests: Iterable[Path] = (b.SCANNER_TESTS, b.LIFECYCLE_TESTS, b.REDUCER_TESTS),
    policy: Path = b.POLICY_REVIEW,
) -> dict[str, Any]:
    helpers = make_helpers(decoder, reducer, scanner=scanner, lifecycle=lifecycle, tests=tests, policy=policy)
    return {
        "schema": b.HELPER_SCHEMA,
        "status": "sealed",
        "version": "0978-v1",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "scope": "fresh variable-layout scanner/lifecycle ledger/adapter runner and plan; no workload started",
        "helpers": helpers,
        "retained_lineage": {
            "base_0977_runner": dependency(Path("/home/zhuhe/litchi-performance-0977/tools/run_0977_v2.py"), scope="immutable reference; native launch contract retained"),
            "base_0977_plan_binding": dependency(Path("/home/zhuhe/litchi-performance-0977/tools/plan_binding_0977_v2.py"), scope="immutable reference; custody checks retained"),
            "base_0977_reducer": dependency(b.RETAINED_REDUCER_0977, scope="immutable delegated custody/oracle implementation; exact SHA required"),
            "base_0973_reducer": dependency(b.RETAINED_REDUCER_0973, scope="loaded by retained 0977 reducer; immutable arithmetic implementation"),
            "base_0973_protocol_common": dependency(b.RETAINED_PROTOCOL_0973, scope="loaded by retained 0977 reducer; immutable protocol helper"),
            "base_0971_protocol_common": dependency(b.RETAINED_PROTOCOL_0971, scope="loaded by retained 0973 protocol helper; immutable dependency"),
            "base_0969_protocol_common": dependency(b.RETAINED_PROTOCOL_0969, scope="loaded by retained 0971 protocol helper; immutable dependency"),
            "base_0978_reducer_v1": dependency(b.RETAINED_REDUCER_0978_V1, scope="prior frozen 0978 reducer; superseded by versioned v2 helper and retained byte-for-byte"),
            "base_0977_helper_seal": dependency(Path("/home/zhuhe/litchi-performance-0977/receipts/helper-seal-0977-v2.json"), scope="immutable sealed helper lineage"),
            "affinity_bridge_0977": dependency(b.BRIDGE, scope="retained byte-for-byte; exact native handoff and command boundary"),
            "bridge_lineage_0976": dependency(b.BRIDGE_LINEAGE, scope="historical handoff lineage"),
        },
        "policy": binding(policy),
        "workload_started": False,
        "build_started": False,
    }


def make_plan(
    seal_path: Path,
    plan_path: Path,
    *,
    decoder: Path = b.FOCUSED_DECODER,
    reducer: Path = b.REDUCER,
    scanner: Path = b.SCANNER,
    lifecycle: Path = b.LIFECYCLE,
    tests: Iterable[Path] = (b.SCANNER_TESTS, b.LIFECYCLE_TESTS, b.REDUCER_TESTS),
    policy: Path = b.POLICY_REVIEW,
) -> dict[str, Any]:
    tests = _paths(tests)
    seal = load(seal_path)
    helpers = make_helpers(decoder, reducer, scanner=scanner, lifecycle=lifecycle, tests=tests, policy=policy)
    if seal.get("helpers") != helpers:
        raise b.BindingError("helper seal does not match current helper files")
    dependencies = {
        "perf": binding(b.PERF),
        "taskset": binding(b.TASKSET),
        "python": binding(b.PYTHON),
        "launcher": binding(b.LAUNCHER),
        "bridge": dependency(b.BRIDGE, scope="retained 0977 bridge; exact native handoff and command boundary"),
        "bridge_lineage": dependency(b.BRIDGE_LINEAGE, scope="handoff lineage; 0977 bridge retained byte-for-byte"),
        "decoder_so": binding(b.DECODER_SO),
        "decoder_source": binding(b.DECODER_SOURCE),
        "decoder_map": binding(b.DECODER_MAP),
        "decoder_validator": binding(b.DECODER_VALIDATOR),
        "libdw": binding(b.LIBDW),
        "libelf": binding(b.LIBELF),
        "decoder_header": binding(b.HEADER),
        "focused_decode": binding(decoder),
        "control_target": binding(b.CONTROL_TARGET),
        "protocol_design": binding(DESIGN),
        "policy_review": binding(policy),
        "scanner": binding(scanner),
        "lifecycle": binding(lifecycle),
        "reducer": binding(reducer),
        "scanner_tests": binding(b.SCANNER_TESTS),
        "lifecycle_tests": binding(b.LIFECYCLE_TESTS),
        "reducer_tests": binding(b.REDUCER_TESTS),
        "replay_tool": binding(b.REPLAY_TOOL),
        "replay_receipt": binding(b.REPLAY_RECEIPT),
        "scanner_controls": binding(b.SCANNER_CONTROLS),
        "raw_test_review": binding(b.RAW_TEST_REVIEW),
        "control_command_receipt": binding(b.CONTROL_COMMAND_RECEIPT),
        "control_test": binding(b.CONTROL_TEST),
        "control_audit": binding(b.CONTROL_AUDIT),
        "control_equivalence": binding(b.CONTROL_EQUIVALENCE),
        "retained_reducer_0977": dependency(b.RETAINED_REDUCER_0977, scope="retained reducer delegated by fresh 0978 reducer; exact immutable SHA"),
        "retained_reducer_0973": dependency(b.RETAINED_REDUCER_0973, scope="loaded by retained 0977 reducer; exact immutable closure"),
        "retained_protocol_0973": dependency(b.RETAINED_PROTOCOL_0973, scope="loaded by retained 0977 reducer; exact immutable closure"),
        "retained_protocol_0971": dependency(b.RETAINED_PROTOCOL_0971, scope="loaded by retained 0973 protocol helper; exact immutable closure"),
        "retained_protocol_0969": dependency(b.RETAINED_PROTOCOL_0969, scope="loaded by retained 0971 protocol helper; exact immutable closure"),
        "retained_reducer_0978_v1": dependency(b.RETAINED_REDUCER_0978_V1, scope="prior frozen 0978 reducer; exact immutable lineage snapshot"),
    }
    plan = {
        "schema": b.PLAN_SCHEMA,
        "status": "pre_capture",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "plan_path": str(plan_path.resolve()),
        "design": binding(DESIGN),
        "policy_review": binding(policy),
        "adr_goal_revalidation": binding(ADR_RECEIPT),
        "capability_probe": binding(b.CAPABILITY),
        "perf_requalification": binding(b.PERF_REQUALIFICATION_SOURCE),
        "perf_requalification_posthoc": binding(b.PERF_REQUALIFICATION),
        "perf_tool_drift": binding(b.PERF_TOOL_DRIFT),
        "source_manifest": binding(b.SOURCE_MANIFEST),
        "source_file_manifest": binding(b.SOURCE_FILE_MANIFEST),
        "binary": binding(b.BINARY),
        "fixture": binding(b.FIXTURE),
        "python": binding(b.PYTHON),
        "taskset": binding(b.TASKSET),
        "launcher": binding(b.LAUNCHER),
        "dependencies": dependencies,
        "helpers": helpers,
        # The decoder validates this direct binding at entry in addition to
        # the helper-seal ``helpers.decoder`` row.
        "focused_decode": binding(decoder),
        "helper_seal": binding(seal_path),
        "helper_seal_sha256": b.sha256(seal_path),
        "decoder_reviews": {
            "final_v2_review": binding(b.DECODER_REVIEW),
            "noop_posthoc": binding(b.NOOP_REVIEW),
            "focused_controls": binding(b.FOCUSED_CONTROLS),
        },
        "protocol": {
            "design": binding(DESIGN),
            "policy_review": binding(policy),
            "scanner": binding(scanner),
            "lifecycle": binding(lifecycle),
            "tests": [binding(path) for path in tests],
            "replay_tool": binding(b.REPLAY_TOOL),
            "replay_receipt": binding(b.REPLAY_RECEIPT),
            "scanner_controls": binding(b.SCANNER_CONTROLS),
            "raw_test_review": binding(b.RAW_TEST_REVIEW),
            "control_command_receipt": binding(b.CONTROL_COMMAND_RECEIPT),
            "control_test": binding(b.CONTROL_TEST),
            "control_audit": binding(b.CONTROL_AUDIT),
            "control_equivalence": binding(b.CONTROL_EQUIVALENCE),
            "variable_layout": True,
            "lifecycle_records_are_authoritative": True,
            "valid_empty_is_unknown": True,
            "unknown_period_cutoff": None,
            "malformed_or_unconserved_fails_closed": True,
        },
        "batch": {
            "children": b.CHILDREN, "serial": True, "cpu": b.CPU, "recorder_cpu": b.RECORDER_CPU,
            "period": b.PERIOD, "stack_cap": b.STACK_CAP, "warmups": b.WARMUPS, "samples": b.SAMPLES,
            "minimum_marker_samples": 100, "marker_period_positive": True, "raw_lost_sum": 0,
            "terminal_lost_samples": 0, "structural_conservation_required": True,
            "valid_empty_is_unknown": True, "unknown_period_cutoff": None,
            "stop_on_first_failure": True, "retry": False,
        },
        "execution": {
            "children": b.CHILDREN, "serial": True, "retry": False,
            "timeout_seconds": b.TIMEOUT_SECONDS, "stop_on_first_failure": True,
            "quiescence_required": True,
        },
        "perf": {
            "path": str(b.PERF.resolve()), "sha256": b.sha256(b.PERF), "event": "instructions:u", "period": b.PERIOD,
            "callgraph": "dwarf,32768", "stack_cap": b.STACK_CAP, "record_mode": ["--per-thread"],
            "record_memory": ["-m", "16M"],
            "sample_flags": ["--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--user-callchains"],
        },
        "perf_contract": b.PERF_CONTRACT,
        "environment": b.environment(),
        "workload": {
            "case": "docfloat", "edits": 1, "retention": "default",
            "target_argv_template": ["<binary>", "--case", "docfloat", "--input", "<fixture>", "--edits", "1", "--retention", "default", "--warmup", "5", "--samples", "80"],
            "same_binary_fixture_all_children": True,
            "same_pid_exec_chain": ["perf", "affinity_bridge_0977.py", "launch_baseline_v2.py", "bound_workload"],
        },
        "control": {
            "target_script": binding(b.CONTROL_TARGET),
            "seconds": 0.25,
            "check_interval": 0.02,
            "output_root": str((ROOT / "capture/affinity-control-v1").resolve()),
            "scope": "retained 0977 harmless stdlib control; exact bridge/Popen boundary; no Office marker decoder",
            "raw_gates": {"raw_lost_sum": 0, "terminal_lost_samples": 0, "cpu": b.CPU, "single_pid": True},
            "reuse_boundary": "exact 0977 command/affinity boundary; independently command-self-tested before GO",
            "equivalence_receipt": binding(b.CONTROL_EQUIVALENCE),
        },
        "marker": {
            "build_id": "e5c457384eeebd057932c444efda77618f137399",
            "calls": 85, "dso_path": str(b.BINARY.resolve()),
            "marker_scope": "5 warmups + 80 measured calls combined",
            "owner_scope": "all 12 children; missing owner membership unknown",
            "p50_scope": "80 measured samples only",
            "symbol": "doc_retention_probe_0942_cfb_profile::profiled_measured_case",
        },
        "pilot_gate": {"child": 1, "lost_events": 0, "marker_period_share_positive": True, "min_marker_samples": 100, "on_failure": "stop_without_retry"},
        "child_gate": {
            "all_children": True, "raw_lost_sum": 0, "terminal_lost_samples": 0,
            "min_marker_samples": 100, "marker_period_share_positive": True,
            "structural_conservation": True, "valid_empty_is_unknown": True,
            "unknown_period_cutoff": None, "on_failure": "stop_without_retry",
        },
        "report_contract": {
            "case": "docfloat", "edits": 1, "retention": "default_8mib", "warmups": 5,
            "samples_requested": 80, "sample_indices": list(range(80)),
            "elapsed_ns_identity": "samples[i].elapsed_ns == elapsed_ns[i]",
            "unknown_period_reported": True,
        },
        "raw_stack_contract": {
            "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
            "period": b.PERIOD, "stack_cap": b.STACK_CAP,
            "bounds": "0 <= dyn_size <= captured_stack_size <= 32768",
            "raw_lost_field": "raw_lost_sum", "terminal_lost_field": "terminal_lost_samples",
            "raw_scope": "all PERF_RECORD_SAMPLE records without misc filtering",
            "task_stream_authority": "zero-loss raw task stream is lifetime authority",
            "variable_layout": True, "valid_zero_payload_is_not_truncation": True,
        },
        "terminal_contract": {
            "field": "PERF_RECORD_LOST_SAMPLES", "required_value": 0,
            "source": "decoder-owned offline perf report --stdio",
            "runner_precomputes": False, "raw_loss_not_added": True,
        },
        "task_stream_contract": {
            "header_required": ["task = 1", "comm = 1", "comm_exec = 1", "enable_on_exec = 1", "sample_stack_user = 32768", "read_format = ID|LOST"],
            "inheritance": "inherit absent or zero; inherit=1 rejects",
            "thread_map": "exactly one expected target PID/TID",
            "COMM": "perf-exec, python3.14, launcher, and bound binary may share one PID/TID",
            "FORK": "none", "CLONE": "none", "EXIT": "main PID/TID only",
            "sample_gate": "every admitted sample has expected PID/TID and CPU31",
            "raw_authority": "raw zero-loss task stream; bounded /proc observations are descriptive",
        },
        "bridge_handoff": {
            "target_cpu": b.CPU, "recorder_cpu": b.RECORDER_CPU, "parent_all_tids": True,
            "same_pid_exec": True, "startup_handoff_samples_retained": True,
            "target_side_stdlib_checks": False,
            "target_identity": "0970 launcher receipt and raw COMM/THREAD_MAP/EXIT records",
            "bridge_helper": str(b.BRIDGE.resolve()),
        },
        "decoder_interface": {
            "function": "decode_child(plan_path, child_dir)",
            "returns": "complete packet with direct lifecycle/attribution ledgers and structural gate fields; canonical full raw-stack, lifecycle, attribution, offline stdout/stderr, and command-receipt sidecars are mandatory and hash-bound; compact ABI summary is optional diagnostic data; raises DecodeFailure on malformed or unbound input",
            "offline_only": True, "native_ld_preload": "absent",
            "scanner_function": "scan_child(plan, child_dir, base_packet)",
            "filenames": list(b.DECODER_CANONICAL_FILENAMES + ("<child>.target-layout.json", "<child>.decoder.json") + b.DECODER_OPTIONAL_FILENAMES),
            "native_runner_filenames": ["<child>.perf.data", "<child>.report.json", "<child>.launch.json", "<child>.command.json", "<child>.terminal.json", "<child>.perf.stderr", "<child>.bridge.json", "<child>.affinity-observations.json"],
            "artifact_bindings": list(b.DECODER_CANONICAL_ARTIFACTS),
            "decoder_terminal_report": "decoder invokes perf report --stdio offline and binds the exact terminal LOST_SAMPLES text; runner does not precompute it",
            "decoder_owned_receipts": [
                "<child>.perf-header.command.json", "<child>.terminal-report.command.json",
                "<child>.perf-evlist-v.command.json", "<child>.perf-mmap-dump.command.json",
                "<child>.corrected.command.json", "<child>.terminal-summary.json", "<child>.loss.json",
            ],
            "modes": ["corrected_final_v2_plus_0978_lifecycle_abi"],
        },
        "outputs": {"capture_root": str((ROOT / "capture/instruction-profile-v1").resolve()), "analysis_root": str((ROOT / "analysis").resolve())},
        "native_environment": {"LC_ALL": "C", "LANG": "C", "TZ": "UTC", "LD_PRELOAD": "absent"},
        "command_receipt_contract": {"cwd": str(b.REPO.resolve()), "no_secret_environment": True, "LD_PRELOAD": "absent"},
        "interpretation": {
            "fresh_cohort": True, "not_a_retry_of": "/home/zhuhe/litchi-performance-0977",
            "diagnostic_only_until_reduction": True, "no_workload_started": True,
            "valid_empty_samples": "retained and classified as explicit unknown",
            "unknown_period": "descriptive/bounded; no cutoff and no optimization admission",
        },
        "no_build_or_workload": True,
    }
    return plan


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--seal-output", type=Path)
    ap.add_argument("--helper-seal", type=Path)
    ap.add_argument("--decoder", type=Path, default=b.FOCUSED_DECODER)
    ap.add_argument("--reducer", type=Path, default=b.REDUCER)
    ap.add_argument("--scanner", type=Path, default=b.SCANNER)
    ap.add_argument("--lifecycle", type=Path, default=b.LIFECYCLE)
    ap.add_argument("--policy", type=Path, default=b.POLICY_REVIEW)
    ap.add_argument("--test", dest="tests", action="append", type=Path)
    ap.add_argument("--output", type=Path)
    args = ap.parse_args()
    try:
        decoder, reducer = args.decoder.resolve(), args.reducer.resolve()
        tests = _paths(args.tests if args.tests else (b.SCANNER_TESTS, b.LIFECYCLE_TESTS, b.REDUCER_TESTS))
        scanner, lifecycle, policy = args.scanner.resolve(), args.lifecycle.resolve(), args.policy.resolve()
        if args.seal_output:
            if args.helper_seal or args.output:
                raise b.BindingError("--seal-output is standalone")
            b.write_exclusive(args.seal_output.resolve(), make_seal(decoder, reducer, scanner=scanner, lifecycle=lifecycle, tests=tests, policy=policy))
            print(json.dumps({"status": "pass", "path": str(args.seal_output.resolve()), "sha256": b.sha256(args.seal_output)}, sort_keys=True))
            return 0
        if args.helper_seal is None or args.output is None:
            raise b.BindingError("--helper-seal, --reducer, and --output are required")
        plan = make_plan(args.helper_seal.resolve(), args.output.resolve(), decoder=decoder, reducer=reducer, scanner=scanner, lifecycle=lifecycle, tests=tests, policy=policy)
        b.validate_seal(plan, args.output.resolve())
        b.write_exclusive(args.output.resolve(), plan)
        b.validate_plan(args.output.resolve())
        print(json.dumps({"status": "pass", "path": str(args.output.resolve()), "sha256": b.sha256(args.output)}, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0978 plan generator rejected: {exc}", file=__import__("sys").stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
