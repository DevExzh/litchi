#!/usr/bin/env python3
"""Serial 0969 runner with a baseline-only A/A launch gate.

The normal command is ``--check-plan``.  ``--run`` is fail-closed and requires
a root-authorized plan, bound binaries, and an empty output directory.  It
never retries a child and never invokes Cargo.  The ``--self-test`` path only
uses synthetic packets and launches no workload.
"""
from __future__ import annotations

import argparse
import json
import os
import signal
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from protocol_common_v4 import (
    ALLOCATOR_PAIRS,
    ALLOCATOR_SAMPLES,
    ALLOCATOR_WARMUPS,
    BOOTSTRAP_RESAMPLES,
    BOOTSTRAP_SEED,
    CI_LIMIT,
    CASES,
    EDITS,
    NATIVE_PAIRS,
    NATIVE_SAMPLES,
    NATIVE_WARMUPS,
    ProtocolError,
    check_file,
    check_packet,
    check_plan,
    check_probe_report,
    cell_id,
    expected_cells,
    load_json,
    reduce_aa_packets,
    validate_go,
    write_json_exclusive,
)

SCHEMA = "litchi.performance.0969.runner.v4"
SYNTHETIC_SHA = "f" * 64
MAX_OUTPUT_BYTES = 8 * 1024 * 1024
TIMEOUT_SECONDS = 900.0


def utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def kill_process(process: subprocess.Popen[bytes]) -> None:
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass


def read_last_json(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    if len(data) > MAX_OUTPUT_BYTES:
        raise ProtocolError(f"stdout exceeds {MAX_OUTPUT_BYTES} bytes")
    lines = [line for line in data.decode("utf-8", errors="strict").splitlines() if line.strip()]
    if not lines:
        raise ProtocolError("child stdout has no JSON report")
    value = json.loads(lines[-1])
    if not isinstance(value, dict):
        raise ProtocolError("child JSON report is not an object")
    return value


def launch_packet(
    *,
    executable: Path,
    args: list[str],
    role: str,
    packet_dir: Path,
    context: dict[str, Any],
    plan: dict[str, Any],
    cell: dict[str, Any],
    allocator: bool,
) -> dict[str, Any]:
    packet_dir.mkdir(parents=True, exist_ok=True)
    packet_id = str(context["packet_id"])
    stdout_path = packet_dir / f"{packet_id}.stdout"
    stderr_path = packet_dir / f"{packet_id}.stderr"
    rss_path = packet_dir / f"{packet_id}.rss"
    command = ["/usr/bin/time", "-f", "%M", "-o", str(rss_path), str(executable), *args]
    terminal: dict[str, Any] = {"returncode": None, "timed_out": False, "launch_error": None}
    started = utc()
    start_ns = time.monotonic_ns()
    try:
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = subprocess.Popen(command, stdout=stdout, stderr=stderr, start_new_session=True)
            try:
                process.wait(timeout=TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                terminal["timed_out"] = True
                kill_process(process)
                process.wait()
            terminal["returncode"] = process.returncode
    except OSError as exc:
        terminal["launch_error"] = str(exc)
    elapsed_wall_ns = time.monotonic_ns() - start_ns
    rss_kb: int | None = None
    validation_error: str | None = None
    if rss_path.is_file():
        try:
            rss_kb = int(rss_path.read_text(encoding="ascii").strip())
        except (OSError, UnicodeDecodeError, ValueError):
            validation_error = "invalid /usr/bin/time RSS output"
    packet: dict[str, Any] = {
        "schema": SCHEMA,
        "packet_id": packet_id,
        "created_utc": started,
        "role": role,
        "command": command,
        "binary": {"path": str(executable), "sha256": context.get("binary_sha256")},
        "context": context,
        "terminal": terminal,
        "wall_ns_excluded_from_metric": elapsed_wall_ns,
        "stdout": {"path": str(stdout_path), "bytes": stdout_path.stat().st_size if stdout_path.exists() else 0},
        "stderr": {"path": str(stderr_path), "bytes": stderr_path.stat().st_size if stderr_path.exists() else 0},
        "rss": {"path": str(rss_path), "bytes": rss_path.stat().st_size if rss_path.exists() else 0, "kb": rss_kb},
        "report": None,
        "validation_error": validation_error,
    }
    if validation_error is None and terminal["returncode"] == 0 and not terminal["timed_out"] and terminal["launch_error"] is None:
        try:
            report = read_last_json(stdout_path)
            check_probe_report(report, cell=cell, fixture=plan["fixtures"][cell["case"]], allocator=allocator)
            packet["report"] = report
        except (OSError, UnicodeError, ValueError, ProtocolError, json.JSONDecodeError) as exc:
            packet["validation_error"] = str(exc)
    elif terminal["returncode"] is None:
        packet["validation_error"] = terminal["launch_error"] or "child did not launch"
    elif terminal["timed_out"]:
        packet["validation_error"] = "timeout; no retry permitted"
    elif terminal["returncode"] != 0:
        packet["validation_error"] = f"child exit status {terminal['returncode']}"
    return packet


def write_packet(packet_root: Path, lane: str, packet: dict[str, Any]) -> None:
    write_json_exclusive(packet_root / lane / f"{packet['packet_id']}.json", packet)


def packet_path(packet_root: Path, packet: dict[str, Any]) -> Path:
    lane = "aa" if packet.get("role") == "aa_baseline" else (packet.get("context") or {}).get("lane")
    if lane not in {"aa", "native", "allocator"}:
        raise ProtocolError(f"{packet.get('packet_id')}: packet lane is missing")
    return packet_root / lane / f"{packet['packet_id']}.json"


def write_capture_manifest(
    output_root: Path,
    plan: dict[str, Any],
    packets: list[dict[str, Any]],
    go_binding: dict[str, Any] | None,
) -> None:
    manifest = {
        "schema": SCHEMA,
        "plan": plan.get("_path"),
        "packets": [str(packet_path(output_root, packet)) for packet in packets],
        "aa_readiness": str(output_root / "aa-readiness-v4.json"),
        "no_retry": True,
    }
    if go_binding is not None:
        manifest["root_go"] = {"path": go_binding["path"], "sha256": go_binding["sha256"]}
    write_json_exclusive(output_root / "capture-manifest-v4.json", manifest)


def load_bound_aa(plan: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Load and revalidate a completed baseline-only A/A receipt."""
    receipt_binding = plan.get("aa_readiness_receipt") or {}
    manifest_binding = plan.get("aa_capture_manifest") or {}
    receipt_path = Path(receipt_binding.get("path", ""))
    manifest_path = Path(manifest_binding.get("path", ""))
    check_file(str(receipt_path), receipt_binding.get("sha256"), "A/A readiness receipt", allow_unbound=False)
    check_file(str(manifest_path), manifest_binding.get("sha256"), "A/A capture manifest", allow_unbound=False)
    receipt = load_json(receipt_path)
    manifest = load_json(manifest_path)
    if manifest.get("schema") != SCHEMA or manifest.get("no_retry") is not True:
        raise ProtocolError("A/A capture manifest schema/retry binding drifted")
    raw_paths = manifest.get("packets")
    if not isinstance(raw_paths, list) or len(raw_paths) != 2 * NATIVE_PAIRS * len(expected_cells()):
        raise ProtocolError("A/A capture manifest packet count drifted")
    if len(set(raw_paths)) != len(raw_paths):
        raise ProtocolError("A/A capture manifest contains duplicate packet paths")
    packets: list[dict[str, Any]] = []
    for raw_path in raw_paths:
        if not isinstance(raw_path, str):
            raise ProtocolError("A/A manifest packet path is not a string")
        path = Path(raw_path)
        if not path.is_absolute():
            path = manifest_path.parent / path
        packet = load_json(path)
        check_packet(packet, plan)
        if packet.get("role") != "aa_baseline":
            raise ProtocolError(f"{path}: bound A/A manifest contains a non-A/A packet")
        packets.append(packet)
    readiness = reduce_aa_packets(packets, plan)
    if receipt != readiness:
        raise ProtocolError("bound A/A readiness receipt does not equal its packet reduction")
    if readiness.get("status") != "pass":
        raise ProtocolError("bound A/A readiness receipt is not a pass")
    return packets, readiness


def binary_for(plan: dict[str, Any], role: str) -> tuple[Path, str | None]:
    row = (plan.get("binaries") or {}).get(role) or {}
    return Path(row.get("path", "")), row.get("sha256")


def child_args(cell: dict[str, Any], warmups: int, samples: int, fixture: dict[str, Any]) -> list[str]:
    # This is the retained perf-cfb-profile CLI shape.  The plan binds the
    # executable and fixture; no shell interpolation is used.
    return [
        "--case", str(cell["case"]),
        "--input", str(fixture["path"]),
        "--edits", str(cell["edits"]),
        "--retention", str(cell["retention"]),
        "--warmup", str(warmups),
        "--samples", str(samples),
    ]


def run_aa(plan: dict[str, Any], packet_root: Path) -> list[dict[str, Any]]:
    packets: list[dict[str, Any]] = []
    executable, digest = binary_for(plan, "native_baseline")
    for identity in expected_cells():
        case, edit_text, retention = identity.split("/")
        cell = {"cell_id": identity, "case": case, "edits": int(edit_text[1:]), "retention": retention}
        fixture = plan["fixtures"][case]
        for pair in range(1, NATIVE_PAIRS + 1):
            order = "AB" if pair <= NATIVE_PAIRS // 2 else "BA"
            for position in ("first", "second"):
                packet_id = f"aa-{case}-e{cell['edits']}-{retention}-p{pair:02d}-{position}"
                context = {
                    "packet_id": packet_id,
                    "lane": "aa",
                    "variant": "baseline",
                    "cell": cell,
                    "pair": pair,
                    "aa_order": order,
                    "aa_position": position,
                    "binary_sha256": digest,
                    "fixture_sha256": fixture["sha256"],
                    "source_manifest_sha256": plan["source"]["baseline"]["manifest_sha256"],
                }
                packet = launch_packet(
                    executable=executable,
                    args=child_args(cell, NATIVE_WARMUPS, NATIVE_SAMPLES, fixture),
                    role="aa_baseline",
                    packet_dir=packet_root / "aa",
                    context=context,
                    plan=plan,
                    cell=cell,
                    allocator=False,
                )
                write_packet(packet_root, "aa", packet)
                packets.append(packet)
    return packets


def run_real_lane(plan: dict[str, Any], packet_root: Path, lane: str) -> list[dict[str, Any]]:
    if lane not in {"native", "allocator"}:
        raise ProtocolError(f"unknown real lane {lane}")
    pairs = NATIVE_PAIRS if lane == "native" else ALLOCATOR_PAIRS
    warmups = NATIVE_WARMUPS if lane == "native" else ALLOCATOR_WARMUPS
    samples = NATIVE_SAMPLES if lane == "native" else ALLOCATOR_SAMPLES
    packets: list[dict[str, Any]] = []
    for identity in expected_cells():
        case, edit_text, retention = identity.split("/")
        cell = {"cell_id": identity, "case": case, "edits": int(edit_text[1:]), "retention": retention}
        fixture = plan["fixtures"][case]
        for pair in range(1, pairs + 1):
            relative_order = "baseline_then_candidate" if pair <= pairs // 2 else "candidate_then_baseline"
            variants = ("baseline", "candidate") if relative_order == "baseline_then_candidate" else ("candidate", "baseline")
            for variant in variants:
                role = f"{lane}_{variant}"
                executable, digest = binary_for(plan, role)
                packet_id = f"{lane}-{case}-e{cell['edits']}-{retention}-p{pair:02d}-{variant}"
                context = {
                    "packet_id": packet_id,
                    "lane": lane,
                    "variant": variant,
                    "cell": cell,
                    "pair": pair,
                    "relative_order": relative_order,
                    "binary_sha256": digest,
                    "fixture_sha256": fixture["sha256"],
                    "source_manifest_sha256": plan["source"][variant]["manifest_sha256"],
                }
                packet = launch_packet(
                    executable=executable,
                    args=child_args(cell, warmups, samples, fixture),
                    role=role,
                    packet_dir=packet_root / lane,
                    context=context,
                    plan=plan,
                    cell=cell,
                    allocator=lane == "allocator",
                )
                write_packet(packet_root, lane, packet)
                packets.append(packet)
    return packets


def capture_sequence(
    plan: dict[str, Any],
    output_root: Path,
    *,
    aa_runner: Callable[[dict[str, Any], Path], list[dict[str, Any]]] = run_aa,
    real_runner: Callable[[dict[str, Any], Path, str], list[dict[str, Any]]] = run_real_lane,
    aa_packets: list[dict[str, Any]] | None = None,
    aa_only: bool = False,
    go_binding: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Run or reuse A/A, then launch formal lanes only after a pass."""
    output_root.mkdir(parents=True, exist_ok=True)
    aa_reused = aa_packets is not None
    if aa_packets is None:
        aa_packets = aa_runner(plan, output_root)
    readiness = reduce_aa_packets(aa_packets, plan)
    write_json_exclusive(output_root / "aa-readiness-v4.json", readiness)
    stage = {
        "schema": "litchi.performance.0969.capture_stage.v4",
        "aa_status": readiness.get("status"),
        "timed_controls_started": False,
        "formal_candidate_launch_permitted": readiness.get("status") == "pass",
        "real_lanes_started": False,
        "no_retry": True,
        "aa_reused": aa_reused,
    }
    if readiness.get("status") != "pass":
        stage["status"] = "stopped_before_candidate_formal"
        write_json_exclusive(output_root / "capture-stage-v4.json", stage)
        return stage
    all_packets = list(aa_packets)
    if aa_only:
        stage["status"] = "aa_complete"
        stage["packet_count"] = len(all_packets)
        write_json_exclusive(output_root / "capture-stage-v4.json", stage)
        write_capture_manifest(output_root, plan, all_packets, go_binding)
        return stage
    for lane in ("native", "allocator"):
        all_packets.extend(real_runner(plan, output_root, lane))
    stage["status"] = "capture_complete"
    stage["real_lanes_started"] = True
    stage["packet_count"] = len(all_packets)
    write_json_exclusive(output_root / "capture-stage-v4.json", stage)
    write_capture_manifest(output_root, plan, all_packets, go_binding)
    return stage


def synthetic_report(cell: dict[str, Any], *, elapsed: int = 100) -> dict[str, Any]:
    oracle = {
        "oracle_ok": True,
        "source_structurally_valid": True,
        "expected_structurally_valid": True,
        "output_structurally_valid": True,
        "semantic_reopen_ok": True,
        "semantic_directory_metadata_matches_expected": True,
        "unchanged_stream_bytes_match_expected": True,
        "unchanged_source_output_stream_bytes_match": True,
        "output_stream_paths_match_expected": True,
        "output_stream_bytes_match_expected": True,
        "root_clsid_preserved": True,
        "storage_clsids_preserved": True,
        "failure_reasons": [],
    }
    correctness = {name: True for name in (
        "all_stages_change_text", "exact_cross_route_outputs", "final_paragraph_semantic_readback",
        "final_public_reader_reopen", "exact_noop_source", "noop_shares_source_allocation",
        "forward_patch_exact_output", "inverse_patch_exact_source", "foreign_patch_refused",
        "stale_patch_refused", "same_value_preserves_retention", "operation_limit_refusal_atomic",
        "retained_capacities_within_ceiling", "zero_route_clears_retention", "release_route_clears_every_stage",
        "untouched_streams_preserved", "directory_metadata_preserved", "semantic_oracles_passed",
        "semantic_change_count_matches_edit_count", "allocator_retained_matches_output_capacity",
    )}
    allocator = {"allocated_bytes": 1000, "deallocated_bytes": 900, "allocation_calls": 10, "peak_live_bytes": 500, "retained_bytes": 100}
    sample = {"index": 0, "output_sha256": "a" * 64, "output_bytes": 1, "output_capacity": 1, "logical_vec_capacity_bytes": 1, "exact_output_equal_reference": True, "final_public_reader_reopen": True, "final_paragraph_semantic_readback": True, "revision_count": 0, "paragraph_count": 1, "semantic_change_count": cell["edits"], "paragraph_zero_sha256": "paragraph", "source_unchanged": False, "elapsed_ns": elapsed, "allocation": allocator, "oracle": oracle}
    rows = [{**sample, "index": index, "elapsed_ns": elapsed, "allocation": None} for index in range(NATIVE_SAMPLES)]
    return {"schema": "litchi.performance.0917.doc_retention.v1", "case": cell["case"], "edits": cell["edits"], "retention": {"zero": "zero", "default": "default_8mib", "release": "release_8mib"}[cell["retention"]], "source_sha256": SYNTHETIC_SHA, "source_bytes": 1, "warmups": NATIVE_WARMUPS, "samples_requested": NATIVE_SAMPLES, "allocator_instrumented": False, "timing_claim": True, "elapsed_ns": [elapsed] * NATIVE_SAMPLES, "samples": rows, "route_outputs": [{"oracle": oracle}, {"oracle": oracle}, {"oracle": oracle}], "correctness": correctness, "output_sha256": "a" * 64, "output_bytes": 1, "output_capacity": 1, "reference_output_sha256": "b" * 64, "reference_output_bytes": 1}


def synthetic_aa_packets(plan: dict[str, Any], *, noisy: bool = False, missing: bool = False) -> list[dict[str, Any]]:
    packets: list[dict[str, Any]] = []
    cell = {"cell_id": "docnohf/e1/zero", "case": "docnohf", "edits": 1, "retention": "zero"}
    for pair in range(1, NATIVE_PAIRS + 1):
        order = "AB" if pair <= NATIVE_PAIRS // 2 else "BA"
        for position in ("first", "second"):
            if missing and pair == NATIVE_PAIRS and position == "second":
                continue
            report = synthetic_report(cell, elapsed=200 if noisy and order == "AB" and position == "second" else 100)
            report["source_sha256"] = SYNTHETIC_SHA
            packet = {"schema": SCHEMA, "packet_id": f"aa-{pair}-{position}", "role": "aa_baseline", "binary": {"path": "", "sha256": None}, "context": {"cell": cell, "pair": pair, "aa_order": order, "aa_position": position, "fixture_sha256": SYNTHETIC_SHA}, "terminal": {"returncode": 0, "timed_out": False, "launch_error": None}, "rss": {"kb": 100}, "report": report, "validation_error": None}
            packets.append(packet)
    return packets


def self_test() -> int:
    plan = {
        "fixtures": {"docnohf": {"path": "", "bytes": 1, "sha256": SYNTHETIC_SHA}},
        "binaries": {"native_baseline": {"path": "", "sha256": None}},
        "aa_noise": {"cells": ["docnohf/e1/zero"]},
    }
    for packet in synthetic_aa_packets(plan):
        packet["binary"] = {"path": "", "sha256": None}
    passed = reduce_aa_packets(synthetic_aa_packets(plan), plan)
    if passed.get("status") != "pass":
        raise SystemExit(f"A/A pass self-test failed: {passed}")
    failed = reduce_aa_packets(synthetic_aa_packets(plan, noisy=True), plan)
    if failed.get("status") != "readiness_fail":
        raise SystemExit(f"A/A readiness-fail self-test failed: {failed}")
    incomplete = reduce_aa_packets(synthetic_aa_packets(plan, missing=True), plan)
    if incomplete.get("status") != "incomplete":
        raise SystemExit(f"A/A incomplete self-test failed: {incomplete}")
    malformed = synthetic_aa_packets(plan)[0]
    malformed["context"] = []
    malformed_result = reduce_aa_packets([malformed], plan)
    if malformed_result.get("status") != "incomplete":
        raise SystemExit(f"malformed packet self-test failed: {malformed_result}")
    malformed_report = synthetic_report({"case": "docnohf", "edits": 1, "retention": "zero"})
    malformed_report["samples"][1]["index"] = 0
    try:
        check_probe_report(
            malformed_report,
            cell={"case": "docnohf", "edits": 1, "retention": "zero"},
            fixture={"sha256": SYNTHETIC_SHA, "bytes": 1},
            allocator=False,
        )
    except ProtocolError:
        pass
    else:
        raise SystemExit("sample index mutation self-test failed")
    calls = {"formal": 0}
    with tempfile.TemporaryDirectory(prefix="litchi-0969-runner-self-test-") as directory:
        root = Path(directory)
        def fake_aa(_plan: dict[str, Any], _root: Path) -> list[dict[str, Any]]:
            return synthetic_aa_packets(plan)
        def fake_real(_plan: dict[str, Any], _root: Path, _lane: str) -> list[dict[str, Any]]:
            calls["formal"] += 1
            return []
        stage = capture_sequence(plan, root, aa_runner=fake_aa, real_runner=fake_real)
        if stage.get("status") != "capture_complete" or calls["formal"] != 2:
            raise SystemExit(f"A/A pass/go self-test failed: {stage} {calls}")
    print(json.dumps({"schema": SCHEMA, "status": "self_test_pass", "aa_pass_go": True, "aa_noise_fail_readiness_fail": True, "aa_missing_incomplete": True, "malformed_packet_incomplete": True, "sample_index_mutation_rejected": True}))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--check-plan", action="store_true")
    parser.add_argument("--allow-unbound", action="store_true")
    parser.add_argument("--run", action="store_true")
    parser.add_argument("--aa-only", action="store_true", help="run only the baseline A/A stage")
    parser.add_argument("--formal-only", action="store_true", help="reuse a bound A/A receipt and run formal lanes")
    parser.add_argument("--go", type=Path, help="explicit root GO receipt required for --run")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--output-root", type=Path)
    args = parser.parse_args()
    if args.self_test:
        return self_test()
    if args.plan is None:
        print("runner rejected input: --plan is required unless --self-test", file=sys.stderr)
        return 2
    if args.aa_only and args.formal_only:
        print("runner rejected input: --aa-only and --formal-only are mutually exclusive", file=sys.stderr)
        return 2
    if (args.aa_only or args.formal_only) and not args.run:
        print("runner rejected input: --aa-only/--formal-only require --run", file=sys.stderr)
        return 2
    if args.run and not (args.aa_only or args.formal_only):
        print("runner rejected input: --run requires exactly one of --aa-only/--formal-only", file=sys.stderr)
        return 2
    if args.run and args.go is None:
        print("runner rejected input: --go is required for --run", file=sys.stderr)
        return 2
    if args.run and args.output_root is None:
        print("runner rejected input: --output-root is required for --run", file=sys.stderr)
        return 2
    try:
        plan_path = args.plan.resolve()
        if args.aa_only:
            plan = check_plan(plan_path, allow_unbound=False, for_aa=True)
        elif args.formal_only:
            plan = check_plan(plan_path, allow_unbound=False, for_capture=True)
        else:
            plan = check_plan(plan_path, allow_unbound=args.allow_unbound or not args.run, for_capture=False)
        if not args.run:
            print(json.dumps({"schema": SCHEMA, "status": "preflight_pass", "plan": str(plan_path), "no_workload_started": True}))
            return 0
        binary_roles = ("native_baseline",) if args.aa_only else ("native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate")
        go_binding = validate_go(args.go, plan_path, plan, binary_roles=binary_roles)
        root = args.output_root.resolve()
        if root.exists() and any(root.iterdir()):
            raise ProtocolError(f"capture output root is not empty; no retry: {root}")
        root.mkdir(parents=True, exist_ok=True)
        plan = {**plan, "_path": str(plan_path)}
        if args.formal_only:
            bound_aa, _ = load_bound_aa(plan)
            result = capture_sequence(plan, root, aa_packets=bound_aa, go_binding=go_binding)
        else:
            result = capture_sequence(plan, root, aa_only=args.aa_only, go_binding=go_binding)
        print(json.dumps(result))
        return 0 if result.get("status") in {"capture_complete", "aa_complete"} else 2
    except (OSError, ProtocolError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(f"runner rejected input: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
