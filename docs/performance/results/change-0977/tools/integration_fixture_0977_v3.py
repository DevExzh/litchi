#!/usr/bin/env python3
"""Offline decoder/reducer integration fixture for one retained 0975 child.

This script never starts perf, a workload, or a build.  It hard-links the
immutable 0975 capture and routes the decoder's four perf subcommands to the
retained 0975/0976 text artifacts.  Decoder outputs are created in a fresh
external directory.  The all-12 exercise duplicates the already decoded
packet in memory solely to test reducer aggregation separately from capture.
"""
from __future__ import annotations

import copy
import datetime as dt
import hashlib
import importlib.util
import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0977")
OUT = ROOT / "integration-v9"
CHILD = OUT / "capture" / "child-01"
PLAN_PATH = OUT / "plan-integration.json"
PACKET_PATH = CHILD / "child-01.packet.json"
RECEIPT = OUT / "integration-receipt-v1.json"

SRC_CHILD = Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01")
SRC_ASLR = Path("/home/zhuhe/litchi-performance-0976/aslr-cache-final-v2-0975-v3")
SRC_META = ROOT / "integration-source-v1"
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
PERF = Path("/usr/bin/perf")
SHIM_SOURCE = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.c")
SHIM_MAP = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.map")
SHIM_SO = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.so")
PARSER_SOURCE_MANIFEST = Path("/home/zhuhe/litchi-performance-0969/manifests/source-manifest-v1.json")

BINARY_SHA = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
HISTORICAL_PERF_SHA = "174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    return {"path": str(path.absolute()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as f:
        json.dump(value, f, indent=2, sort_keys=True, allow_nan=False)
        f.write("\n")


def link_or_fail(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise RuntimeError(f"refusing to overwrite fixture output: {destination}")
    os.link(source, destination)


def copy_or_fail(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise RuntimeError(f"refusing to overwrite fixture output: {destination}")
    shutil.copy2(source, destination)


def import_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def make_plan(decoder: Any) -> dict[str, Any]:
    plan = {
        "schema": "litchi.performance.0977.plan.v1",
        "status": "pre_capture",
        "plan_path": str(PLAN_PATH.absolute()),
        "binary": {**binding(BINARY), "build_id": BUILD_ID},
        # The installed profiler changed after the retained 0975 capture.  The
        # offline fixture binds the currently installed executable and records
        # this explicitly in its receipt; production 0977 custody remains the
        # sealed historical hash in focuseddecode_0977.PERF_SHA256.
        "perf": {
            "path": str(PERF), "sha256": sha256(PERF), "event": "instructions:u",
            "period": 1_000_000, "callgraph": "dwarf,32768", "stack_cap": 32768,
        },
        "marker": {
            "build_id": BUILD_ID, "calls": 85, "dso_path": str(BINARY),
            "marker_scope": "5 warmups + 80 measured calls combined",
            "owner_scope": "all 12 children; missing owner membership unknown",
            "p50_scope": "80 measured samples only", "symbol": MARKER,
        },
        "child_gate": {
            "all_children": True, "lost_events": 0,
            "marker_period_share_positive": True, "min_marker_samples": 100,
            "on_failure": "stop_without_retry",
        },
        "pilot_gate": {
            "child": 1, "lost_events": 0,
            "marker_period_share_positive": True, "min_marker_samples": 100,
            "on_failure": "stop_without_retry",
        },
        "source_manifest": {
            "bytes": PARSER_SOURCE_MANIFEST.stat().st_size,
            "path": str(PARSER_SOURCE_MANIFEST),
            "sha256": sha256(PARSER_SOURCE_MANIFEST),
        },
        "batch": {
            "children": 12, "serial": True, "cpu": 31, "period": 1_000_000,
            "stack_cap": 32768, "warmups": 5, "samples": 80,
            "minimum_marker_samples": 100, "marker_period_positive": True,
            "raw_lost_sum": 0, "terminal_lost_samples": 0,
            "empty_callchain_count": 0, "stop_on_first_failure": True, "retry": False,
        },
        "decoder_interface": {"offline_only": True, "native_ld_preload": "absent"},
        "perf_tool_drift": binding(ROOT / "receipts" / "perf-tool-drift-0977-v1.json"),
        "perf_requalification_posthoc": binding(ROOT / "receipts" / "current-perf-requalification-0977-v1-posthoc-v2.json"),
        "shim_source": binding(SHIM_SOURCE),
        "shim_map": binding(SHIM_MAP),
        "shim_so": binding(SHIM_SO),
    }
    return plan


def prepare_fixture() -> dict[str, Any]:
    if OUT.exists():
        raise RuntimeError(f"fixture output already exists: {OUT}")
    CHILD.mkdir(parents=True)
    # Large capture inputs are hard-linked and never opened for writing.
    link_or_fail(SRC_CHILD / "child-01.perf.data", CHILD / "child-01.perf.data")
    for name in ("launch.json", "report.json"):
        copy_or_fail(SRC_CHILD / f"child-01.{name}", CHILD / f"child-01.{name}")
    terminal = json.loads((SRC_CHILD / "child-01.terminal.json").read_text(encoding="utf-8"))
    terminal_report = CHILD / "child-01.terminal-report.txt"
    terminal_report.write_text("# Total Lost Samples: 0\n", encoding="utf-8")
    terminal["terminal_report"] = str(terminal_report)
    terminal["fixture_note"] = "offline integration fixture; retained terminal fields copied from 0975"
    write_json(CHILD / "child-01.terminal.json", terminal)
    return {
        "perf_data": binding(CHILD / "child-01.perf.data"),
        "perf_mmap_dump_source": binding(SRC_CHILD / "child-01.perf-mmap-dump.txt"),
        "terminal_source": binding(SRC_CHILD / "child-01.terminal.json"),
        "terminal": binding(CHILD / "child-01.terminal.json"),
        "terminal_report": binding(terminal_report),
        "script_source": binding(SRC_ASLR / "original.final-v2.corrected.script.txt"),
        "log_source": binding(SRC_ASLR / "original.final-v2.dwfl.jsonl"),
        "evlist_source": binding(SRC_META / "child-01.perf-evlist-v.txt"),
        "binary": binding(BINARY),
    }


def prepare_runner_contract(decoder: Any, reducer: Any, plan: dict[str, Any]) -> dict[str, Any]:
    """Create the runner-side custody fields needed by the reducer boundary.

    These files are synthetic fixture metadata only.  They describe the
    retained launch and deliberately never claim that a 0977 native runner
    executed.
    """
    plan.update({
        "fixture": binding(Path("/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc")),
        "python": binding(Path("/usr/bin/python3.14")),
        "launcher": binding(Path("/home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py")),
        "taskset": binding(Path("/usr/bin/taskset")),
        "helpers": {"bridge": binding(ROOT / "tools" / "affinity_bridge_0977.py")},
        "command_receipt_contract": {"cwd": "/home/zhuhe/code/litchi"},
    })
    expected = reducer._expected_native_record_argv(plan, CHILD)
    command = {
        "schema": "litchi.performance.0977.command.fixture.v1",
        "argv": expected,
        "cwd": "/home/zhuhe/code/litchi",
        "cwd_policy": "synthetic offline boundary metadata; no native command launched",
        "environment": {"LANG": "C", "LC_ALL": "C", "LD_PRELOAD": "absent", "TZ": "UTC"},
        "native_workload": False,
    }
    command_path = CHILD / "child-01.command.json"
    if command_path.exists():
        command_path.unlink()
    write_json(command_path, command)
    terminal_path = CHILD / "child-01.terminal.json"
    terminal = json.loads(terminal_path.read_text(encoding="utf-8"))
    terminal["command"] = expected
    terminal["affinity_gate"] = {"passed": True, "scope": "synthetic fixture metadata"}
    terminal["process_group_cleanup"] = {"passed": True, "scope": "synthetic fixture metadata"}
    terminal["fixture_scope"] = "offline decoder/reducer integration; no native runner execution"
    terminal_path.unlink()
    write_json(terminal_path, terminal)
    observations = {
        "schema": "litchi.performance.0977.affinity_observations.fixture.v1",
        "rows": [{"bridge_handoff_ready": True, "scope": "synthetic fixture metadata"}],
        "fixture_scope": "does not prove a native 0977 handoff",
    }
    write_json(CHILD / "child-01.affinity-observations.json", observations)
    return {"command": binding(command_path), "terminal": binding(terminal_path), "affinity": binding(CHILD / "child-01.affinity-observations.json")}


def run_decoder(decoder: Any, plan: dict[str, Any], source_bindings: dict[str, Any]) -> dict[str, Any]:
    src_header = SRC_CHILD / "child-01.perf-header.txt"
    src_buildid = SRC_CHILD / "child-01.perf-buildid.txt"
    src_mmap = SRC_CHILD / "child-01.perf-mmap-dump.txt"
    src_script = SRC_ASLR / "original.final-v2.corrected.script.txt"
    src_log = SRC_ASLR / "original.final-v2.dwfl.jsonl"
    src_evlist = SRC_META / "child-01.perf-evlist-v.txt"

    def fake_run(argv: list[str], env: dict[str, str], output: Path, stderr: Path, *, timeout: int = 600) -> None:
        if output.exists() or stderr.exists():
            raise RuntimeError(f"decoder attempted to overwrite output: {output} / {stderr}")
        if len(argv) >= 2 and argv[1] == "report":
            source = src_header
        elif len(argv) >= 2 and argv[1] == "buildid-list":
            source = src_buildid
        elif len(argv) >= 2 and argv[1] == "script" and "-D" in argv:
            source = src_mmap
        elif len(argv) >= 2 and argv[1] == "evlist":
            source = src_evlist
        elif len(argv) >= 2 and argv[1] == "script" and "--show-lost-events" in argv:
            source = src_script
            # The shim output is paired with this command's decoded script.
            log_output = output.parent / f"{output.stem.replace('.perf-script-no-inline', '')}.dwfl.jsonl"
            if log_output.exists():
                raise RuntimeError(f"decoder attempted to overwrite log: {log_output}")
            os.link(src_log, log_output)
        else:
            raise RuntimeError(f"unexpected offline decoder argv: {argv}")
        output.parent.mkdir(parents=True, exist_ok=True)
        os.link(source, output)
        stderr.parent.mkdir(parents=True, exist_ok=True)
        stderr.touch()

    decoder._run_to_file = fake_run
    packet = decoder.decode_child(
        plan, CHILD,
        json.loads((CHILD / "child-01.terminal.json").read_text(encoding="utf-8")),
        json.loads((CHILD / "child-01.launch.json").read_text(encoding="utf-8")),
    )
    # The real runner would emit these sideband files after the decoder.  Add
    # bounded synthetic counterparts so the reducer's custody/loss boundary
    # is exercised without pretending a new runner invocation occurred.
    terminal_summary_path = CHILD / "child-01.terminal-summary.json"
    write_json(terminal_summary_path, {
        "schema": "litchi.performance.0977.terminal-summary.fixture.v1",
        "lost_samples": 0,
        "fixture_scope": "copied retained 0975 terminal zero-loss value",
    })
    loss_path = CHILD / "child-01.loss.json"
    write_json(loss_path, {"schema": "litchi.performance.0977.loss.fixture.v1", "raw_lost_sum": 0, "terminal_lost_samples": 0})
    packet["runner_gate"] = {
        "passed": True,
        "affinity": {"passed": True, "scope": "synthetic fixture metadata"},
        "process_group_cleanup": {"passed": True, "scope": "synthetic fixture metadata"},
    }
    packet["artifacts"]["terminal_summary"] = binding(terminal_summary_path)
    packet["artifacts"]["terminal-summary.json"] = packet["artifacts"]["terminal_summary"]
    packet["artifacts"]["loss.json"] = binding(loss_path)
    packet["artifacts"]["child-01.affinity-observations.json"] = binding(CHILD / "child-01.affinity-observations.json")
    write_json(PACKET_PATH, packet)
    return packet


def run_reducer(reducer: Any, plan: dict[str, Any], packet: dict[str, Any]) -> dict[str, Any]:
    normalized = reducer._decode_packet(packet, PACKET_PATH, plan)
    # The reducer's pure all-12 function is exercised independently from the
    # one-child decoder boundary.  The duplicated packet is synthetic only;
    # no additional capture or decoder call is implied.
    fake_packets = []
    for i in range(1, 13):
        row = copy.deepcopy(normalized)
        row["packet_id"] = f"child-{i:02d}"
        row["child_number"] = i
        row["report"] = copy.deepcopy(normalized["report"])
        fake_packets.append(row)
    manifest = OUT / "synthetic-12-manifest.json"
    write_json(manifest, {
        "schema": "litchi.performance.0977.capture_manifest.v1",
        "status": "complete", "no_retry": True, "serial": True,
        "children_requested": 12, "children_completed": 12,
        "plan": {"path": str(PLAN_PATH), "bytes": PLAN_PATH.stat().st_size, "sha256": sha256(PLAN_PATH)},
        "packets": [f"synthetic:{i:02d}" for i in range(1, 13)],
        "fixture_scope": "in-memory duplicate of one successfully decoded packet; analytics-only",
    })
    result = reducer._reduce_packets(PLAN_PATH, manifest, plan, fake_packets)
    return result


def negative_mutation(reducer: Any, plan: dict[str, Any], packet: dict[str, Any]) -> dict[str, Any]:
    bad = copy.deepcopy(packet)
    bad["perf"]["corrected"]["sample_count"] += 1
    try:
        reducer._decode_packet(bad, PACKET_PATH, plan)
    except Exception as exc:  # exact fail-closed boundary is the assertion
        return {"status": "rejected", "exception": type(exc).__name__, "message": str(exc)[:500]}
    raise RuntimeError("negative packet mutation was accepted")


def main() -> int:
    started = dt.datetime.now(dt.timezone.utc).isoformat()
    decoder_path = ROOT / "tools" / "focuseddecode_0977_v2.py"
    decoder_before = binding(decoder_path)
    decoder = import_module("focuseddecode_0977_v2_fixture", decoder_path)
    reducer = import_module("reduce_0977_fixture_v3", ROOT / "tools" / "reduce_0977.py")
    source_bindings = prepare_fixture()
    plan = make_plan(decoder)
    runner_bindings = prepare_runner_contract(decoder, reducer, plan)
    write_json(PLAN_PATH, plan)
    packet = run_decoder(decoder, plan, source_bindings)
    single = reducer._decode_packet(packet, PACKET_PATH, plan)
    all12 = run_reducer(reducer, plan, packet)
    negative = negative_mutation(reducer, plan, packet)
    decoder_after = binding(decoder_path)
    if decoder_before != decoder_after:
        raise RuntimeError(f"decoder helper changed during final integration: before={decoder_before} after={decoder_after}")
    receipt = {
        "schema": "litchi.performance.0977.decoder-reducer.integration.v1",
        "status": "pass",
        "created_utc": started,
        "workload_started": False,
        "build_started": False,
        "perf_commands_started": False,
        "fixture_root": str(OUT),
        "source_bindings": source_bindings,
        "synthetic_runner_bindings": runner_bindings,
        "installed_perf": {"path": str(PERF), "sha256": sha256(PERF), "bytes": PERF.stat().st_size},
        "perf_binding": {"path": str(PERF), "sha256": sha256(PERF), "bytes": PERF.stat().st_size, "historical_0975_sha256": HISTORICAL_PERF_SHA, "validation_override": False},
        "decoder_helper_before": decoder_before,
        "decoder_helper": decoder_after,
        "reducer_helper": binding(ROOT / "tools" / "reduce_0977.py"),
        "perf_tool_drift": binding(ROOT / "receipts" / "perf-tool-drift-0977-v1.json"),
        "perf_requalification_posthoc": binding(ROOT / "receipts" / "current-perf-requalification-0977-v1-posthoc-v2.json"),
        "decoder": {"packet": binding(PACKET_PATH), "packet_id": packet["packet_id"], "sample_count": packet["perf"]["corrected"]["sample_count"], "marker_sample_count": packet["perf"]["corrected"]["marker_sample_count"], "marker_period": packet["perf"]["corrected"]["marker_period"]},
        "single_packet_reducer": {"status": "pass", "normalized_packet_id": single["packet_id"], "parsed_sample_count": single["_parsed"]["corrected"]["sample_count"], "parsed_marker_sample_count": single["_parsed"]["corrected"]["marker_sample_count"]},
        "all12_analytics": {"status": "pass", "children": len(all12.get("children", [])), "scope": "synthetic in-memory duplicate packet; reducer arithmetic/associations only"},
        "negative_mutation": negative,
        "limitations": ["The retained raw capture is historical; this fixture runs only offline decoder/reducer logic and never launches the current perf executable against a workload.", "The all-12 result is analytics-only duplication of one decoded child and is not cohort evidence."],
    }
    write_json(RECEIPT, receipt)
    print(json.dumps({"status": "pass", "receipt": str(RECEIPT), "receipt_sha256": sha256(RECEIPT), "packet": binding(PACKET_PATH)}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
