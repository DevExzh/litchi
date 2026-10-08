#!/usr/bin/env python3
"""Replay one retained 0977 child through the 0978 attribution proof.

This is an offline-only integration replay.  It reads an already retained raw
scanner sidecar and lifecycle stream, reparses the corrected perf-script text,
rebuilds the lifecycle ledger, and runs the reducer's exact linkage and
conservation checks.  It never invokes perf, a workload, Cargo, or a build.
Inputs are immutable and every bound file is hashed before and after replay.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import sys
from pathlib import Path
from typing import Any, Mapping

ROOT = Path("/home/zhuhe/litchi-performance-0978")
CAPTURE = ROOT / "cpu-profile-child02-v1/capture/child-02"
ANALYSIS = ROOT / "analysis/reducer-replay-child02-v1"
TOOLS = ROOT / "tools"
RAW_SUMMARY = ANALYSIS / "raw-summary.json"
LIFECYCLE_STREAM = ANALYSIS / "lifecycle-stream.json"
SAVED_LEDGER = ANALYSIS / "ledger-replay.json"
CORRECTED = CAPTURE / "child-02.corrected.perf-script-no-inline.txt"
RAW_DUMP = CAPTURE / "child-02.perf-mmap-dump.txt"
PERF_DATA = CAPTURE / "child-02.perf.data"
EVLIST = CAPTURE / "child-02.perf-evlist-v.txt"
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
INTEGRATION_RECEIPT = ROOT / "analysis/child-02-build-ledger-integration-v2.json"
MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
EXPECTED_PID = 3849325
EXPECTED_CPU = 31
PERIOD = 1_000_000


class ReplayFailure(RuntimeError):
    pass


def fail(message: str) -> None:
    raise ReplayFailure(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    path = path.absolute()
    if not path.is_file():
        fail(f"missing bound input: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        fail(f"cannot load {path}: {exc}")
    if not isinstance(value, dict):
        fail(f"JSON root is not an object: {path}")
    return value


def load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        fail(f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:  # noqa: BLE001
        fail(f"import failed for {path}: {exc}")
    return module


def lifecycle_rows(built: Mapping[str, Any], corrected: list[Mapping[str, Any]]) -> list[dict[str, Any]]:
    samples = built.get("samples")
    if not isinstance(samples, list) or len(samples) != len(corrected):
        fail("built lifecycle sample count does not match corrected stream")
    rows: list[dict[str, Any]] = []
    for sample, corrected_row in zip(samples, corrected):
        if not isinstance(sample, Mapping):
            fail("built lifecycle sample is not an object")
        row = dict(sample)
        raw_identity = sample.get("raw_identity")
        if not isinstance(raw_identity, Mapping):
            fail("built lifecycle sample lacks raw identity")
        for key in ("time_ns", "pid", "tid", "cpu", "period"):
            row[key] = raw_identity.get(key)
        owners = sorted(
            {
                f"{frame.get('dso')}::{frame.get('symbol_base')}"
                for frame in corrected_row.get("frames", [])
                if frame.get("resolved") is True
                and frame.get("dso") not in {None, "", "[unknown]"}
                and frame.get("symbol_base") not in {None, "", "[unknown]", "unknown"}
            }
        )
        row["inclusive_owners"] = owners if row.get("marker") else []
        rows.append(row)
    return rows


def write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    path = path.absolute()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
    except FileExistsError:
        fail(f"refusing to overwrite existing output: {path}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="new replay receipt; must not already exist")
    parser.add_argument("--ledger-output", type=Path, required=True, help="new derived ledger; must not already exist")
    parser.add_argument("--lifecycle-stream", type=Path, default=LIFECYCLE_STREAM, help="retained lifecycle stream JSON")
    parser.add_argument("--stream-generation-receipt", type=Path, default=INTEGRATION_RECEIPT, help="receipt binding the stream's generating helper")
    args = parser.parse_args(argv)

    lifecycle_stream_path = args.lifecycle_stream.resolve()
    stream_generation_receipt_path = args.stream_generation_receipt.resolve()

    source_paths = {
        "replay_script": Path(__file__),
        "reducer": TOOLS / "reduce_0978.py",
        "focused_decoder": TOOLS / "focuseddecode_0978.py",
        "lifecycle_helper": TOOLS / "lifecycle_ledger_0978.py",
        "raw_scanner": TOOLS / "raw_stack_0978.py",
        "raw_summary": RAW_SUMMARY,
        "lifecycle_stream": lifecycle_stream_path,
        "saved_ledger": SAVED_LEDGER,
        "corrected_script": CORRECTED,
        "raw_mmap_dump": RAW_DUMP,
        "perf_data": PERF_DATA,
        "perf_evlist": EVLIST,
        "binary": BINARY,
        "stream_generation_receipt": stream_generation_receipt_path,
    }
    before = {name: binding(path) for name, path in source_paths.items()}
    script_sha = before["replay_script"]["sha256"]
    raw_obj = load_json(RAW_SUMMARY)
    stream_obj = load_json(lifecycle_stream_path)
    saved_obj = load_json(SAVED_LEDGER)
    generation_obj = load_json(stream_generation_receipt_path)

    # The retained stream's generation receipt is separate from this replay.
    # Bind its recorded parser hash explicitly; do not silently relabel the
    # retained stream as having been generated by a later helper revision.
    recorded_generation = generation_obj.get("inputs", {}).get("lifecycle_helper", {})
    if not isinstance(recorded_generation, Mapping):
        fail("stream-generation receipt lacks lifecycle helper binding")
    generation_helper_sha = recorded_generation.get("sha256")
    if not isinstance(generation_helper_sha, str) or len(generation_helper_sha) != 64:
        fail("stream-generation receipt has no lifecycle helper SHA-256")

    if raw_obj.get("status") != "complete" or raw_obj.get("sample_count") != 7201:
        fail("retained raw scanner sidecar is not complete with 7201 samples")
    if stream_obj.get("status") != "complete" or stream_obj.get("sample_descriptor_count") != 7201:
        fail("retained lifecycle stream is not complete with 7201 samples")
    if stream_obj.get("raw_lost_sum") != 0 or stream_obj.get("terminal_lost_values") not in ([], [0]):
        fail("retained lifecycle stream has nonzero loss")
    raw_rows = raw_obj.get("sample_ledger")
    if not isinstance(raw_rows, list):
        fail("raw scanner sidecar lacks sample_ledger")

    reducer = load_module("reducer_0978_replay", TOOLS / "reduce_0978.py")
    focused = load_module("focuseddecode_0978_replay", TOOLS / "focuseddecode_0978.py")
    lifecycle = load_module("lifecycle_ledger_0978_replay", TOOLS / "lifecycle_ledger_0978.py")
    corrected_focus, corrected_lost = focused._parse_corrected(
        CORRECTED,
        expected_pid=EXPECTED_PID,
        expected_cpu=EXPECTED_CPU,
        expected_period=PERIOD,
        marker=MARKER,
        marker_dso=str(BINARY),
    )
    if corrected_lost != 0 or len(corrected_focus) != len(raw_rows):
        fail(f"corrected stream census/loss mismatch: rows={len(corrected_focus)} loss={corrected_lost}")
    identity = [
        {
            "index": row["index"],
            "time_ns": row["time_ns"],
            "pid": row["pid"],
            "tid": row["tid"],
            "cpu": row["cpu"],
            "period": row["period"],
            "frames": row["frames"],
            "marker": any(frame.get("symbol_base") == MARKER and frame.get("dso") == str(BINARY) for frame in row["frames"]),
        }
        for row in corrected_focus
    ]
    focused._pair_samples(identity, raw_rows)
    cfg = {"binary": BINARY, "build_id": BUILD_ID, "marker": MARKER, "marker_dso": str(BINARY)}
    attributions, _corrected_with_raw = focused._make_attributions(cfg, lifecycle, stream_obj, identity)
    built = lifecycle.build_ledger(
        stream_obj,
        raw_rows,
        expected_pid=EXPECTED_PID,
        expected_cpu=EXPECTED_CPU,
        expected_period=PERIOD,
        target_path=str(BINARY),
        target_build_id=BUILD_ID,
        target_comm=BINARY.name[:15],
        terminal_lost_samples=0,
        marker=MARKER,
        marker_dso=str(BINARY),
        attributions=attributions,
    )
    if built.get("status") != "complete":
        fail("independent lifecycle build did not return complete")

    corrected = reducer._parse_corrected_text(CORRECTED, marker=MARKER, marker_dso=str(BINARY), expected_period=PERIOD)
    raw = reducer._parse_structured_rows({"sample_ledger": raw_rows}, "raw replay", marker=MARKER, marker_dso=str(BINARY))
    rows = lifecycle_rows(built, corrected)
    actual = reducer._derive_child(corrected, raw, rows, built["ledger"], marker_dso=str(BINARY))
    attr_obj = {"attributions": {str(index): value for index, value in attributions.items()}}
    reducer._validate_exact_attribution_linkage(
        attr_obj,
        corrected,
        marker=MARKER,
        marker_dso=str(BINARY),
        target_build_id=BUILD_ID,
        allow_synthetic=False,
    )
    reducer._validate_marker_map_linkage(
        lifecycle,
        stream_obj,
        attributions,
        expected_pid=EXPECTED_PID,
        target_comm=BINARY.name[:15],
        marker_dso=str(BINARY),
        target_build_id=BUILD_ID,
    )
    reducer._compare_built_ledger_to_actual(built["ledger"], actual)
    saved_ledger = saved_obj.get("ledger")
    saved_ledger_match = saved_ledger == built.get("ledger")

    # Emit the current independent result as a new versioned sidecar.  The
    # original ledger-replay.json is retained as an immutable comparison input.
    ledger_output = args.ledger_output.absolute()
    ledger_payload = {
        "schema": "litchi.performance.0978.lifecycle_ledger.v1",
        "status": "complete",
        "scope": "one retained 0977 child-02 offline replay; no 0978 native custody",
        "target": built.get("target"),
        "lifecycle": built.get("lifecycle"),
        "gates": built.get("gates"),
        "ledger": built.get("ledger"),
        "samples": built.get("samples"),
    }
    write_exclusive(ledger_output, ledger_payload)

    after = {name: binding(path) for name, path in source_paths.items()}
    for name in source_paths:
        if before[name] != after[name]:
            fail(f"bound input changed during replay: {name}")

    result_ledger = built["ledger"]
    receipt = {
        "schema": "litchi.performance.0978.child02_offline_reducer_replay.v3",
        "status": "pass",
        "scope": {
            "kind": "offline_analysis_replay",
            "child": "child-02",
            "capture_origin": "retained 0977 capture copied into 0978 custody directory",
            "native_actions": {"build": False, "recording": False, "workload": False},
            "cohort_admitted": False,
            "0978_native_plan_claim": False,
        },
        "command": {
            "argv": [sys.executable, "-B", str(Path(__file__).absolute()), "--output", str(args.output.absolute()), "--ledger-output", str(ledger_output), "--lifecycle-stream", str(lifecycle_stream_path), "--stream-generation-receipt", str(stream_generation_receipt_path)],
            "cwd": os.getcwd(),
            "offline_only": True,
            "no_perf_subprocess": True,
        },
        "source_bindings": {
            "replay_script": before["replay_script"],
            "reducer": before["reducer"],
            "focused_decoder": before["focused_decoder"],
            "lifecycle_helper_used_for_replay": before["lifecycle_helper"],
            "raw_scanner_source": before["raw_scanner"],
            "stream_generation_receipt": before["stream_generation_receipt"],
            "stream_generation_lifecycle_helper_sha256": generation_helper_sha,
        },
        "inputs": {
            name: before[name]
            for name in ("raw_summary", "lifecycle_stream", "saved_ledger", "corrected_script", "raw_mmap_dump", "perf_data", "perf_evlist", "binary")
        },
        "outputs": {
            "derived_ledger": binding(ledger_output),
            "receipt": {"path": str(args.output.absolute()), "script_hash_bound_before_write": script_sha},
        },
        "checks": {
            "raw_sample_count": len(raw_rows),
            "lifecycle_record_count": stream_obj.get("record_count"),
            "lifecycle_sample_descriptor_count": stream_obj.get("sample_descriptor_count"),
            "corrected_sample_count": len(corrected),
            "corrected_lost_events": corrected_lost,
            "exact_frame_and_marker_linkage": True,
            "active_map_build_id_linkage": True,
            "saved_ledger_matches_current_helper": saved_ledger_match,
            "input_hashes_stable_before_after": True,
        },
        "ledger_summary": {
            **{
                key: result_ledger.get(key)
                for key in ("D", "M", "R", "T", "U_target", "U_marker", "marker_sample_count", "conservation")
            },
            # Tknown is a reducer-derived view; the lifecycle ledger stores
            # T and U_target rather than duplicating this scalar.
            "Tknown": actual.get("Tknown"),
        },
        "limitations": [
            "one child only; no all-12 reduction or cohort admission",
            "the retained capture originated from 0977; this replay does not claim a 0978 native command/plan custody result",
            "D/M/R/T are sampled instruction periods; R means marker ancestry not observed and may include missing/truncated marker ancestry",
            "no speedup, exact timed-region, optimization, or causal owner claim",
            "the original ledger-replay.json is preserved and compared; the derived ledger output is a fresh version",
        ],
    }
    write_exclusive(args.output, receipt)
    print(json.dumps({"status": "pass", "receipt": str(args.output.absolute()), "ledger": str(ledger_output), "saved_ledger_matches_current_helper": saved_ledger_match, "D": result_ledger.get("D"), "M": result_ledger.get("M"), "R": result_ledger.get("R"), "T": result_ledger.get("T"), "U_marker": result_ledger.get("U_marker"), "U_target": result_ledger.get("U_target")}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ReplayFailure as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}, sort_keys=True), file=sys.stderr)
        raise SystemExit(1)
