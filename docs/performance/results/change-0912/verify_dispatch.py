"""Independently verify the retained 0912 GDB dispatch witnesses.

This reducer is read-only with respect to captured children.  It hashes and
reads the retained plan, commands, GDB logs, dispatch state machines, and
workload receipts; it never launches GDB, the benchmark, a profiler, or a
compiler.  The only write is the requested analysis JSON.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
BINARY = PACKET.parent / "change-0910" / "binaries" / "native"
CAPTURE_SCRIPT = PACKET / "capture.py"
PLAN = PACKET / "plan.json"
PLAN_GDB = PACKET / "plan-gdb.json"
SYMBOL_RANGES = PACKET / "symbol-ranges.json"
SYMBOLS = PACKET / "symbols.txt"
DISPATCH_GDB = PACKET / "dispatch_gdb.py"
DISPATCH_HELPER = PACKET / "dispatch.py"
OWNER = "litchi_pptx::package::model::Package::opened_presentation_transaction"
PARENT = "litchi_perf_baseline::pptx_owned_benchmark::timed_iteration"
SETUP = "litchi_perf_baseline::build_pptx_source_edit_corpus"
HARDWARE = "sha2::sha256::x86_sha::compress"
SOFTWARE = "sha2::sha256::soft::unroll::compress"
ROUTES = ("materialized", "owned")
GDB_WARNING = "/usr/bin/gdb: warning: Couldn't determine a path for the index cache directory."


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def check_descriptor(row: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(row, dict), f"{label}: descriptor is not an object")
    actual = descriptor(Path(row["path"]))
    require(actual == row, f"{label}: descriptor changed")
    return actual


def load_capture_module():
    spec = importlib.util.spec_from_file_location("capture_0912_verify", CAPTURE_SCRIPT)
    require(spec is not None and spec.loader is not None, "cannot load capture helper")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def parse_maps(text: str, binary_path: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for line in text.splitlines():
        fields = line.split()
        if len(fields) < 6 or fields[5] != binary_path:
            continue
        try:
            rows.append(
                {
                    "start": int(fields[0], 16),
                    "end": int(fields[1], 16),
                    "size": int(fields[2], 16),
                    "offset": int(fields[3], 16),
                    "perms": fields[4],
                    "path": fields[5],
                }
            )
        except ValueError as error:
            fail(f"invalid mapping row: {line!r}: {error}")
    return rows


def in_range(value: int, row: dict[str, Any]) -> bool:
    return row["start"] <= value < row["end"]


def compact_frame(frame: dict[str, Any], bias: int) -> dict[str, Any]:
    pc = int(frame["pc"])
    return {"name": frame.get("name"), "pc": pc, "offset": pc - bias}


def verify_perf_failure(plan_gdb: dict[str, Any]) -> dict[str, Any]:
    """Bind the rejected first perf child without admitting its samples."""

    folder = PACKET / "capture" / "materialized-a"
    command_path = folder / "command.json"
    command = read_json(command_path)
    require(command["exit_code"] == 0, "rejected perf child did not exit zero")
    command_descriptor = descriptor(command_path)
    raw_rows = [row for row in command["artifacts"] if row["path"].endswith("/perf.data")]
    require(len(raw_rows) == 1, "rejected perf child must retain one perf.data descriptor")
    raw = check_descriptor(raw_rows[0], "rejected perf raw data")

    debug_meta = read_json(folder / "sample-debug.json")
    debug_stdout = Path(debug_meta["stdout"]["path"])
    check_descriptor(debug_meta["stdout"], "perf sample-debug stdout")
    try:
        debug_text = gzip.decompress(debug_stdout.read_bytes()).decode("utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"cannot decode perf lost-event log: {error}")
    lost = [int(value) for value in re.findall(r"PERF_RECORD_LOST lost (\d+)", debug_text)]
    require(lost == [47, 36, 64, 99, 26, 32], f"unexpected lost-event chunks: {lost}")
    require(sum(lost) == 304, "unexpected lost-event total")

    script_meta = read_json(folder / "decode.json")
    script_path = Path(script_meta["script"]["path"])
    check_descriptor(script_meta["script"], "rejected perf script")
    try:
        script_text = gzip.decompress(script_path.read_bytes()).decode("utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"cannot decode perf script: {error}")
    owner_frames = script_text.count(OWNER)
    hardware_frames = script_text.count(HARDWARE)
    require(owner_frames == 0, f"rejected perf script unexpectedly resolves owner: {owner_frames}")
    require(hardware_frames == 1, f"expected one unscoped hardware frame, got {hardware_frames}")
    require(OWNER not in script_text[script_text.find(HARDWARE) :], "hardware frame became scoped unexpectedly")

    return {
        "status": "rejected",
        "reason": "lost perf records and missing exact owner stack scope",
        "child": "materialized-a",
        "command": command_descriptor,
        "raw_perf_data": raw,
        "lost_chunks": len(lost),
        "lost_records": sum(lost),
        "lost_chunk_values": lost,
        "resolved_owner_frame_occurrences": owner_frames,
        "resolved_hardware_frame_occurrences": hardware_frames,
        "remaining_perf_children_stopped": True,
        "perf_samples_reused": False,
        "timings_reused": False,
    }


def verify_symbols(symbol_data: dict[str, Any]) -> dict[str, dict[str, Any]]:
    expected = {
        OWNER: (0x43E780, 0x177),
        PARENT: (0x3BD7A0, 0x1DBF),
        SETUP: (0x3A2C40, 0x2765),
        "sha2::sha256::compress256": (0x5CB2D0, 0x4B),
        HARDWARE: (0x5CB380, 0x3BC),
        SOFTWARE: (0x5CB740, 0x272D),
    }
    rows = {row["name"]: row for row in symbol_data["ranges"]}
    require(set(rows) == set(expected), "symbol range set changed")
    for name, (address, size) in expected.items():
        row = rows[name]
        require(row["address"] == address and row["size"] == size, f"symbol range changed: {name}")
        require(row["end"] == address + size, f"symbol end changed: {name}")
    lines = [line.strip() for line in SYMBOLS.read_text(encoding="utf-8").splitlines() if line.strip()]
    require(len(lines) == len(expected), "symbols.txt row count changed")
    for name, (address, size) in expected.items():
        expected_line = f"{address:016x} {size:016x} "
        require(any(line.startswith(expected_line) and line.endswith(" " + name) for line in lines), f"symbols.txt missing {name}")
    return rows


def verify_leg(
    kind: str,
    index: int,
    plan: dict[str, Any],
    plan_gdb: dict[str, Any],
    symbol_rows: dict[str, dict[str, Any]],
    capture_module: Any,
) -> dict[str, Any]:
    folder = PACKET / "dispatch" / kind
    start_path = folder / "start.json"
    command_path = folder / "command.json"
    dispatch_path = folder / "dispatch.json"
    semantics_path = folder / "semantics.json"
    start = read_json(start_path)
    command = read_json(command_path)
    dispatch = read_json(dispatch_path)
    saved_semantics = read_json(semantics_path)

    binary = plan_gdb["binary"]["path"]
    script_path = plan_gdb["gdb_script"]["path"]
    artifacts_path = str((folder / "artifacts").resolve())
    expected_argv = [
        "/usr/bin/taskset", "-c", "12", "/usr/bin/gdb", "--batch", "-nx", "-nh",
        "-iex", "set auto-load off", "-iex", "set debuginfod enabled off", "-x",
        script_path, "--args", binary, kind, "0", "1", artifacts_path,
    ]
    expected_environment = dict(
        plan_gdb["environment"],
        LITCHI_0912_DISPATCH_LOG=str(dispatch_path.resolve()),
        LITCHI_0912_SYMBOL_RANGES=str(SYMBOL_RANGES.resolve()),
    )

    require(command["argv"] == expected_argv, f"{kind}: GDB argv changed")
    require(start["argv"] == expected_argv, f"{kind}: start argv changed")
    require(command["environment"] == expected_environment, f"{kind}: command environment changed")
    require(start["environment"] == expected_environment, f"{kind}: start environment changed")
    require(command["cwd"] == plan_gdb["cwd"] == start["cwd"], f"{kind}: cwd changed")
    require(command["exit_code"] == 0, f"{kind}: GDB exited nonzero")
    require(start["plan"] == descriptor(PLAN_GDB), f"{kind}: start plan custody changed")
    require(command["plan"] == descriptor(PLAN_GDB), f"{kind}: command plan custody changed")
    require(start["binary"] == plan_gdb["binary"] == command["binary"], f"{kind}: binary custody changed")
    require(start["helper"] == plan_gdb["helper"] == command["helper"], f"{kind}: helper custody changed")
    require(start["gdb_script"] == plan_gdb["gdb_script"] == command["gdb_script"], f"{kind}: GDB script custody changed")
    require(start["row"] == {"index": index, "kind": kind}, f"{kind}: start row changed")
    require(command["row"] == {"index": index, "kind": kind}, f"{kind}: command row changed")

    required_artifacts = {
        str((folder / "artifacts" / "output.pptx").resolve()),
        str((folder / "artifacts" / "report.json").resolve()),
        str((folder / "artifacts" / "source.pptx").resolve()),
        str(dispatch_path.resolve()),
        str(start_path.resolve()),
        str((folder / "stdout.txt").resolve()),
        str((folder / "stderr.txt").resolve()),
    }
    actual_artifacts = {row["path"] for row in command["artifacts"]}
    require(actual_artifacts == required_artifacts, f"{kind}: command artifact census changed")
    for artifact in command["artifacts"]:
        check_descriptor(artifact, f"{kind}: command artifact")

    stderr_lines = (folder / "stderr.txt").read_text(encoding="utf-8").splitlines()
    require(stderr_lines == [GDB_WARNING], f"{kind}: unexpected GDB stderr")
    stdout = (folder / "stdout.txt").read_text(encoding="utf-8")
    require("exited normally" in stdout, f"{kind}: target did not report normal exit")

    semantic_row = {"index": index, "kind": kind}
    expected_semantics = capture_module.validate_semantics(plan_gdb, semantic_row, folder)
    require(saved_semantics == expected_semantics, f"{kind}: semantics receipt differs from independent helper validation")

    require(dispatch["status"] == "pass", f"{kind}: dispatch status is not pass")
    require(dispatch["errors"] == [], f"{kind}: dispatch errors present")
    require(dispatch["exit_code"] == 0, f"{kind}: target exit code missing")
    require(dispatch["startup"] == "", f"{kind}: startup unexpectedly nonempty")
    require(dispatch["mapping_load_bias"] > 0, f"{kind}: invalid mapping load bias")
    bias = dispatch["mapping_load_bias"]
    mappings = parse_maps(dispatch["mappings"], binary)
    zero_maps = [row for row in mappings if row["offset"] == 0]
    require(len(zero_maps) == 1 and zero_maps[0]["start"] == bias, f"{kind}: mapping base mismatch")
    require(zero_maps[0]["perms"].startswith("r--"), f"{kind}: zero-offset mapping permissions changed")

    runtime = dispatch["runtime_symbols"]
    require(set(runtime) == set(symbol_rows), f"{kind}: runtime symbol set changed")
    for name, row in symbol_rows.items():
        value = runtime[name]
        require(value["file_address"] == row["address"], f"{kind}: {name} file address mismatch")
        require(value["size"] == row["size"], f"{kind}: {name} size mismatch")
        require(value["load_bias"] == bias, f"{kind}: {name} load bias mismatch")
        require(value["start"] == bias + row["address"], f"{kind}: {name} runtime start mismatch")
        require(value["end"] == bias + row["end"], f"{kind}: {name} runtime end mismatch")

    entries = dispatch["owner_entries"]
    returns = dispatch["owner_returns"]
    require([entry["context"] for entry in entries] == ["setup", "measured"], f"{kind}: owner entry order changed")
    require([entry["context"] for entry in returns] == ["setup", "measured"], f"{kind}: owner return order changed")
    require([entry["index"] for entry in entries] == [0, 1], f"{kind}: owner entry indices changed")
    require([row["owner_entry"] for row in returns] == [0, 1], f"{kind}: owner return indices changed")
    require(all(entry["entry_pc"] == runtime[OWNER]["start"] for entry in entries), f"{kind}: owner entry PC mismatch")
    require(all(entry["unwind_error"] is None for entry in entries), f"{kind}: owner entry unwind error")
    require(all(row["backend_breakpoints_disabled"] is True for row in returns), f"{kind}: return backend state not disabled")

    expected_parents = {"setup": SETUP, "measured": PARENT}
    caller_rows: list[dict[str, Any]] = []
    for entry in entries:
        context = entry["context"]
        parent_name = expected_parents[context]
        return_pc = entry["return_address"]
        return_offset = return_pc - bias
        parent_row = symbol_rows[parent_name]
        require(in_range(return_offset, {"start": parent_row["address"], "end": parent_row["end"]}), f"{kind}: {context} raw return caller mismatch")
        frames = entry["frames"]
        names = [frame.get("name") for frame in frames]
        require(names.count(OWNER) == 1 and names.count(parent_name) == 1, f"{kind}: {context} entry backtrace caller mismatch")
        owner_frame = next(frame for frame in frames if frame.get("name") == OWNER)
        parent_frame = next(frame for frame in frames if frame.get("name") == parent_name)
        require(in_range(owner_frame["pc"] - bias, {"start": symbol_rows[OWNER]["address"], "end": symbol_rows[OWNER]["end"]}), f"{kind}: {context} owner frame range mismatch")
        require(in_range(parent_frame["pc"] - bias, {"start": parent_row["address"], "end": parent_row["end"]}), f"{kind}: {context} parent frame range mismatch")
        caller_rows.append({
            "index": entry["index"],
            "context": context,
            "entry_pc": entry["entry_pc"],
            "entry_rsp": entry["entry_rsp"],
            "return_pc": return_pc,
            "return_offset": return_offset,
            "caller": parent_name,
            "caller_range": [parent_row["address"], parent_row["end"]],
            "frame_backtrace_exact": True,
        })

    return_rows: list[dict[str, Any]] = []
    for entry, row in zip(entries, returns):
        require(row["context"] == entry["context"] and row["owner_entry"] == entry["index"], f"{kind}: return does not match entry")
        require(row["pc"] == entry["return_address"], f"{kind}: return PC does not match saved RSP word")
        require(row["rsp"] == entry["entry_rsp"] + 8, f"{kind}: return RSP is not entry RSP+8")
        return_rows.append({
            "owner_entry": row["owner_entry"],
            "context": row["context"],
            "return_pc": row["pc"],
            "entry_rsp": entry["entry_rsp"],
            "return_rsp": row["rsp"],
            "rsp_delta": row["rsp"] - entry["entry_rsp"],
            "backend_breakpoints_disabled": row["backend_breakpoints_disabled"],
        })

    events = dispatch["events"]
    require([event["kind"] for event in events] == ["owner_entry", "owner_return", "owner_entry", "backend", "owner_return"], f"{kind}: event order changed")
    require(events[0] == {"backend_breakpoints_enabled": {"hardware": False, "software": False}, "context": "setup", "kind": "owner_entry", "owner_entry": 0}, f"{kind}: setup enable event changed")
    require(events[1] == {"backend_breakpoints_disabled": True, "context": "setup", "kind": "owner_return", "owner_entry": 0}, f"{kind}: setup return event changed")
    require(events[2] == {"backend_breakpoints_enabled": {"hardware": True, "software": True}, "context": "measured", "kind": "owner_entry", "owner_entry": 1}, f"{kind}: measured enable event changed")
    require(events[3] == {"backend": "hardware", "breakpoint_disabled_after_hit": True, "kind": "backend", "owner_entry": 1}, f"{kind}: backend event changed")
    require(events[4] == {"backend_breakpoints_disabled": True, "context": "measured", "kind": "owner_return", "owner_entry": 1}, f"{kind}: measured return event changed")

    hits = dispatch["backend_hits"]
    require(len(hits) == 1, f"{kind}: expected exactly one backend hit")
    hit = hits[0]
    require(hit["kind"] == "hardware" and hit["symbol"] == HARDWARE, f"{kind}: backend is not exact hardware SHA")
    require(hit["owner_entry"] == 1, f"{kind}: backend hit outside measured owner")
    require(hit["pc"] == runtime[HARDWARE]["start"], f"{kind}: hardware PC is not exact breakpoint range start")
    require(hit["frame_scope_available"] is True and hit["unwind_error"] is None, f"{kind}: hardware backtrace scope unavailable")
    frames = hit["frames"]
    names = [frame.get("name") for frame in frames]
    require(names.count(HARDWARE) == 1, f"{kind}: hardware frame count changed")
    require(names.count(OWNER) == 1 and names.count(PARENT) == 1, f"{kind}: hardware backtrace lacks owner/parent")
    require(SOFTWARE not in names and SETUP not in names, f"{kind}: unexpected software/setup frame in hardware hit")
    owner_index = names.index(OWNER)
    parent_index = names.index(PARENT)
    require(owner_index < parent_index, f"{kind}: owner is not below measured parent in backtrace")
    owner_frame = frames[owner_index]
    parent_frame = frames[parent_index]
    require(in_range(owner_frame["pc"] - bias, {"start": symbol_rows[OWNER]["address"], "end": symbol_rows[OWNER]["end"]}), f"{kind}: hit owner frame outside range")
    require(in_range(parent_frame["pc"] - bias, {"start": symbol_rows[PARENT]["address"], "end": symbol_rows[PARENT]["end"]}), f"{kind}: hit parent frame outside range")
    require(OWNER in hit["backtrace"] and PARENT in hit["backtrace"] and HARDWARE in hit["backtrace"], f"{kind}: textual backtrace lacks exact frames")
    require("owner-entry" in hit["scope_proof"] and "matched owner return" in hit["scope_proof"], f"{kind}: scope proof changed")

    return {
        "route": kind,
        "index": index,
        "status": "pass",
        "command": {
            "argv": command["argv"],
            "cwd": command["cwd"],
            "environment": command["environment"],
            "exit_code": command["exit_code"],
            "descriptor": descriptor(command_path),
        },
        "custody": {
            "start": descriptor(start_path),
            "dispatch": descriptor(dispatch_path),
            "semantics": descriptor(semantics_path),
            "artifacts_verified": True,
            "plan_descriptor": descriptor(PLAN_GDB),
            "binary_descriptor": plan_gdb["binary"],
            "helper_descriptor": plan_gdb["helper"],
            "gdb_script_descriptor": plan_gdb["gdb_script"],
        },
        "semantic_validation": {
            "method": "capture.validate_semantics",
            "status": "pass",
            "operations": expected_semantics["operations"],
            "source": expected_semantics["source"],
            "output": expected_semantics["output"],
        },
        "mapping": {
            "load_bias": bias,
            "load_bias_hex": hex(bias),
            "zero_offset_mapping": zero_maps[0],
            "runtime_ranges_match_map_base": True,
        },
        "owner_scope": {
            "entries": caller_rows,
            "returns": return_rows,
            "matched_returns": True,
            "return_rsp_is_entry_rsp_plus_8": True,
            "event_order_exact": True,
            "backend_breakpoints_disabled_outside_measured_owner": True,
        },
        "hardware_backend": {
            "status": "observed",
            "symbol": HARDWARE,
            "pc": hit["pc"],
            "pc_offset": hit["pc"] - bias,
            "owner_entry": hit["owner_entry"],
            "frame_scope_available": hit["frame_scope_available"],
            "exact_owner_and_parent_frames": True,
            "full_backtrace_retained": True,
            "software_backend_hit": False,
            "first_hit_breakpoint_disabled": True,
        },
        "warnings": {
            "gdb_index_cache_warning_only": True,
            "stderr": GDB_WARNING,
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=PACKET / "dispatch-analysis.json")
    args = parser.parse_args()

    plan = read_json(PLAN)
    plan_gdb = read_json(PLAN_GDB)
    symbol_data = read_json(SYMBOL_RANGES)
    capture_module = load_capture_module()

    require(plan_gdb["schema"] == "litchi.native-sha-dispatch-0912.v1", "GDB plan schema changed")
    require(plan_gdb["binary"] == plan["binary"], "GDB and perf binary custody differ")
    require(plan_gdb["parent"] == plan["evidence_parent"], "GDB parent custody differs")
    require(plan_gdb["cpu"] == 12 and plan_gdb["samples"] == 1 and plan_gdb["warmup"] == 0, "GDB plan sampling changed")
    require(plan_gdb["schedule"] == [{"index": 0, "kind": "materialized"}, {"index": 1, "kind": "owned"}], "GDB schedule changed")
    require(plan_gdb["cwd"] == "/tmp/litchi-0912-gdb-run", "GDB cwd changed")
    require(plan_gdb["environment"] == {"DEBUGINFOD_URLS": "", "LANG": "C", "LC_ALL": "C", "PATH": "/usr/bin:/bin", "TZ": "UTC"}, "GDB environment changed")
    require(plan_gdb["gdb_version"] == "GNU gdb (Ubuntu 17.1-2ubuntu1) 17.1", "GDB version changed")
    require(plan_gdb["binary"] == descriptor(BINARY), "retained binary changed")
    for row in plan_gdb["inputs"]:
        check_descriptor(row, "GDB plan input")
    for key in ("helper", "gdb_script"):
        check_descriptor(plan_gdb[key], f"GDB plan {key}")
    for row in plan_gdb["tools"]:
        check_descriptor(row, "GDB plan tool")

    symbol_rows = verify_symbols(symbol_data)
    perf_failure = verify_perf_failure(plan_gdb)
    legs = [verify_leg(kind, index, plan, plan_gdb, symbol_rows, capture_module) for index, kind in enumerate(ROUTES)]

    analysis = {
        "schema": "litchi.native-sha-dispatch-analysis-0912.v1",
        "status": "pass",
        "decision": "admit_abi_scoped_native_hardware_sha_reachability",
        "reviewer": "owning_profile_review_0911",
        "scope": "Independent offline verification of two retained GDB witnesses. Native backend reachability only; no perf attribution, timing, Ir, or route comparison.",
        "checks": {
            "plan_helper_command_environment_custody": True,
            "binary_and_symbol_ranges": True,
            "capture_validate_semantics_samples_1": True,
            "runtime_ranges_match_pie_map_base": True,
            "raw_return_address_caller_classification": True,
            "owner_entry_return_event_order": True,
            "matched_return_pc_and_rsp_plus_8": True,
            "backend_breakpoints_disabled_outside_measured_owner": True,
            "hardware_sha_pc_and_frame_checks": True,
            "gdb_index_cache_warning_bound": True,
            "rejected_perf_packet_retained": True,
        },
        "perf_primary": perf_failure,
        "legs": legs,
        "dispatch": {
            "routes": list(ROUTES),
            "hardware_backend_observed_in_each_measured_owner": True,
            "software_backend_observed_in_measured_owner": False,
            "full_backtraces_include_owner_and_timed_iteration": True,
            "both_children_target_exit_zero": True,
            "native_timing_claim": False,
            "perf_sample_claim": False,
        },
        "hash_bindings": {
            "plan_gdb": descriptor(PLAN_GDB),
            "verify_dispatch": descriptor(Path(__file__)),
            "dispatch_helper": descriptor(DISPATCH_HELPER),
            "gdb_script": descriptor(DISPATCH_GDB),
            "symbol_ranges": descriptor(SYMBOL_RANGES),
            "symbols": descriptor(SYMBOLS),
            "binary": descriptor(BINARY),
            "protocol_review_gdb": descriptor(PACKET / "protocol-review-gdb.md"),
            "first_perf_failure_command": perf_failure["command"],
        },
    }
    args.output = args.output.resolve()
    args.output.write_text(json.dumps(analysis, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": analysis["status"], "output": str(args.output), "legs": len(legs)}, sort_keys=True))


if __name__ == "__main__":
    main()

