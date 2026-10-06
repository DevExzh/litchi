#!/usr/bin/env python3
"""Serial, fail-closed runner for the prospective 0941 XLSX controls.

The runner is intentionally inert until ``xlsx-plan-v4.json`` has status
``authorized`` and ``--execute`` is supplied. It launches each fixed child
exactly once, records the outer terminal receipt before parsing the probe
report, and aborts on the first failure without retrying or replacing a
measurement.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import time


ROOT = Path(__file__).resolve().parent
PLAN_PATH = ROOT / "xlsx-plan-v4.json"
SCRIPT_PATH = ROOT / "xlsx-capture-v4.py"
RUNTIME_ENV = {
    "PATH": "/usr/bin:/bin",
    "LANG": "C",
    "LC_ALL": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
}
FIXTURE_NAMES = {
    "generated-xlsx": "generated-xlsx-medium-source.xlsx",
    "conditional-formatting": "conditional-formatting-source.xlsx",
    "small-xlsx": "small-xlsx-source.xlsx",
    "boundary-xlsx": "boundary-xlsx-source.xlsx",
}
EXPECTED_BINARY_IDENTITY = {
    "native": ("litchi-xlsx-allocation-probe-0779", "none", None),
    "allocator": ("litchi-xlsx-allocation-probe-0779-alloc", "system_allocator_operation_scoped", "serialized_region_peak_v3"),
}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def read(path: Path) -> dict:
    try:
        value = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"cannot read JSON {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise SystemExit(f"JSON object required: {path}")
    return value


def write(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def artifact(path: Path) -> dict:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def validate_plan(plan: dict, require_authorized: bool = True) -> None:
    require(plan.get("schema") == "litchi.performance.0941.xlsx-controls-plan.v4", "plan schema changed")
    if require_authorized:
        require(plan.get("status") == "authorized", "capture requires explicit plan status authorized")
    source = plan["source_and_probe"]
    require(source["probe_lock_sha256"] == "9f5e7da95e9af1ec0a35c21fd1c2daa6081d14c54031553cdca56a8804d26d3f", "probe lock witness changed")
    require(source["toolchain"] == "1.95.0", "toolchain changed")
    require(plan["scope"]["cpu"] == 12, "CPU changed")
    require(plan["scope"]["no_retry_or_replacement"] is True, "retry policy changed")
    require(plan["scope"]["terminal_first"] is True, "terminal-first policy changed")
    require(plan["scope"]["native"] == {
        "blocks": 6, "warmup": 3, "samples": 30,
        "orders": [["baseline", "candidate"], ["candidate", "baseline"],
                    ["baseline", "candidate"], ["candidate", "baseline"],
                    ["candidate", "baseline"], ["baseline", "candidate"]],
        "children_per_cell": 12, "total_children": 72,
    }, "native lane changed")
    require(plan["scope"]["allocator"] == {
        "blocks": 2, "warmup": 0, "samples": 3,
        "orders": [["baseline", "candidate"], ["candidate", "baseline"]],
        "children_per_cell": 4, "total_children": 24,
    }, "allocator lane changed")
    require(plan.get("runtime_environment") == {
        "PATH": "/usr/bin:/bin",
        "LANG": "C",
        "LC_ALL": "C",
        "TZ": "UTC",
        "RUST_BACKTRACE": "0",
        "unset": ["RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "CARGO_HOME", "RUSTUP_TOOLCHAIN", "CARGO_INCREMENTAL"],
    }, "runtime environment changed")
    expected_cells = [
        ("generated-xlsx/open", "generated-xlsx", "open", "Sheet1", "primary"),
        ("generated-xlsx/lifecycle", "generated-xlsx", "lifecycle", "Sheet1", "primary"),
        ("conditional-formatting/open", "conditional-formatting", "open", "Home", "primary"),
        ("conditional-formatting/lifecycle", "conditional-formatting", "lifecycle", "Home", "primary"),
        ("small-xlsx/open", "small-xlsx", "open", "Sheet1", "boundary-control"),
        ("boundary-xlsx/open", "boundary-xlsx", "open", "Sheet1", "boundary-control"),
    ]
    require(len(plan["cells"]) == len(expected_cells), "cell count changed")
    seen: set[str] = set()
    for cell, expected in zip(plan["cells"], expected_cells):
        require(tuple(cell.get(key) for key in ("id", "fixture", "phase", "sheet", "group")) == expected, f"cell changed: {cell}")
        require(cell["id"] not in seen, f"duplicate cell: {cell['id']}")
        seen.add(cell["id"])
        require("iwork" not in cell["id"].lower(), "iWork cell present")
    require(len(seen) == len(expected_cells), "cell identity is not unique")
    tools = plan.get("tool_witnesses", {})
    for name in ("time", "taskset"):
        tool = Path(tools[name]["path"])
        require(tool.is_file() and sha(tool) == tools[name]["sha256"], f"{name} witness changed")
    script = Path(plan["script_witness"]["path"])
    require(script == SCRIPT_PATH and sha(script) == plan["script_witness"]["sha256"], "runner script witness changed")
    validator = Path(plan["validator_witness"]["path"])
    require(validator.is_file() and sha(validator) == plan["validator_witness"]["sha256"], "validator witness changed")
    for leg in ("baseline", "candidate"):
        source_leg = source["legs"][leg]
        source_receipt = Path(source_leg["source_receipt"])
        require(source_receipt.is_file(), f"{leg} source receipt missing")
        require(sha(source_receipt) == plan["source_receipt_witnesses"][leg]["sha256"], f"{leg} source receipt changed")
        require(sha(Path(source_leg["probe_root"], "Cargo.lock")) == source["probe_lock_sha256"], f"{leg} probe lock changed")
        for variant in ("normal", "allocator"):
            binary = source_leg["normal_binary"] if variant == "normal" else source_leg["allocator_binary"]
            witness = plan["binary_witnesses"][leg][variant]
            require(Path(binary) == Path(witness["path"]), f"{leg} {variant} binary path changed")
            require(Path(binary).is_file() and os.access(binary, os.X_OK), f"{leg} {variant} binary missing")
            require(artifact(Path(binary)) == witness, f"{leg} {variant} binary witness changed")
        for lane in ("native", "allocator"):
            for kind in ("start", "terminal"):
                witness = plan["build_witnesses"][leg][lane][kind]
                path = Path(witness["path"])
                require(path.is_file() and sha(path) == witness["sha256"], f"{leg} {lane} {kind} witness changed")


def fixture_path(plan: dict, leg: str, fixture_id: str) -> Path:
    path = Path(plan["source_and_probe"]["legs"][leg]["fixture_root"]) / FIXTURE_NAMES[fixture_id]
    expected = plan["fixtures"][fixture_id]
    require(path.is_file(), f"fixture missing: {path}")
    require(path.stat().st_size == expected["bytes"] and sha(path) == expected["sha256"], f"fixture identity changed: {path}")
    return path


def binary_path(plan: dict, leg: str, lane: str) -> Path:
    key = "normal_binary" if lane == "native" else "allocator_binary"
    path = Path(plan["source_and_probe"]["legs"][leg][key])
    witness = plan["binary_witnesses"][leg]["normal" if lane == "native" else "allocator"]
    require(path == Path(witness["path"]), f"binary path witness changed: {path}")
    require(path.is_file() and os.access(path, os.X_OK), f"binary missing or not executable: {path}")
    require(artifact(path) == witness, f"binary hash/size witness changed: {path}")
    return path


def kill_group(process: subprocess.Popen[bytes]) -> None:
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        process.wait()


def process_group_clear(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    return False


def read_rss(path: Path) -> int:
    require(path.is_file(), f"/usr/bin/time RSS file missing: {path}")
    lines = [line.strip() for line in path.read_text().splitlines() if line.strip()]
    require(len(lines) == 1 and lines[0].isdigit(), f"invalid raw max RSS: {path}")
    return int(lines[0])


def nonnegative_int(value: object, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0, f"{label} is not a nonnegative integer")


def validate_report(report: dict, plan: dict, cell: dict, leg: str, lane: str, samples: int, warmup: int) -> None:
    expected_source = plan["fixtures"][cell["fixture"]]
    require(report.get("schema") == "litchi.xlsx.allocation-probe.v1", "probe schema changed")
    require(report.get("tool") == "litchi-xlsx-allocation-attribution-probe-0779", "probe identity changed")
    require(report.get("phase") == cell["phase"], "report phase changed")
    expected_scope = {
        "open": "Workbook::open(source) only",
        "lifecycle": "Workbook::open, semantic A1 edit and commit, then save_with_durability(destination, NoSync)",
    }[cell["phase"]]
    require(report.get("timing_scope") == expected_scope, "report timing scope changed")
    require(report.get("durability") == ("no-sync" if cell["phase"] == "lifecycle" else None), "report durability changed")
    require(report.get("sheet") == cell["sheet"], "report sheet changed")
    require(report.get("address") == "A1", "report address changed")
    require(report.get("marker") == "litchi-perf-0638-ordinary-save", "marker changed")
    require(report.get("warmup") == warmup and report.get("samples_requested") == samples, "sample configuration changed")
    source = report.get("source", {})
    require(source.get("bytes") == expected_source["bytes"] and source.get("sha256") == expected_source["sha256"], "report source identity changed")
    expected_binary, expected_instrumentation, expected_revision = EXPECTED_BINARY_IDENTITY[lane]
    allocator = report.get("allocator", {})
    expected_allocator = "CountingSystemAllocator(std::alloc::System)" if lane == "allocator" else "Rust system allocator"
    require(tuple(allocator.get(key) for key in ("binary", "instrumentation", "allocator", "counter_revision")) == (expected_binary, expected_instrumentation, expected_allocator, expected_revision), "report allocator identity changed")
    require(isinstance(report.get("samples"), list) and len(report["samples"]) == samples, "sample count changed")
    for index, sample in enumerate(report["samples"]):
        require(isinstance(sample, dict), f"sample is not an object at {index}")
        require(isinstance(sample.get("index"), int) and not isinstance(sample.get("index"), bool) and sample.get("index") == index, f"sample index changed at {index}")
        nonnegative_int(sample.get("elapsed_ns"), f"elapsed sample {index}")
        if cell["phase"] == "lifecycle":
            verification = sample.get("verification")
            require(isinstance(verification, dict), f"lifecycle verification missing at {index}")
            require((verification.get("reopened"), verification.get("sheet"), verification.get("address"), verification.get("expected"), verification.get("actual"), verification.get("marker_matches")) == (True, cell["sheet"], "A1", "litchi-perf-0638-ordinary-save", "litchi-perf-0638-ordinary-save", True), f"lifecycle verification changed at {index}")
            expected = plan["expected_published_outputs_from_0779"][cell["id"]]
            nonnegative_int(sample.get("published_bytes"), f"published bytes {index}")
            require(sample.get("published_bytes") == expected["bytes"] and sample.get("published_sha256") == expected["sha256"], "published output witness changed")
        else:
            require(sample.get("published_bytes") is None and sample.get("published_sha256") is None, "open unexpectedly published output")
        if lane == "native":
            require(sample.get("allocation") is None, f"native report contains allocation metrics at {index}")
        else:
            allocation = sample.get("allocation")
            require(isinstance(allocation, dict) and allocation.get("status") == "measured", "allocator sample is not measured")
            require(allocation.get("scope") == "operation_global_system_allocator", "allocator scope changed")
            fields = ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes")
            for field in fields:
                nonnegative_int(allocation.get(field), f"allocation {field} {index}")
            require(allocation["allocation_calls"] >= allocation["reallocation_calls"], f"allocation/reallocation counter ordering changed at {index}")
            require(allocation["live_bytes_after"] == allocation["live_bytes_before"] + allocation["allocated_bytes"] - allocation["deallocated_bytes"], f"allocation live-byte conservation changed at {index}")
            require(allocation["region_peak_live_bytes"] >= allocation["live_bytes_before"] and allocation["region_peak_live_bytes"] >= allocation["live_bytes_after"], f"allocation region peak ordering changed at {index}")
            require(allocation["peak_live_bytes_before"] >= allocation["live_bytes_before"] and allocation["peak_live_bytes_after"] >= allocation["live_bytes_after"] and allocation["peak_live_bytes_after"] >= allocation["peak_live_bytes_before"] and allocation["peak_live_bytes_after"] >= allocation["region_peak_live_bytes"], f"allocation peak ordering changed at {index}")
            require(allocation["failed_allocation_calls"] == 0, f"failed allocation observed at {index}")


def run_child(plan: dict, output_root: Path, lane: str, block: int, cell: dict, leg: str) -> dict:
    leg_info = plan["source_and_probe"]["legs"][leg]
    source = fixture_path(plan, leg, cell["fixture"])
    binary = binary_path(plan, leg, lane)
    lane_cfg = plan["scope"][lane]
    slug = cell["id"].replace("/", "__")
    stem = f"{block:02d}-{slug}-{leg}"
    stdout_path = output_root / f"{stem}.stdout-v2.log"
    stderr_path = output_root / f"{stem}.stderr-v2.log"
    rss_path = output_root / f"{stem}.maxrss-kib-v2.txt"
    report_path = output_root / f"{stem}.report-v2.json"
    terminal_path = output_root / f"{stem}.terminal-v2.json"
    for path in (stdout_path, stderr_path, rss_path, report_path, terminal_path):
        require(not path.exists(), f"refusing to overwrite {path}")
    command = [
        "/usr/bin/time", "-f", "%M", "-o", str(rss_path), "--",
        "/usr/bin/taskset", "--cpu-list", str(plan["scope"]["cpu"]), str(binary),
        "--source", str(source), "--sheet", cell["sheet"], "--phase", cell["phase"],
        "--samples", str(lane_cfg["samples"]), "--warmup", str(lane_cfg["warmup"]), "--output", str(report_path),
    ]
    binary_before = artifact(binary)
    source_artifact = artifact(source)
    started = time.time()
    stdout_stream = stdout_path.open("wb")
    stderr_stream = stderr_path.open("wb")
    process = subprocess.Popen(
        command,
        cwd=leg_info["probe_root"],
        stdout=stdout_stream,
        stderr=stderr_stream,
        start_new_session=True,
        env=RUNTIME_ENV,
    )
    timed_out = False
    try:
        process.wait(timeout=plan["scope"]["deadline_seconds"])
    except subprocess.TimeoutExpired:
        timed_out = True
        kill_group(process)
    stdout_stream.close()
    stderr_stream.close()
    ended = time.time()
    group_clear = process_group_clear(process.pid)
    if not group_clear:
        kill_group(process)
        group_clear = process_group_clear(process.pid)
    rss_artifact = artifact(rss_path) if rss_path.is_file() else {"path": str(rss_path), "missing": True}
    binary_after_terminal = artifact(binary)
    tool_witnesses = plan["tool_witnesses"]
    terminal = {
        "schema": "litchi.performance.0941.xlsx-child-terminal.v3",
        "lane": lane,
        "block": block,
        "cell": cell["id"],
        "leg": leg,
        "pid": process.pid,
        "process_group": process.pid,
        "command": command,
        "cwd": leg_info["probe_root"],
        "environment": RUNTIME_ENV,
        "environment_exact": True,
        "binary_before": binary_before,
        "binary_after": binary_after_terminal,
        "source": source_artifact,
        "script": plan["script_witness"],
        "time": tool_witnesses["time"],
        "taskset": tool_witnesses["taskset"],
        "started_utc": datetime.fromtimestamp(started, timezone.utc).isoformat().replace("+00:00", "Z"),
        "ended_utc": datetime.fromtimestamp(ended, timezone.utc).isoformat().replace("+00:00", "Z"),
        "elapsed_seconds": ended - started,
        "return_code": process.returncode,
        "timed_out": timed_out,
        "process_group_clear": group_clear,
        "stdout": artifact(stdout_path),
        "stderr": artifact(stderr_path),
        "rss": rss_artifact,
        "report": {"path": str(report_path)},
        "terminal_recorded_before_report_parse": True,
    }
    # This write is deliberately before reading the report or semantically
    # parsing any child result.
    write(terminal_path, terminal)
    require(not timed_out and process.returncode == 0 and group_clear, f"child failed or process group leaked; terminal retained at {terminal_path}")
    require(binary_after_terminal == binary_before, "binary changed during child")
    report = read(report_path)
    samples = lane_cfg["samples"]
    warmup = lane_cfg["warmup"]
    validate_report(report, plan, cell, leg, lane, samples, warmup)
    rss_kib = read_rss(rss_path)
    return {
        "schema": "litchi.performance.0941.xlsx-child-receipt.v3",
        "lane": lane,
        "block": block,
        "cell": cell["id"],
        "fixture": cell["fixture"],
        "phase": cell["phase"],
        "leg": leg,
        "command": command,
        "binary": artifact(binary),
        "source": source_artifact,
        "report": artifact(report_path),
        "terminal": artifact(terminal_path),
        "stdout": artifact(stdout_path),
        "stderr": artifact(stderr_path),
        "max_rss_kib": rss_kib,
        "rss": rss_artifact,
    }


def capture(plan: dict, lane: str, output_root: Path) -> dict:
    require(not output_root.exists(), f"refusing to overwrite capture directory {output_root}")
    output_root.mkdir(parents=True)
    receipts: list[dict] = []
    lane_cfg = plan["scope"][lane]
    try:
        for cell in plan["cells"]:
            for block, order in enumerate(lane_cfg["orders"]):
                for leg in order:
                    row = run_child(plan, output_root, lane, block, cell, leg)
                    receipts.append(row)
                    write(output_root / "receipts-v3.json", receipts)
                    print(f"{lane} {cell['id']} block={block} {leg} passed", flush=True)
    except Exception as exc:
        write(output_root / "aborted-v3.json", {"schema": "litchi.performance.0941.xlsx-capture-aborted.v3", "lane": lane, "ended_utc": utc_now(), "children_completed": len(receipts), "error": str(exc), "receipts": receipts})
        raise
    write(output_root / "receipts-v3.json", receipts)
    complete = {"schema": "litchi.performance.0941.xlsx-capture-complete.v3", "lane": lane, "ended_utc": utc_now(), "children": len(receipts), "receipts": artifact(output_root / "receipts-v3.json"), "expected_children": plan["scope"][lane]["total_children"]}
    write(output_root / "complete-v3.json", complete)
    return complete


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--lane", choices=("native", "allocator"), required=True)
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--plan", type=Path, default=PLAN_PATH)
    args = parser.parse_args()
    plan = read(args.plan)
    if not args.execute:
        validate_plan(plan, require_authorized=False)
        print(json.dumps({"status": plan["status"], "capture": "not_started", "reason": "pass --execute after explicit authorization"}, indent=2, sort_keys=True))
        return 0
    validate_plan(plan, require_authorized=True)
    output_root = Path(plan["output_layout"]["root"]) / args.lane
    result = capture(plan, args.lane, output_root)
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except RuntimeError as exc:
        raise SystemExit(f"capture refused: {exc}") from exc
