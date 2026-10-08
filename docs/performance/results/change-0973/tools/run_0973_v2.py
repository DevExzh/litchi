#!/usr/bin/env python3
"""Serial 0973 diagnostic runner; no retry and explicit root GO required."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import signal
import subprocess
import sys
from pathlib import Path
from typing import Any

import protocol_common_0973_v2 as p
import raw_stack_0973_v2 as raw_stack


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def pin_cpu() -> None:
    os.sched_setaffinity(0, {p.CPU})


def native_env() -> dict[str, str]:
    result = os.environ.copy()
    result.update(p.ENVIRONMENT)
    result.pop("LD_PRELOAD", None)
    for key in list(result):
        if key.startswith("LITCHI_0973_"):
            result.pop(key, None)
    return result


def run_process(argv: list[str], *, stdout: Any = None, stderr: Any = None,
                timeout: int | None = None, pin: bool = False) -> subprocess.CompletedProcess[str]:
    if stdout is None:
        stdout = subprocess.PIPE
    if stderr is None:
        stderr = subprocess.PIPE
    return subprocess.run(
        argv, stdin=subprocess.DEVNULL, stdout=stdout, stderr=stderr,
        text=True, env=native_env(), timeout=timeout,
        preexec_fn=pin_cpu if pin else None, start_new_session=pin, check=False,
    )


def terminate_group(proc: subprocess.Popen[Any], *, grace_seconds: float = 1.0) -> None:
    try:
        os.killpg(proc.pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    try:
        proc.wait(timeout=grace_seconds)
        return
    except subprocess.TimeoutExpired:
        pass
    try:
        os.killpg(proc.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    proc.wait(timeout=grace_seconds)


def perf_script_argv(perf_data: Path) -> list[str]:
    return [
        p.PERF_PATH, "script", "--show-lost-events", "--no-inline",
        "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso",
        "-i", str(perf_data),
    ]


def child_record_argv(plan: dict[str, Any], child_dir: Path) -> list[str]:
    binary = Path(plan["binary"]["path"])
    fixture = Path(plan["fixture"]["path"])
    data = child_dir / f"{child_dir.name}.perf.data"
    metadata = child_dir / f"{child_dir.name}.launch.json"
    report = child_dir / f"{child_dir.name}.report.json"
    return [
        p.TASKSET_PATH, "--cpu-list", str(p.CPU), p.PERF_PATH, "record",
        "-e", "instructions:u", "-c", str(p.PERIOD), "--call-graph", p.CALLGRAPH,
        "--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache",
        "--user-callchains", "-o", str(data), "--", p.PYTHON_PATH,
        p.LAUNCHER_PATH, "--metadata", str(metadata), "--requested-affinity", "[31]",
        "--", str(binary), "--case", "docfloat", "--input", str(fixture),
        "--edits", "1", "--retention", "default", "--warmup", str(p.WARMUPS),
        "--samples", str(p.SAMPLES),
    ]


def write_binding(path: Path, value: Any) -> dict[str, Any]:
    p.write_json_exclusive(path, value)
    return p.file_binding(path)


def run_decoder_mode(mode: str, perf_data: Path, child_dir: Path,
                     layout: dict[str, Any], so: Path, binary: Path,
                     expected_pid: int) -> tuple[Path, Path, dict[str, Any]]:
    script_path = child_dir / f"{child_dir.name}.{mode}.perf-script-no-inline.txt"
    stderr_path = child_dir / f"{child_dir.name}.{mode}.perf.stderr"
    log_path = child_dir / f"{child_dir.name}.{mode}.dwfl.jsonl"
    env = native_env()
    argv = perf_script_argv(perf_data)
    if mode in {"noop", "corrected"}:
        env.update({
            "LITCHI_0973_SHIM_MODE": mode,
            "LITCHI_0973_SHIM_LOG": str(log_path),
            "LITCHI_0973_TARGET_PATH": str(binary),
            "LITCHI_0973_TARGET_NAME": binary.name,
            "LITCHI_0973_CORRECTED_BASE": layout["corrected_base"],
            "LITCHI_0973_PERF_EXPECTED_START": layout["perf_expected_start"],
            "LD_PRELOAD": str(so),
        })
    cp = subprocess.run(argv, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE, text=True, env=env, timeout=300,
                        check=False)
    if cp.returncode != 0:
        raise p.ProtocolError(f"{mode} perf script failed: {cp.stderr[-500:]}")
    p.write_text_exclusive(script_path, cp.stdout)
    p.write_text_exclusive(stderr_path, cp.stderr)
    if mode == "stock":
        log_path = child_dir / f"{child_dir.name}.stock-no-log"
        p.write_text_exclusive(log_path, "stock mode has no LD_PRELOAD diagnostic log\n")
    elif not log_path.is_file():
        raise p.ProtocolError(f"{mode} shim did not create a log")
    summary = p.parse_perf_script_strict(
        cp.stdout, marker_dso=str(binary.absolute()), expected_pid=expected_pid, expected_cpu=p.CPU,
    )
    return script_path, stderr_path, summary


def decode_child(child_dir: Path, perf_data: Path, layout: dict[str, Any],
                 plan: dict[str, Any], launch: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Path]]:
    binary = Path(plan["binary"]["path"])
    so = Path(plan["shim"]["shared_object"]["path"])
    expected_pid = int(launch["exec_pid"])
    stock_script, stock_stderr, stock = run_decoder_mode("stock", perf_data, child_dir, layout, so, binary, expected_pid)
    noop_script, noop_stderr, noop = run_decoder_mode("noop", perf_data, child_dir, layout, so, binary, expected_pid)
    corrected_script, corrected_stderr, corrected = run_decoder_mode("corrected", perf_data, child_dir, layout, so, binary, expected_pid)
    if stock_script.read_bytes() != noop_script.read_bytes() or stock_stderr.read_bytes() != noop_stderr.read_bytes():
        raise p.ProtocolError("stock/no-op decoder control is not byte-identical")
    inline_path = child_dir / f"{child_dir.name}.perf-script-inline.txt"
    inline = run_process([p.PERF_PATH, "script", "--show-lost-events", "--inline", "-i", str(perf_data)], timeout=300)
    if inline.returncode != 0:
        raise p.ProtocolError(f"perf inline decode failed: {inline.stderr!r}")
    p.write_text_exclusive(inline_path, inline.stdout)
    header_path = child_dir / f"{child_dir.name}.perf-header.txt"
    header = run_process([p.PERF_PATH, "report", "--header-only", "--stdio", "-i", str(perf_data)], timeout=300)
    if header.returncode != 0:
        raise p.ProtocolError(f"perf header decode failed: {header.stderr!r}")
    p.write_text_exclusive(header_path, header.stdout)
    buildid_path = child_dir / f"{child_dir.name}.perf-buildid.txt"
    buildid = run_process([p.PERF_PATH, "buildid-list", "-i", str(perf_data)], timeout=300)
    if buildid.returncode != 0:
        raise p.ProtocolError(f"perf buildid-list failed: {buildid.stderr!r}")
    p.write_text_exclusive(buildid_path, buildid.stdout)
    p.check_buildid_export(buildid.stdout, binary_path=str(binary.absolute()), build_id=p.BINARY_BUILD_ID)
    # Do not use a marker-share number as a timed-region percentage.
    stock["control_stdout_sha256"] = p.sha256(stock_script)
    noop["control_stdout_sha256"] = p.sha256(noop_script)
    corrected["control_stdout_sha256"] = p.sha256(corrected_script)
    summary = {
        "stock": stock,
        "noop": noop,
        "corrected": corrected,
        "control": {
            "stock_noop_stdout_byte_identical": True,
            "stock_noop_stderr_byte_identical": True,
            "sample_count_identity": len({stock["sample_count"], noop["sample_count"], corrected["sample_count"]}) == 1,
            "total_period_identity": len({stock["total_period"], noop["total_period"], corrected["total_period"]}) == 1,
            "all_modes_lost_events_zero": all(row["lost_events"] == 0 for row in (stock, noop, corrected)),
        },
        "marker_scope": "5 warmups + 80 measured calls combined; no exact timed-region attribution",
        "physical_leaf_scope": "frame 0 only; unresolved frame 0 is explicit unknown_leaf, never reassigned to an ancestor",
    }
    artifacts = {
        "stock_script": stock_script, "stock_stderr": stock_stderr,
        "stock_no_log": child_dir / f"{child_dir.name}.stock-no-log",
        "noop_script": noop_script, "noop_stderr": noop_stderr,
        "noop_log": child_dir / f"{child_dir.name}.noop.dwfl.jsonl",
        "corrected_script": corrected_script, "corrected_stderr": corrected_stderr,
        "corrected_log": child_dir / f"{child_dir.name}.corrected.dwfl.jsonl",
        "inline": inline_path, "header": header_path, "buildid": buildid_path,
    }
    for name, path in artifacts.items():
        if not path.is_file():
            raise p.ProtocolError(f"missing decoder artifact {name}")
    return summary, artifacts


def child_gate(corrected: dict[str, Any]) -> tuple[bool, dict[str, Any]]:
    reasons: list[str] = []
    if corrected.get("marker_sample_count") < 100:
        reasons.append("marker_sample_count_lt_100")
    if corrected.get("lost_events") != 0:
        reasons.append("lost_events_nonzero")
    share = corrected.get("marker_period_share")
    if share is None or share <= 0:
        reasons.append("marker_period_share_nonpositive")
    return not reasons, {"passed": not reasons, "reasons": reasons}


def run_child(plan: dict[str, Any], child_number: int, capture_root: Path) -> dict[str, Any]:
    child_id = f"child-{child_number:02d}"
    child_dir = capture_root / child_id
    child_dir.mkdir()
    binary = Path(plan["binary"]["path"])
    report_path = child_dir / f"{child_id}.report.json"
    stderr_path = child_dir / f"{child_id}.perf.stderr"
    terminal_path = child_dir / f"{child_id}.terminal.json"
    packet_path = child_dir / f"{child_id}.packet.json"
    perf_data = child_dir / f"{child_id}.perf.data"
    command = child_record_argv(plan, child_dir)
    command_path = child_dir / f"{child_id}.command.json"
    p.write_json_exclusive(command_path, {"argv": command, "environment": {**p.ENVIRONMENT, "LD_PRELOAD": "absent"}})
    started = utc()
    process_pid: int | None = None
    timed_out = False
    launch_error = False
    with report_path.open("xb") as report_handle, stderr_path.open("xb") as stderr_handle:
        try:
            proc = subprocess.Popen(command, stdout=report_handle, stderr=stderr_handle,
                                    stdin=subprocess.DEVNULL, env=native_env(),
                                    start_new_session=True, preexec_fn=pin_cpu)
            process_pid = proc.pid
            try:
                returncode = proc.wait(timeout=int(plan["execution"]["timeout_seconds"]))
            except subprocess.TimeoutExpired as exc:
                terminate_group(proc)
                returncode = proc.returncode if proc.returncode is not None else 124
                stderr_handle.write((f"timeout: {exc}; process group terminated and reaped\n").encode())
                timed_out = True
        except OSError as exc:
            returncode = 127
            stderr_handle.write((f"launch error: {exc}\n").encode())
            launch_error = True
    terminal = {"schema": "litchi.performance.0973.terminal.v2", "started_utc": started, "finished_utc": utc(), "returncode": returncode, "timed_out": timed_out, "launch_error": launch_error, "process_pid": process_pid, "command": command, "native_ld_preload": "absent"}
    p.write_json_exclusive(terminal_path, terminal)
    if returncode != 0 or timed_out or not perf_data.is_file() or not report_path.is_file():
        packet = {"schema": p.PACKET_SCHEMA, "packet_id": child_id, "status": "terminal_failure", "terminal": terminal}
        p.write_json_exclusive(packet_path, packet)
        raise p.ProtocolError(f"{child_id}: terminal failure rc={returncode}")
    report = p.load_json(report_path)
    p.validate_report(report)
    metadata = child_dir / f"{child_id}.launch.json"
    launch = p.check_launch_metadata(metadata, packet_id=child_id, binary=binary)
    if launch["exec_pid"] != launch["pid"]:
        raise p.ProtocolError(f"{child_id}: launcher PID changed")
    layout = p.derive_target_layout(perf_data, binary, expected_pid=launch["exec_pid"])
    layout_path = child_dir / f"{child_id}.target-layout.json"
    p.write_json_exclusive(layout_path, layout)
    mmap_dump_path = child_dir / f"{child_id}.perf-mmap-dump.txt"
    dump = run_process([p.PERF_PATH, "script", "-D", "-i", str(perf_data)], timeout=300)
    if dump.returncode != 0:
        raise p.ProtocolError(f"{child_id}: perf dump failed")
    p.write_text_exclusive(mmap_dump_path, dump.stdout)
    if p.sha256(mmap_dump_path) != layout["perf_mmap_dump"]["sha256"]:
        raise p.ProtocolError(f"{child_id}: MMAP2 dump changed during layout binding")
    raw_stack_summary = raw_stack.scan_dump(
        mmap_dump_path, perf_data=perf_data, binary=binary,
        expected_pid=int(launch["exec_pid"]), expected_cpu=p.CPU,
        expected_period=p.PERIOD, stack_cap=p.STACK_CAP, layout=layout,
    )
    raw_stack_path = child_dir / f"{child_id}.raw-stack.json"
    p.write_json_exclusive(raw_stack_path, raw_stack_summary)
    p.check_raw_stack_summary(raw_stack_summary, plan=plan)
    summary, decoded = decode_child(child_dir, perf_data, layout, plan, launch)
    artifacts = {
        "report": report_path, "stderr": stderr_path, "perf_data": perf_data,
        "launch": metadata, "command": command_path, "terminal": terminal_path,
        "layout": layout_path, "mmap_dump": mmap_dump_path,
        "raw_stack": raw_stack_path,
        **decoded,
    }
    gate_passed, gate = child_gate(summary["corrected"])
    packet = {
        "schema": p.PACKET_SCHEMA, "packet_id": child_id, "status": "complete",
        "child_number": child_number, "terminal": terminal, "report": report,
        "launch": launch, "layout": layout, "shim_config": {
            "target_path": str(binary), "target_name": binary.name,
            "corrected_base": layout["corrected_base"], "perf_expected_start": layout["perf_expected_start"],
        }, "perf": summary, "gate": gate, "artifacts": {name: p.file_binding(path) for name, path in artifacts.items()},
        "context": {"cell": {"case": "docfloat", "edits": 1, "retention": "default"}, "binary_sha256": plan["binary"]["sha256"], "fixture_sha256": plan["fixture"]["sha256"], "marker_calls": p.CALLS},
    }
    p.write_json_exclusive(packet_path, packet)
    return packet


def synthetic_tests() -> None:
    text = f"367/367 [031] 1.0: {p.PERIOD} instructions:u:\n\t1000 helper+0x1 (/x)\n\n"
    row = p.parse_perf_script_strict(text, expected_pid=367, expected_cpu=31)
    assert row["sample_count"] == 1 and row["empty_callchain_count"] == 0
    bad = text.replace("367/367", "367/368")
    try:
        p.parse_perf_script_strict(bad, expected_pid=367, expected_cpu=31)
    except p.ProtocolError:
        pass
    else:
        raise AssertionError("PID/TID mismatch accepted")
    unknown = f"367/367 [031] 1.0: {p.PERIOD} instructions:u:\n\t1000 [unknown] ([unknown])\n\n"
    row = p.parse_perf_script_strict(unknown, expected_pid=367, expected_cpu=31)
    assert row["physical_leaf_unknown_period"] == 0 and row["physical_leaf_unknown_samples"] == 0 and row["deepest_resolved_proxy_period"] == {}
    assert row["marker_ancestry"] == "no observed marker ancestry"
    assert native_env().get("LD_PRELOAD") is None


def self_test() -> int:
    synthetic_tests()
    try:
        reject_test_only_boundary({"synthetic_boundary": {"enabled": True}})
    except p.ProtocolError:
        pass
    else:
        raise AssertionError("synthetic boundary reached capture admission")
    raw = Path("/home/zhuhe/litchi-performance-0971/capture/instruction-profile-v2/child-01/child-01.perf.data")
    layout = p.derive_target_layout(raw, Path(p.BINARY_PATH), expected_pid=3683783)
    assert layout["corrected_base"] == "0x64d1e4e5b000"
    assert layout["perf_expected_start"] == "0x64d1e4e5c000"
    print(json.dumps({"schema": "litchi.performance.0973.self_test.v2", "status": "pass", "synthetic": "pass", "dynamic_layout": {"corrected_base": layout["corrected_base"], "perf_expected_start": layout["perf_expected_start"]}, "workload_started": False}, indent=2, sort_keys=True))
    return 0


def load_plan(path: Path) -> dict[str, Any]:
    plan = p.load_json(path.resolve())
    p.check_plan_bindings(plan, path.resolve())
    return plan


def reject_test_only_boundary(plan: dict[str, Any]) -> None:
    """Keep the explicit synthetic perf boundary out of capture admission."""
    if (plan.get("synthetic_boundary") or {}).get("enabled") is True:
        raise p.ProtocolError("synthetic_boundary is test-only and cannot be used for --run")


def run(plan_path: Path, go_path: Path) -> int:
    plan = load_plan(plan_path)
    reject_test_only_boundary(plan)
    go = p.load_json(go_path.resolve())
    p.check_go(go, plan, plan_path.resolve(), go_path.resolve())
    capture_root = Path(plan["outputs"]["capture_root"]).resolve()
    if capture_root.exists():
        raise p.ProtocolError(f"capture root already exists; refusing overwrite: {capture_root}")
    allowed = sorted(os.sched_getaffinity(0))
    if p.CPU not in allowed:
        raise p.ProtocolError(f"CPU {p.CPU} is not in runner affinity {allowed}")
    capture_root.mkdir(parents=True)
    packet_paths: list[str] = []
    failure: dict[str, Any] | None = None
    for child_number in range(1, p.CHILDREN + 1):
        try:
            packet = run_child(plan, child_number, capture_root)
            packet_paths.append(str(capture_root / f"child-{child_number:02d}/child-{child_number:02d}.packet.json"))
            corrected = packet["perf"]["corrected"]
            gate_status = "pass" if packet.get("gate", {}).get("passed") is True else "failed"
            print(json.dumps({
                "schema": "litchi.performance.0973.child_progress.v2",
                "child": child_number,
                "status": "gated",
                "gate": gate_status,
                "marker_sample_count": corrected["marker_sample_count"],
                "lost_events": corrected["lost_events"],
                "empty_callchain_count": corrected["empty_callchain_count"],
            }, sort_keys=True), flush=True)
            if gate_status != "pass":
                failure = {"reason": "per_child_marker_gate_failed", "child": child_number, "gate": packet.get("gate")}
                break
        except Exception as exc:
            failure = {"reason": "child_failed", "child": child_number, "type": type(exc).__name__, "error": str(exc)}
            print(json.dumps({"schema": "litchi.performance.0973.child_progress.v2", "child": child_number, "status": "failed", "error": str(exc)}, sort_keys=True), flush=True)
            break
    manifest = {"schema": p.MANIFEST_SCHEMA, "status": "failed" if failure else "complete", "created_utc": utc(), "plan": {"path": str(plan_path.resolve()), "sha256": p.sha256(plan_path)}, "no_retry": True, "children_requested": p.CHILDREN, "packets": packet_paths, "failure": failure}
    p.write_json_exclusive(capture_root / "capture-manifest-v2.json", manifest)
    print(json.dumps(manifest, indent=2, sort_keys=True))
    return 2 if failure else 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--check-plan", type=Path)
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--plan", type=Path)
    ap.add_argument("--go", type=Path)
    args = ap.parse_args()
    try:
        if args.self_test:
            if args.check_plan or args.run or args.plan or args.go:
                raise p.ProtocolError("--self-test is standalone")
            return self_test()
        if args.check_plan:
            if args.run or args.plan or args.go:
                raise p.ProtocolError("--check-plan is standalone")
            plan = load_plan(args.check_plan)
            print(json.dumps({"schema": p.PLAN_SCHEMA, "status": "pass", "plan": str(args.check_plan.resolve()), "sha256": p.sha256(args.check_plan.resolve())}, indent=2, sort_keys=True))
            return 0
        if args.run:
            if args.plan is None or args.go is None:
                raise p.ProtocolError("--run requires --plan and --go")
            return run(args.plan.resolve(), args.go.resolve())
        raise p.ProtocolError("choose --self-test, --check-plan, or --run")
    except Exception as exc:
        print(f"0973 runner rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
