#!/usr/bin/env python3
"""Serial 0971 instruction-profile runner.

The only executable workload admitted by ``--self-test`` is ``/usr/bin/true``.
``--run`` requires an explicit plan and root GO and is intentionally fail
closed.  A failed child is recorded once and is never retried.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import signal
import subprocess
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Any

import protocol_common_0971_v2 as p


ROOT = Path(__file__).resolve().parents[1]
LAUNCHER = Path("/home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py")
PYTHON = Path("/usr/bin/python3.14")
PERF = Path("/usr/bin/perf")
TASKSET = Path("/usr/bin/taskset")


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def pin_cpu() -> None:
    os.sched_setaffinity(0, {p.CPU})


def env() -> dict[str, str]:
    result = os.environ.copy()
    result.update({"LC_ALL": "C", "LANG": "C", "TZ": "UTC"})
    return result


def run_process(argv: list[str], *, stdout: Any = None, stderr: Any = None, timeout: int | None = None, pin: bool = False) -> subprocess.CompletedProcess[str]:
    preexec = pin_cpu if pin else None
    if stdout is None:
        stdout = subprocess.PIPE
    if stderr is None:
        stderr = subprocess.PIPE
    return subprocess.run(argv, stdout=stdout, stderr=stderr, stdin=subprocess.DEVNULL, text=True, env=env(), timeout=timeout, preexec_fn=preexec, start_new_session=pin, check=False)


def terminate_group(proc: subprocess.Popen[Any], *, grace_seconds: float = 1.0) -> None:
    """Terminate and reap a timed-out child process group exactly once."""
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


def shell_result(path: Path, argv: list[str], cp: subprocess.CompletedProcess[str]) -> dict[str, Any]:
    return {"path": str(path), "argv": argv, "returncode": cp.returncode, "stdout": cp.stdout, "stderr": cp.stderr}


def synthetic_parser_tests() -> None:
    marker = p.MARKER
    text = """367/367 [031] 1.000: 50 instructions:u:\n\t1000 helper+0x1 (/x/bin)\n\t2000 %s+0x2 (/x/bin)\n\n367/367 [031] 2.000: 50 instructions:u:\n\t3000 [unknown] ([unknown])\n\t4000 parent+0x1 (/x/bin)\n\nLost samples: 0\n""" % marker
    parsed = p.parse_perf_script(text, marker_dso="/x/bin", expected_pid=367, expected_cpu=31)
    assert parsed["sample_count"] == 2 and parsed["marker_sample_count"] == 1
    assert parsed["marker_period"] == 50 and parsed["lost_events"] == 0
    assert parsed["unresolved_frames"] == 1
    assert p.parse_perf_script(text.replace("\n\n367/367", "\n...\n\n367/367"), marker_dso="/x/bin", expected_pid=367, expected_cpu=31)["stack_truncated"] == 1
    try:
        p.parse_perf_script(text, marker_dso="other-binary", expected_pid=367, expected_cpu=31)
    except p.ProtocolError:
        pass
    else:
        raise AssertionError("marker in an unexpected DSO was accepted")
    constant = p.parse_perf_script("1/1 [031] 1.0: 1 instructions:u:\n\t1000 %s (/x/bin)\n" % marker, marker_dso="/x/bin", expected_pid=1, expected_cpu=31)
    assert constant["marker_period_share"] == 1.0
    for bad in ("garbage", "1/1 [031] 1.0: xx instructions:u:\n"):
        try:
            p.parse_perf_script(bad)
        except p.ProtocolError:
            pass
        else:
            raise AssertionError("malformed perf text was accepted")
    for bad in (
        "1/1 [031] 1.0: 0 instructions:u:\n\t1000 x (/x/bin)\n",
        "2/1 [031] 1.0: 1 instructions:u:\n\t1000 x (/x/bin)\n",
        "1/1 [030] 1.0: 1 instructions:u:\n\t1000 x (/x/bin)\n",
    ):
        try:
            p.parse_perf_script(bad, marker_dso="/x/bin", expected_pid=1, expected_cpu=31)
        except p.ProtocolError:
            pass
        else:
            raise AssertionError("invalid perf sample identity was accepted")


def timeout_group_smoke() -> dict[str, Any]:
    """Verify a timeout kills/reaps a harmless shell descendant."""
    proc = subprocess.Popen(["/bin/sh", "-c", "sleep 30"], stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True, env=env())
    terminate_group(proc, grace_seconds=0.5)
    if proc.poll() is None:
        raise AssertionError("timeout group was not reaped")
    return {"returncode": proc.returncode, "descendant": "sleep 30", "group_terminated": True}


def wrapper_smoke() -> dict[str, Any]:
    """Run the exact wrapper shape around true; this is not a native workload."""
    with tempfile.TemporaryDirectory(prefix="litchi-0971-wrapper-smoke-") as raw:
        root = Path(raw)
        data = root / "true.perf.data"
        metadata = root / "true.launch.json"
        argv = [
            str(TASKSET), "--cpu-list", str(p.CPU), str(PERF), "record", "-q",
            "-e", "instructions:u", "-c", str(p.PERIOD), "--call-graph", "dwarf,16384",
            "--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache",
            "--user-callchains", "-o", str(data), "--", str(PYTHON), str(LAUNCHER),
            "--metadata", str(metadata), "--requested-affinity", "[31]", "--", "/usr/bin/true",
        ]
        cp = run_process(argv, timeout=120, pin=True)
        if cp.returncode != 0 or not data.is_file() or not metadata.is_file():
            raise p.ProtocolError(f"wrapper smoke failed: rc={cp.returncode}, stderr={cp.stderr!r}")
        launch = p.check_launch_metadata(metadata, packet_id="true", binary=Path("/usr/bin/true"))
        decoded = run_process([str(PERF), "script", "--show-lost-events", "--no-inline", "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso", "-i", str(data)], timeout=120)
        if decoded.returncode != 0:
            raise p.ProtocolError(f"perf script smoke failed: {decoded.stderr!r}")
        parsed = p.parse_perf_script(decoded.stdout)
        inline = run_process([str(PERF), "script", "--show-lost-events", "--inline", "-i", str(data)], timeout=120)
        if inline.returncode != 0:
            raise p.ProtocolError(f"inline perf script smoke failed: {inline.stderr!r}")
        buildid = run_process([str(PERF), "buildid-list", "-i", str(data)], timeout=120)
        if buildid.returncode != 0 or not buildid.stdout.strip():
            raise p.ProtocolError(f"buildid-list smoke failed: {buildid.stderr!r}")
        return {"record_argv": argv, "record_returncode": cp.returncode, "launch": launch, "decoded": {k: parsed[k] for k in ("sample_count", "total_period", "lost_events")}, "inline_returncode": inline.returncode, "buildid_returncode": buildid.returncode, "buildid_lines": len(buildid.stdout.splitlines())}


def self_test() -> int:
    synthetic_parser_tests()
    smoke = wrapper_smoke()
    timeout = timeout_group_smoke()
    print(json.dumps({"schema": "litchi.performance.0971.self_test.v2", "status": "pass", "synthetic": "pass", "wrapper": smoke, "timeout_group": timeout}, indent=2, sort_keys=True))
    return 0


def load_plan(path: Path) -> dict[str, Any]:
    plan = p.load_json(path.resolve())
    p.check_plan_bindings(plan, path.resolve())
    return plan


def command_for_child(plan: dict[str, Any], child_dir: Path) -> list[str]:
    binary = Path(plan["binary"]["path"])
    fixture = Path(plan["fixture"]["path"])
    metadata = child_dir / f"{child_dir.name}.launch.json"
    perf_data = child_dir / f"{child_dir.name}.perf.data"
    report = child_dir / f"{child_dir.name}.report.json"
    return [
        str(TASKSET), "--cpu-list", str(p.CPU), str(PERF), "record",
        "-e", "instructions:u", "-c", str(p.PERIOD), "--call-graph", "dwarf,16384",
        "--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache",
        "--user-callchains", "-o", str(perf_data), "--", str(PYTHON), str(LAUNCHER),
        "--metadata", str(metadata), "--requested-affinity", "[31]", "--", str(binary),
        "--case", "docfloat", "--input", str(fixture), "--edits", "1", "--retention", "default",
        "--warmup", str(p.WARMUPS), "--samples", str(p.SAMPLES),
    ]


def decode_child(child_dir: Path, perf_data: Path, marker_dso: str, expected_pid: int) -> tuple[dict[str, Any], dict[str, Any]]:
    no_inline_path = child_dir / f"{child_dir.name}.perf-script-no-inline.txt"
    inline_path = child_dir / f"{child_dir.name}.perf-script-inline.txt"
    no_inline = run_process([str(PERF), "script", "--show-lost-events", "--no-inline", "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso", "-i", str(perf_data)], timeout=300)
    if no_inline.returncode != 0:
        raise p.ProtocolError(f"perf no-inline decode failed: {no_inline.stderr!r}")
    p.write_text_exclusive(no_inline_path, no_inline.stdout)
    inline = run_process([str(PERF), "script", "--show-lost-events", "--inline", "-i", str(perf_data)], timeout=300)
    if inline.returncode != 0:
        raise p.ProtocolError(f"perf inline decode failed: {inline.stderr!r}")
    p.write_text_exclusive(inline_path, inline.stdout)
    summary = p.parse_perf_script(no_inline.stdout, marker_dso=marker_dso, expected_pid=expected_pid, expected_cpu=p.CPU)
    header = run_process([str(PERF), "report", "--header-only", "--stdio", "-i", str(perf_data)], timeout=300)
    if header.returncode != 0:
        raise p.ProtocolError(f"perf header decode failed: {header.stderr!r}")
    header_path = child_dir / f"{child_dir.name}.perf-header.txt"
    p.write_text_exclusive(header_path, header.stdout)
    buildid = run_process([str(PERF), "buildid-list", "-i", str(perf_data)], timeout=300)
    if buildid.returncode != 0:
        raise p.ProtocolError(f"perf buildid-list failed: {buildid.stderr!r}")
    buildid_path = child_dir / f"{child_dir.name}.perf-buildid.txt"
    p.write_text_exclusive(buildid_path, buildid.stdout)
    p.check_buildid_export(buildid.stdout, binary_path=str(Path(marker_dso).absolute()), build_id=p.BINARY_BUILD_ID)
    return summary, {"no_inline": str(no_inline_path), "inline": str(inline_path), "header": str(header_path), "buildid": str(buildid_path)}


def run_child(plan: dict[str, Any], child_number: int, root: Path) -> dict[str, Any]:
    child_id = f"child-{child_number:02d}"
    child_dir = root / child_id
    child_dir.mkdir()
    report_path = child_dir / f"{child_id}.report.json"
    stderr_path = child_dir / f"{child_id}.perf.stderr"
    terminal_path = child_dir / f"{child_id}.terminal.json"
    packet_path = child_dir / f"{child_id}.packet.json"
    perf_data = child_dir / f"{child_id}.perf.data"
    command = command_for_child(plan, child_dir)
    command_path = child_dir / f"{child_id}.command.json"
    p.write_json_exclusive(command_path, {"argv": command, "environment": {"LC_ALL": "C", "LANG": "C", "TZ": "UTC"}})
    process_pid: int | None = None
    with report_path.open("xb") as report_handle, stderr_path.open("xb") as stderr_handle:
        started = utc()
        timed_out = False
        launch_error = False
        try:
            proc = subprocess.Popen(command, stdout=report_handle, stderr=stderr_handle, stdin=subprocess.DEVNULL, env=env(), start_new_session=True, preexec_fn=pin_cpu)
            process_pid = proc.pid
            try:
                returncode = proc.wait(timeout=int(plan["execution"]["timeout_seconds"]))
            except subprocess.TimeoutExpired as exc:
                terminate_group(proc)
                returncode = proc.returncode if proc.returncode is not None else 124
                stderr_handle.write((f"timeout: {exc}; process group terminated and reaped\n").encode())
                timed_out = True
            cp = subprocess.CompletedProcess(command, returncode=returncode)
        except OSError as exc:
            cp = subprocess.CompletedProcess(command, returncode=127)
            stderr_handle.write((f"launch error: {exc}\n").encode())
            launch_error = True
        except subprocess.TimeoutExpired as exc:
            cp = subprocess.CompletedProcess(command, returncode=124)
            timed_out = True
            stderr_handle.write((f"timeout: {exc}\n").encode())
    terminal = {"schema": "litchi.performance.0971.terminal.v2", "started_utc": started, "finished_utc": utc(), "returncode": cp.returncode, "timed_out": timed_out, "launch_error": launch_error, "process_pid": process_pid, "command": command}
    p.write_json_exclusive(terminal_path, terminal)
    if cp.returncode != 0 or timed_out or not perf_data.is_file() or not report_path.is_file():
        packet = {"schema": p.PACKET_SCHEMA, "packet_id": child_id, "status": "terminal_failure", "terminal": terminal, "paths": {"report": str(report_path), "stderr": str(stderr_path), "perf_data": str(perf_data)}}
        p.write_json_exclusive(packet_path, packet)
        raise p.ProtocolError(f"{child_id}: terminal failure rc={cp.returncode}")
    report = p.load_json(report_path)
    p.validate_report(report)
    metadata = child_dir / f"{child_id}.launch.json"
    launch = p.check_launch_metadata(metadata, packet_id=child_id, binary=Path(plan["binary"]["path"]))
    summary, decoded = decode_child(child_dir, perf_data, str(Path(plan["binary"]["path"]).absolute()), launch["exec_pid"])
    artifacts = {
        "report": p.file_binding(report_path),
        "stderr": p.file_binding(stderr_path),
        "perf_data": p.file_binding(perf_data),
        "launch": p.file_binding(metadata),
        "command": p.file_binding(command_path),
        "terminal": p.file_binding(terminal_path),
        "no_inline": p.file_binding(Path(decoded["no_inline"])),
        "inline": p.file_binding(Path(decoded["inline"])),
        "header": p.file_binding(Path(decoded["header"])),
        "buildid": p.file_binding(Path(decoded["buildid"])),
    }
    packet = {
        "schema": p.PACKET_SCHEMA,
        "packet_id": child_id,
        "status": "complete",
        "child_number": child_number,
        "terminal": terminal,
        "context": {"cell": {"case": "docfloat", "edits": 1, "retention": "default"}, "binary_sha256": plan["binary"]["sha256"], "fixture_sha256": plan["fixture"]["sha256"], "source_manifest_sha256": plan["source_manifest"]["sha256"], "marker_calls": p.CALLS},
        "report": report,
        "launch": launch,
        "perf": summary,
        "decoded": decoded,
        "artifacts": artifacts,
    }
    p.write_json_exclusive(packet_path, packet)
    return packet


def run(plan_path: Path, go_path: Path) -> int:
    plan = load_plan(plan_path)
    go = p.load_json(go_path.resolve())
    p.check_go(go, plan, plan_path.resolve(), go_path.resolve())
    root = Path(plan["outputs"]["capture_root"]).resolve()
    if root.exists():
        raise p.ProtocolError(f"capture root already exists; refusing overwrite: {root}")
    allowed = sorted(os.sched_getaffinity(0))
    if p.CPU not in allowed:
        raise p.ProtocolError(f"CPU {p.CPU} is not in runner affinity {allowed}")
    root.mkdir(parents=True)
    packets: list[str] = []
    failure: dict[str, Any] | None = None
    for child_number in range(1, p.CHILDREN + 1):
        try:
            packet = run_child(plan, child_number, root)
            packets.append(str(root / f"child-{child_number:02d}/child-{child_number:02d}.packet.json"))
            if child_number == 1:
                pilot = packet["perf"]
                if pilot["marker_sample_count"] < 100 or pilot["lost_events"] != 0 or not isinstance(pilot.get("marker_period_share"), (int, float)) or pilot["marker_period_share"] <= 0:
                    failure = {"reason": "pilot_gate_failed", "pilot": {k: pilot[k] for k in ("marker_sample_count", "lost_events", "marker_period_share")}}
                    break
        except Exception as exc:
            failure = {"reason": "child_failed", "child": child_number, "error": str(exc), "type": type(exc).__name__}
            break
    manifest = {"schema": "litchi.performance.0971.capture_manifest.v2", "status": "failed" if failure else "complete", "created_utc": utc(), "plan": {"path": str(plan_path.resolve()), "sha256": p.sha256(plan_path)}, "no_retry": True, "children_requested": p.CHILDREN, "packets": packets, "failure": failure}
    p.write_json_exclusive(root / "capture-manifest-v2.json", manifest)
    if failure:
        print(json.dumps(manifest, indent=2, sort_keys=True), file=sys.stderr)
        return 2
    print(json.dumps(manifest, indent=2, sort_keys=True))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--check-plan", type=Path)
    parser.add_argument("--run", action="store_true")
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--go", type=Path)
    args = parser.parse_args()
    try:
        if args.self_test:
            if any(value is not None for value in (args.check_plan, args.plan, args.go)) or args.run:
                raise p.ProtocolError("--self-test is standalone")
            return self_test()
        if args.check_plan is not None:
            if args.run or args.plan or args.go:
                raise p.ProtocolError("--check-plan is standalone")
            plan = load_plan(args.check_plan)
            print(json.dumps({"status": "pass", "schema": p.PLAN_SCHEMA, "plan": str(args.check_plan.resolve()), "sha256": p.sha256(args.check_plan.resolve())}, indent=2, sort_keys=True))
            return 0
        if args.run:
            if args.plan is None or args.go is None:
                raise p.ProtocolError("--run requires --plan and --go")
            return run(args.plan.resolve(), args.go.resolve())
        raise p.ProtocolError("choose --self-test, --check-plan, or --run")
    except Exception as exc:
        print(f"0971 runner rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
