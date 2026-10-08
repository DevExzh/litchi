#!/usr/bin/env python3
"""0974 v4 offline reducer: retained 0973 arithmetic plus task-scope gates."""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

import protocol_adapter_0974_v3 as a


def _load_packets(plan_path: Path, manifest_path: Path, plan: dict) -> list[dict]:
    packets = a.BASE_REDUCER._load_packets(plan_path, manifest_path, plan)
    if len(packets) != a.BASE.CHILDREN:
        raise a.ProtocolError(f"v4 reducer requires exactly {a.BASE.CHILDREN} children")
    for packet in packets:
        raw = packet.get("_raw_stack") or {}
        terminal = packet.get("terminal") or {}
        a.V1.require_zero_loss(raw, terminal)
        task = packet.get("task_audit")
        artifacts = packet.get("_artifacts") or {}
        launch = packet.get("launch") or {}
        if not isinstance(task, dict) or task.get("status") != "pass":
            raise a.ProtocolError(f"{packet.get('packet_id')}: task audit missing/failing")
        if task.get("stream") != raw.get("task_stream"):
            raise a.ProtocolError(f"{packet.get('packet_id')}: packet/raw task stream mismatch")
        # Revalidate only immutable header/raw files.  Do not call
        # task_audit_from_files here: its /proc observation belongs to packet
        # admission, and a later PID reuse would make an offline reduction
        # nondeterministic.
        a.validate_recorded_task_audit(
            task,
            dump_path=artifacts["mmap_dump"],
            header_path=artifacts["header"],
            expected_pid=int(launch["exec_pid"]),
            binary_name=Path(plan["binary"]["path"]).name,
        )
        packet_path = Path(packet.get("_packet_path", ""))
        if not packet_path.is_absolute() or packet_path.name == "":
            raise a.ProtocolError(f"{packet.get('packet_id')}: packet path custody missing")
        command = a.load_json(artifacts["command"])
        expected_argv = a.native_record_argv(plan, packet_path.parent)
        if command.get("argv") != expected_argv or terminal.get("command") != expected_argv:
            raise a.ProtocolError(f"{packet.get('packet_id')}: recorded native argv does not match bound per-thread/memory command")
        expected_environment = {**a.BASE.ENVIRONMENT, "LD_PRELOAD": "absent"}
        if command.get("environment") != expected_environment or terminal.get("native_ld_preload") != "absent":
            raise a.ProtocolError(f"{packet.get('packet_id')}: native environment/LD_PRELOAD custody mismatch")
        if not isinstance(command.get("cwd"), str) or not Path(command["cwd"]).is_absolute() or command.get("cwd_policy") != "inherited runner cwd; recorded before launch":
            raise a.ProtocolError(f"{packet.get('packet_id')}: command cwd binding missing or malformed")
        if packet.get("gate", {}).get("passed") is not True:
            raise a.ProtocolError(f"{packet.get('packet_id')}: child gate is not passed")
    return packets


def reduce_packets(plan_path: Path, manifest_path: Path, plan: dict, packets: list[dict]) -> dict:
    result = a.BASE_REDUCER._reduce_packets(plan_path, manifest_path, plan, packets)
    task_by_id = {packet["packet_id"]: packet["task_audit"] for packet in packets}
    for row in result.get("children", []):
        if row.get("packet_id") in task_by_id:
            row["task_audit"] = task_by_id[row["packet_id"]]
    result["schema"] = "litchi.performance.0974.instruction_profile_analysis.v4"
    result["batch"] = {
        "record_mode": {"tokens": ["--per-thread"], "inheritance": "disabled_by_per_thread", "outer_affinity": 31},
        "record_memory": {"tokens": a.MEMORY_TOKENS, "bytes": a.MEMORY_BYTES, "position": "immediately before --call-graph"},
        "task_scope": "same launcher/native PID and one thread; raw task stream gate required per child",
        "retained_v1_decoder": {"plan": a.V1.file_binding(a.V1_PLAN), "helper_seal": a.V1.file_binding(a.V1_HELPER_SEAL)},
        "raw_loss_field": a.RAW_LOST_FIELD,
        "terminal_loss_field": a.TERMINAL_LOST_FIELD,
        "loss_policy": "raw PERF_RECORD_LOST sum and terminal PERF_RECORD_LOST_SAMPLES summary are separate zero gates; never added",
    }
    result["interpretation"] = result.get("interpretation", "") + " 0974 v3 adds per-thread perf scope, immutable raw task/lifecycle revalidation, and recorded post-exit snapshot custody; it remains diagnostic-only."
    return result


def self_test() -> int:
    raw = {a.RAW_LOST_FIELD: 0}
    terminal = {a.TERMINAL_LOST_FIELD: {"record_type": "PERF_RECORD_LOST_SAMPLES", "lost_samples": 0}}
    a.V1.require_zero_loss(raw, terminal)
    stream = a.parse_task_stream("""0 0 0x10 [0x28]: PERF_RECORD_THREAD_MAP nr: 1 thread: 77
0 0 0x20 [0x10]: PERF_RECORD_CPU_MAP: -1
0 0 0x30 [0x38]: PERF_RECORD_COMM: perf-exec:77/77
0 0 0x40 [0x38]: PERF_RECORD_COMM exec: python3.14:77/77
0 0 0x50 [0x38]: PERF_RECORD_COMM exec: doc_retention_p:77/77
""", expected_pid=77, binary_name="doc_retention_probe_0942_cfb_profile")
    if stream["status"] != "pass":
        raise a.ProtocolError("v4 reducer task self-test failed")
    print(json.dumps({"schema": "litchi.performance.0974.reducer_self_test.v4", "status": "pass", "children": a.BASE.CHILDREN, "workload_started": False}, indent=2, sort_keys=True))
    return 0


def admission_self_test(plan_path: Path) -> int:
    """Exercise the v4 packet-admission gates with twelve synthetic packets.

    This is an offline checker test only.  The retained base loader is replaced
    with an in-memory packet source so no workload or capture artifact is
    launched; the v4 admission function itself still performs task, loss,
    command, environment, cwd, and exact-argv checks.
    """
    plan = a.load_plan(plan_path.resolve())
    raw_text = """0 0 0x10 [0x28]: PERF_RECORD_THREAD_MAP nr: 1 thread: 77
0 0 0x20 [0x10]: PERF_RECORD_CPU_MAP: -1
0 0 0x30 [0x38]: PERF_RECORD_COMM: perf-exec:77/77
0 0 0x40 [0x38]: PERF_RECORD_COMM exec: python3.14:77/77
0 0 0x50 [0x38]: PERF_RECORD_COMM exec: doc_retention_p:77/77
"""
    header_text = "# event : name = instructions:u, task = 1, comm = 1, enable_on_exec = 1, comm_exec = 1, read_format = ID|LOST, sample_stack_user = 32768\n"
    original_loader = a.BASE_REDUCER._load_packets
    try:
        with tempfile.TemporaryDirectory(prefix="litchi-0974-admission-") as temp:
            root = Path(temp)
            packets: list[dict] = []
            for index in range(1, a.BASE.CHILDREN + 1):
                child_id = f"child-{index:02d}"
                child = root / child_id
                child.mkdir()
                pid = 70000 + index
                dump = raw_text.replace("77/77", f"{pid}/{pid}").replace("thread: 77", f"thread: {pid}")
                header_path = child / "header.txt"
                dump_path = child / "dump.txt"
                command_path = child / "command.json"
                packet_path = child / "packet.json"
                header_path.write_text(header_text, encoding="utf-8")
                dump_path.write_text(dump, encoding="utf-8")
                stream = a.parse_task_stream(dump, expected_pid=pid, binary_name=Path(plan["binary"]["path"]).name)
                if stream["status"] != "pass":
                    raise a.ProtocolError(f"synthetic stream construction failed: {stream['errors']}")
                absent = {
                    "phase": "before_packet_decode",
                    "pid": pid,
                    "scope": "target /proc/<pid>/task only",
                    "lifetime_authority": "raw zero-loss task stream",
                    "live_history_claim": False,
                    "status": "absent",
                    "task_ids": [],
                    "reason": "process_not_present_at_post_exit_observation",
                    "observation": "post_exit_observation",
                }
                task = {
                    "status": "pass",
                    "header": a.parse_header_contract(header_text),
                    "stream": stream,
                    "proc_snapshots": {"before": absent, "after": {**absent, "phase": "after_packet_decode"}},
                    "absence_is_allowed": True,
                    "scope": "per-thread target task only; no full /proc scan",
                    "snapshot_scope": "bounded post-exit observations; not live execution history",
                }
                expected = a.native_record_argv(plan, child)
                command = {"argv": expected, "environment": {**a.BASE.ENVIRONMENT, "LD_PRELOAD": "absent"}, "cwd": str(root), "cwd_policy": "inherited runner cwd; recorded before launch"}
                a.BASE.write_json_exclusive(command_path, command)
                terminal = {"command": expected, "native_ld_preload": "absent", a.TERMINAL_LOST_FIELD: {"record_type": "PERF_RECORD_LOST_SAMPLES", "lost_samples": 0}}
                packets.append({
                    "packet_id": child_id,
                    "_packet_path": str(packet_path),
                    "_artifacts": {"command": command_path, "mmap_dump": dump_path, "header": header_path},
                    "_raw_stack": {a.RAW_LOST_FIELD: 0, "task_stream": stream},
                    "terminal": terminal,
                    "task_audit": task,
                    "launch": {"exec_pid": pid},
                    "gate": {"passed": True},
                })

            def fake_loader(_plan_path: Path, _manifest_path: Path, _plan: dict) -> list[dict]:
                return packets

            a.BASE_REDUCER._load_packets = fake_loader
            _load_packets(plan_path.resolve(), root / "manifest.json", plan)
            mutated = dict(packets[0])
            mutated["_artifacts"] = dict(packets[0]["_artifacts"])
            bad_command_path = root / "bad-command.json"
            bad_command = json.loads(command_path.read_text(encoding="utf-8"))
            bad_command["argv"] = [token for token in bad_command["argv"] if token != "--per-thread"]
            a.BASE.write_json_exclusive(bad_command_path, bad_command)
            mutated["_artifacts"]["command"] = bad_command_path
            packets[0] = mutated
            try:
                _load_packets(plan_path.resolve(), root / "manifest.json", plan)
            except a.ProtocolError:
                pass
            else:
                raise a.ProtocolError("synthetic missing --per-thread command mutation was accepted")
    finally:
        a.BASE_REDUCER._load_packets = original_loader
    print(json.dumps({"schema": "litchi.performance.0974.reducer_admission_self_test.v4", "status": "pass", "children": a.BASE.CHILDREN, "workload_started": False, "checks": ["raw/header reparse", "recorded post-exit snapshots", "exact native argv", "environment/cwd custody", "separate loss gates", "missing --per-thread rejection"]}, indent=2, sort_keys=True))
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--admission-self-test", action="store_true")
    ap.add_argument("--plan", type=Path)
    ap.add_argument("--manifest", type=Path)
    ap.add_argument("--output", type=Path)
    args = ap.parse_args()
    try:
        if args.self_test:
            if args.admission_self_test or args.plan or args.manifest or args.output:
                raise a.ProtocolError("--self-test is standalone")
            return self_test()
        if args.admission_self_test:
            if args.plan is None or args.manifest or args.output:
                raise a.ProtocolError("--admission-self-test requires only --plan")
            return admission_self_test(args.plan)
        if args.plan is None or args.manifest is None or args.output is None:
            raise a.ProtocolError("--plan, --manifest, and --output are required")
        plan_path = args.plan.resolve()
        manifest_path = args.manifest.resolve()
        plan = a.load_plan(plan_path)
        packets = _load_packets(plan_path, manifest_path, plan)
        result = reduce_packets(plan_path, manifest_path, plan, packets)
        a.BASE.write_json_exclusive(args.output.resolve(), result)
        print(json.dumps({"status": "pass", "output": str(args.output.resolve()), "sha256": a.sha256(args.output.resolve())}, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0974 v4 reducer rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
