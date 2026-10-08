#!/usr/bin/env python3
"""0974 v2 per-thread adapter layered over immutable v1/0973 decoder."""
from __future__ import annotations

import copy
import importlib.util
import json
import re
import sys
import tempfile
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0974")
V1_PATH = ROOT / "tools/protocol_adapter_0974_v1.py"
V1_SHA256 = "6c2032e2d1cc04990f6f4be0df4542d393ce55f1c1c7d45b97ea62eb5f05269e"
V1_PLAN = ROOT / "plans/plan-0974-pre-capture-v1.json"
V1_HELPER_SEAL = ROOT / "receipts/helper-seal-0974-v1.json"


def _sha256(path: Path) -> str:
    import hashlib
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


if _sha256(V1_PATH) != V1_SHA256:
    raise RuntimeError("immutable 0974 v1 adapter changed")
_spec = importlib.util.spec_from_file_location("protocol_adapter_0974_v1", V1_PATH)
if _spec is None or _spec.loader is None:
    raise RuntimeError("cannot load immutable 0974 v1 adapter")
V1 = importlib.util.module_from_spec(_spec)
sys.modules["protocol_adapter_0974_v1"] = V1
_spec.loader.exec_module(V1)

# The retained adapter's native_record_argv delegates through the mutable
# BASE_RUNNER.child_record_argv attribute.  Capture the actual retained
# function from a fresh module before installing any 0974 wrapper, so calling
# the v5 command builder can never recurse through its own patch.  The fresh
# raw/common modules likewise provide unwrapped originals for idempotent patch
# installation and cleanup when a test invokes the installer more than once.
_RETAINED_RUNNER = V1.load_module("litchi_0974_v5_retained_runner", V1.RETAINED_FILES["runner"])
_RETAINED_COMMON = V1.load_module("litchi_0974_v5_retained_common", V1.RETAINED_FILES["protocol"])
_RETAINED_RAW_STACK = V1.load_module("litchi_0974_v5_retained_raw_stack", V1.RETAINED_FILES["raw_stack"])
ORIGINAL_CHILD_RECORD_ARGV = _RETAINED_RUNNER.child_record_argv
ORIGINAL_WRITE_JSON_EXCLUSIVE = _RETAINED_COMMON.write_json_exclusive
ORIGINAL_SCAN_DUMP = _RETAINED_RAW_STACK.scan_dump
_PATCH_STATE_ATTR = "_litchi_0974_v5_patch_state"

BASE = V1.BASE
BASE_RUNNER = V1.BASE_RUNNER
BASE_REDUCER = V1.BASE_REDUCER
BASE_RAW_STACK = V1.BASE_RAW_STACK
ProtocolError = V1.ProtocolError
load_json = V1.load_json
sha256 = V1.sha256
file_binding = V1.file_binding

PLAN_SCHEMA = "litchi.performance.0974.plan.v3"
GO_SCHEMA = "litchi.performance.0974.root_go.v3"
HELPER_SCHEMA = "litchi.performance.0974.helper_seal.v3"
MEMORY_TOKENS = V1.MEMORY_TOKENS
MEMORY_BYTES = V1.MEMORY_BYTES
RAW_LOST_FIELD = V1.RAW_LOST_FIELD
TERMINAL_LOST_FIELD = V1.TERMINAL_LOST_FIELD

DESIGN_PATH = ROOT / "design-v2.md"
CAPABILITY_PATH = ROOT / "capability-perthread-mmap16m-v1/capability-receipt-v1.json"
FAILED_MMAP_PATH = ROOT / "capability-mmap16m-v1/capability-receipt-v2.json"
FAILED_CPU_PATH = ROOT / "capability-cpu31-mmap16m-v1/capability-receipt-v1.json"
RECORD_SPEC_PATH = ROOT / "analysis/native-record-spec-v2.json"
LIFECYCLE_CONTROL_PATH = ROOT / "lifecycle-negative-control-v1/lifecycle-negative-control-receipt-v1.json"


def _binding(path: str | Path) -> dict[str, Any]:
    return file_binding(path)


def native_record_argv(plan: dict[str, Any], child_dir: Path) -> list[str]:
    """Return the retained command with only --per-thread and -m 16M."""
    # Do not call V1.native_record_argv: it follows BASE_RUNNER's mutable
    # child_record_argv attribute, which points back to this function after
    # installation.
    base_argv = list(ORIGINAL_CHILD_RECORD_ARGV(plan, child_dir))
    if base_argv.count("--call-graph") != 1:
        raise ProtocolError("retained native command has no unique --call-graph")
    index = base_argv.index("--call-graph")
    if MEMORY_TOKENS != ["-m", "16M"] or base_argv.count("-m") != 0:
        raise ProtocolError("retained command already has a memory selector or v5 memory delta changed")
    base_argv[index:index] = MEMORY_TOKENS
    forbidden = ("-a", "--all-cpus", "-C", "--cpu", "--inherit")
    if "--per-thread" in base_argv or any(token in base_argv for token in forbidden):
        raise ProtocolError("retained command already contains a forbidden/per-thread token")
    record_index = base_argv.index("record")
    result = base_argv[:record_index + 1] + ["--per-thread"] + base_argv[record_index + 1:]
    if result[record_index + 1] != "--per-thread" or result.count("--per-thread") != 1:
        raise ProtocolError("--per-thread insertion drift")
    if result[result.index("--call-graph") - 2:result.index("--call-graph")] != MEMORY_TOKENS:
        raise ProtocolError("v5 memory delta insertion drift")
    return result


_COMM = re.compile(r"PERF_RECORD_COMM(?:\s+exec)?:\s*(?P<tail>.+)$")
_THREAD_MAP = re.compile(r"PERF_RECORD_THREAD_MAP\s+nr:\s*(?P<nr>\d+)\s+thread:\s*(?P<threads>.*)$")
_EXIT = re.compile(r"PERF_RECORD_EXIT\((?P<pid>\d+):(?P<tid>\d+)\):\((?P<ppid>\d+):(?P<ptid>\d+)\)")
_FORK = re.compile(r"PERF_RECORD_(?:FORK|CLONE)\((?P<child_pid>\d+):(?P<child_tid>\d+)\):\((?P<parent_pid>\d+):(?P<parent_tid>\d+)\)")


def parse_header_contract(text: str) -> dict[str, Any]:
    event_lines = [line.strip() for line in text.splitlines() if line.strip().startswith("# event : name")]
    if len(event_lines) != 1:
        raise ProtocolError("per-thread header must contain exactly one event line")
    line = event_lines[0]
    required = ("task = 1", "comm = 1", "comm_exec = 1", "enable_on_exec = 1", "sample_stack_user = 32768", "read_format = ID|LOST")
    missing = [token for token in required if token not in line]
    if missing:
        raise ProtocolError(f"per-thread header missing attributes: {missing}")
    inherit = re.search(r"\binherit\s*=\s*(\d+)", line)
    if inherit is not None and inherit.group(1) != "0":
        raise ProtocolError("per-thread header advertises inherit != 0")
    return {"status": "pass", "event_line": line, "task": 1, "inherit": int(inherit.group(1)) if inherit else 0, "inherit_field_absent": inherit is None, "required_attributes": list(required), "scope": "header attributes only; raw lifetime task stream remains authoritative"}


def parse_task_stream(text: str, *, expected_pid: int, binary_name: str) -> dict[str, Any]:
    """Validate bounded raw COMM/FORK/EXIT/THREAD_MAP task records."""
    counts: dict[str, int] = {}
    comms: list[dict[str, Any]] = []
    thread_maps: list[dict[str, Any]] = []
    cpu_maps: list[str] = []
    exits: list[dict[str, int]] = []
    errors: list[str] = []
    for line in text.splitlines():
        if "PERF_RECORD_FORK" in line or "PERF_RECORD_CLONE" in line:
            counts["FORK"] = counts.get("FORK", 0) + 1
            errors.append("unexpected_task_creation_record")
            match = _FORK.search(line)
            if match:
                errors.append(f"fork_pid={match.group('child_pid')}/{match.group('child_tid')}")
            continue
        if "PERF_RECORD_EXIT" in line:
            counts["EXIT"] = counts.get("EXIT", 0) + 1
            match = _EXIT.search(line)
            if not match:
                errors.append("malformed_exit_record")
            else:
                row = {key: int(value) for key, value in match.groupdict().items()}
                exits.append(row)
                if row["pid"] != expected_pid or row["tid"] != expected_pid:
                    errors.append("descendant_exit_record")
            continue
        if "PERF_RECORD_THREAD_MAP" in line:
            counts["THREAD_MAP"] = counts.get("THREAD_MAP", 0) + 1
            match = _THREAD_MAP.search(line)
            if not match:
                errors.append("malformed_thread_map")
            else:
                threads = [int(value) for value in re.findall(r"\d+", match.group("threads"))]
                row = {"nr": int(match.group("nr")), "threads": threads}
                thread_maps.append(row)
                if row != {"nr": 1, "threads": [expected_pid]}:
                    errors.append("thread_map_not_single_expected_pid")
            continue
        if "PERF_RECORD_CPU_MAP" in line:
            counts["CPU_MAP"] = counts.get("CPU_MAP", 0) + 1
            cpu_maps.append(line.split("PERF_RECORD_CPU_MAP:", 1)[-1].strip())
            continue
        if "PERF_RECORD_COMM" in line:
            counts["COMM"] = counts.get("COMM", 0) + 1
            match = _COMM.search(line)
            if not match or ":" not in match.group("tail"):
                errors.append("malformed_comm_record")
                continue
            comm, pid_tid = match.group("tail").rsplit(":", 1)
            if "/" not in pid_tid:
                errors.append("malformed_comm_pid_tid")
                continue
            pid_text, tid_text = pid_tid.strip().split("/", 1)
            try:
                pid, tid = int(pid_text), int(tid_text)
            except ValueError:
                errors.append("malformed_comm_pid_tid")
                continue
            comms.append({"comm": comm.strip(), "pid": pid, "tid": tid})
            if pid != expected_pid or tid != expected_pid:
                errors.append("comm_pid_tid_escaped_expected")
    required_names = {"perf-exec", "python3.14", binary_name[:15]}
    observed_names = {row["comm"] for row in comms}
    missing_names = sorted(name for name in required_names if not any(observed == name or observed.startswith(name) or name.startswith(observed) for observed in observed_names))
    if not thread_maps:
        errors.append("missing_thread_map")
    if len(thread_maps) != 1:
        errors.append("thread_map_count_not_one")
    if not comms:
        errors.append("missing_comm_records")
    if missing_names:
        errors.append("missing_launcher_or_native_comm:" + ",".join(missing_names))
    return {"status": "pass" if not errors else "fail", "expected_pid": expected_pid, "event_counts": counts, "comm_records": comms, "thread_maps": thread_maps, "cpu_maps": cpu_maps, "exit_records": exits, "main_exit_allowed": True, "missing_required_comm_names": missing_names, "errors": errors, "scope": "raw perf task/lifecycle records for target PID only; no unrelated process scan", "authority": "zero-loss raw task stream is lifetime authority; /proc snapshots are bounded observations only"}


def proc_task_snapshot(pid: int, *, phase: str) -> dict[str, Any]:
    """Take a bounded observation without turning access errors into exit."""
    root = Path(f"/proc/{pid}/task")
    common = {
        "phase": phase,
        "pid": pid,
        "scope": "target /proc/<pid>/task only",
        "lifetime_authority": "raw zero-loss task stream",
        "live_history_claim": False,
    }
    try:
        task_ids = sorted(int(entry.name) for entry in root.iterdir() if entry.name.isdecimal())
    except FileNotFoundError:
        # The packet is admitted after the child exits.  This is an explicit
        # post-exit observation, not evidence that the task was absent while
        # the workload was running.
        return {
            **common,
            "status": "absent",
            "task_ids": [],
            "reason": "process_not_present_at_post_exit_observation",
            "observation": "post_exit_observation",
        }
    except PermissionError as exc:
        # Permission errors must remain visible and fail closed; classifying
        # them as an exited process would erase a custody failure.
        return {
            **common,
            "status": "error",
            "task_ids": None,
            "reason": "permission_denied",
            "error_type": type(exc).__name__,
            "observation": "snapshot_error",
        }
    except OSError as exc:
        return {
            **common,
            "status": "error",
            "task_ids": None,
            "reason": "snapshot_os_error",
            "error_type": type(exc).__name__,
            "observation": "snapshot_error",
        }
    return {
        **common,
        "status": "present",
        "task_ids": task_ids,
        "reason": "task_directory_observed",
        "observation": "bounded_task_snapshot",
    }


def _validate_recorded_proc_snapshot(snapshot: Any, *, expected_pid: int, phase: str) -> None:
    if not isinstance(snapshot, dict):
        raise ProtocolError(f"{phase}: recorded /proc task snapshot is not an object")
    if snapshot.get("phase") != phase or snapshot.get("pid") != expected_pid:
        raise ProtocolError(f"{phase}: recorded /proc task identity mismatch")
    if snapshot.get("scope") != "target /proc/<pid>/task only" or snapshot.get("lifetime_authority") != "raw zero-loss task stream":
        raise ProtocolError(f"{phase}: recorded /proc task scope drift")
    if snapshot.get("live_history_claim") is not False:
        raise ProtocolError(f"{phase}: /proc observation makes an unsupported live-history claim")
    status = snapshot.get("status")
    task_ids = snapshot.get("task_ids")
    if status == "absent":
        if task_ids != [] or snapshot.get("observation") != "post_exit_observation":
            raise ProtocolError(f"{phase}: absent snapshot is not an explicit post-exit observation")
        return
    if status == "present":
        if task_ids != [expected_pid] or snapshot.get("observation") != "bounded_task_snapshot":
            raise ProtocolError(f"{phase}: present snapshot is not exactly the launcher task")
        return
    if status == "error":
        raise ProtocolError(f"{phase}: /proc task snapshot failed: {snapshot.get('reason')}")
    raise ProtocolError(f"{phase}: invalid recorded /proc task snapshot status")


def validate_recorded_task_audit(
    task: Any,
    *,
    dump_path: Path,
    header_path: Path,
    expected_pid: int,
    binary_name: str,
) -> None:
    """Revalidate immutable raw/header evidence without querying /proc.

    The snapshots are observations captured at packet admission.  Reopening
    ``/proc/<pid>/task`` during offline reduction would be a different event
    and could even inspect a reused PID, so only the recorded rows are checked.
    """
    if not isinstance(task, dict) or task.get("status") != "pass":
        raise ProtocolError("recorded task audit missing/failing")
    header = parse_header_contract(header_path.read_text(encoding="utf-8", errors="strict"))
    stream = parse_task_stream(dump_path.read_text(encoding="utf-8", errors="strict"), expected_pid=expected_pid, binary_name=binary_name)
    if stream["status"] != "pass":
        raise ProtocolError(f"immutable raw task/lifecycle audit failed: {stream['errors']}")
    if task.get("header") != header:
        raise ProtocolError("recorded task header changed from immutable header artifact")
    if task.get("stream") != stream:
        raise ProtocolError("recorded task stream changed from immutable raw artifact")
    snapshots = task.get("proc_snapshots")
    if not isinstance(snapshots, dict):
        raise ProtocolError("recorded task audit has no proc_snapshots")
    _validate_recorded_proc_snapshot(snapshots.get("before"), expected_pid=expected_pid, phase="before_packet_decode")
    _validate_recorded_proc_snapshot(snapshots.get("after"), expected_pid=expected_pid, phase="after_packet_decode")


def task_audit_from_files(*, dump_path: Path, header_path: Path, expected_pid: int, binary_name: str) -> dict[str, Any]:
    header = parse_header_contract(header_path.read_text(encoding="utf-8", errors="strict"))
    stream = parse_task_stream(dump_path.read_text(encoding="utf-8", errors="strict"), expected_pid=expected_pid, binary_name=binary_name)
    before = proc_task_snapshot(expected_pid, phase="before_packet_decode")
    after = proc_task_snapshot(expected_pid, phase="after_packet_decode")
    if stream["status"] != "pass":
        raise ProtocolError(f"raw task/lifecycle audit failed: {stream['errors']}")
    task = {
        "status": "pass",
        "header": header,
        "stream": stream,
        "proc_snapshots": {"before": before, "after": after},
        "absence_is_allowed": True,
        "scope": "per-thread target task only; no full /proc scan",
        "snapshot_scope": "bounded post-exit observations; not live execution history",
    }
    validate_recorded_task_audit(task, dump_path=dump_path, header_path=header_path, expected_pid=expected_pid, binary_name=binary_name)
    return task


def raw_stack_with_v2_fields(*args: Any, **kwargs: Any) -> dict[str, Any]:
    result = BASE_RAW_STACK.scan_dump(*args, **kwargs)
    dump_path = Path(args[0]) if args else Path(kwargs["dump_path"])
    result.update(V1.parse_raw_lost_sum(dump_path.read_text(encoding="utf-8", errors="strict")))
    binary = Path(kwargs["binary"] if "binary" in kwargs else args[2])
    expected_pid = int(kwargs["expected_pid"] if "expected_pid" in kwargs else args[3])
    result["task_stream"] = parse_task_stream(dump_path.read_text(encoding="utf-8", errors="strict"), expected_pid=expected_pid, binary_name=binary.name)
    if result["task_stream"]["status"] != "pass":
        raise ProtocolError(f"raw task/lifecycle audit failed: {result['task_stream']['errors']}")
    return result


def _check_helper_seal(plan: dict[str, Any]) -> None:
    row = plan.get("helper_seal")
    if not isinstance(row, dict):
        raise ProtocolError("missing v2 helper seal")
    seal = load_json(row["path"])
    if seal.get("schema") != HELPER_SCHEMA or seal.get("status") != "sealed":
        raise ProtocolError("v2 helper seal schema/status drift")
    if plan.get("helpers_seal_sha256") != row.get("sha256") or seal.get("helpers") != plan.get("helpers"):
        raise ProtocolError("v2 helper seal/plan mismatch")
    expected_names = {"protocol", "runner", "reducer", "plan_generator", "raw_stack", "shim_source", "shim_map", "shim_so"}
    if set(seal.get("helpers", {})) != expected_names:
        raise ProtocolError("v2 helper set drift")
    for name, binding in seal["helpers"].items():
        current = file_binding(binding["path"])
        if current != {key: binding[key] for key in ("path", "sha256", "bytes")}:
            raise ProtocolError(f"v2 helper changed: {name}")
    if seal.get("retained_v1_helper_seal") != file_binding(V1_HELPER_SEAL) or seal.get("retained_v1_plan") != file_binding(V1_PLAN):
        raise ProtocolError("v1 seal/plan lineage drift")


def check_plan_bindings(plan: dict[str, Any], plan_path: Path | None = None) -> None:
    if plan.get("schema") != PLAN_SCHEMA or plan.get("status") != "pre_capture":
        raise ProtocolError("v2 plan schema/status mismatch")
    old = load_json(V1_PLAN)
    view = copy.deepcopy(plan)
    for key in ("design", "helper_seal", "helpers_seal_sha256", "helpers", "outputs"):
        view[key] = old[key]
    view["schema"] = V1.PLAN_SCHEMA
    view["plan_path"] = None
    for key in ("retained_0974_design", "record_mode", "task_stream_contract", "proc_task_contract", "record_spec", "capability_lineage", "failed_mmap_probe", "failed_cpu_probe", "lifecycle_control"):
        view.pop(key, None)
    view["capability_probe"] = old["capability_probe"]
    view["raw_stack_contract"] = old["raw_stack_contract"]
    view["perf"] = {key: value for key, value in plan["perf"].items() if key not in {"record_mode", "cpu_selector"}}
    V1.check_plan_bindings(view, None)
    perf = plan.get("perf") or {}
    if perf.get("record_memory") != {"tokens": MEMORY_TOKENS, "bytes": MEMORY_BYTES, "position": "immediately before --call-graph"}:
        raise ProtocolError("v2 memory selector drift")
    if perf.get("record_mode") != {"tokens": ["--per-thread"], "inheritance": "disabled_by_per_thread", "forbidden_tokens": ["-a", "--all-cpus", "-C", "--cpu", "--inherit"], "scope": "one launcher/native exec PID and one native thread; task/FORK audit required per child"}:
        raise ProtocolError("v2 record mode drift")
    if perf.get("cpu_selector") != {"tokens": [], "affinity_scope": "outer taskset --cpu-list 31"}:
        raise ProtocolError("v2 CPU selector drift")
    if plan.get("outputs", {}).get("capture_root") != str((ROOT / "capture/instruction-profile-v1").resolve()):
        raise ProtocolError("v2 output path drift")
    for key, expected in (("design", DESIGN_PATH), ("capability_probe", CAPABILITY_PATH), ("record_spec", RECORD_SPEC_PATH), ("failed_mmap_probe", FAILED_MMAP_PATH), ("failed_cpu_probe", FAILED_CPU_PATH)):
        if plan.get(key, {}).get("path") != str(expected.absolute()) or plan[key].get("sha256") != sha256(expected):
            raise ProtocolError(f"v2 binding drift: {key}")
    if plan.get("lifecycle_control", {}).get("path") != str(LIFECYCLE_CONTROL_PATH.absolute()) or plan["lifecycle_control"].get("sha256") != sha256(LIFECYCLE_CONTROL_PATH):
        raise ProtocolError("v2 lifecycle control binding drift")
    capability = load_json(CAPABILITY_PATH)
    if capability.get("status") != "setup_succeeded_no_workload_samples" or capability.get("result", {}).get("returncode") != 0 or capability.get("scope", {}).get("native_workload_started") is not False:
        raise ProtocolError("v2 capability receipt is not a successful no-workload probe")
    spec_row = load_json(RECORD_SPEC_PATH)
    if spec_row.get("status") != "independent_final_spec" or spec_row.get("record_mode", {}).get("required_tokens") != ["--per-thread"]:
        raise ProtocolError("independent argv spec is not final")
    for path in (FAILED_MMAP_PATH, FAILED_CPU_PATH):
        failure = load_json(path)
        if failure.get("status") not in {"setup_failed_no_retry", "setup_failed"}:
            raise ProtocolError(f"historical failure receipt status drift: {path}")
    lifecycle = load_json(LIFECYCLE_CONTROL_PATH)
    if lifecycle.get("status") != "control_passed_lifecycle_exposure" or lifecycle.get("lifecycle_conclusion", {}).get("native_workload_go") is not False or lifecycle.get("lifecycle_conclusion", {}).get("raw_fork_exposed_child") is not True or lifecycle.get("lifecycle_conclusion", {}).get("raw_fork_exposed_thread") is not True:
        raise ProtocolError("v2 lifecycle negative control is not a passing rejection control")
    expected_task_contract = {"header_required": ["task = 1", "comm = 1", "comm_exec = 1", "enable_on_exec = 1", "sample_stack_user = 32768", "read_format = ID|LOST"], "inheritance": "inherit field absent or zero; inherit=1 rejects", "thread_map": "exactly nr=1 with expected launcher/native PID", "COMM": "same PID/TID only; require perf-exec, python3.14, and native binary identity", "FORK": "none", "CLONE": "none", "EXIT": "main PID/TID only; descendant exit rejects", "CPU_MAP": "record descriptively; sample CPU31 remains strict", "raw_authority": "zero-loss raw task stream is lifetime authority"}
    if plan.get("task_stream_contract") != expected_task_contract:
        raise ProtocolError("v2 task stream contract drift")
    if plan.get("proc_task_contract") != {"scope": "only /proc/<exec_pid>/task", "snapshots": ["before_packet_decode", "after_packet_decode"], "absence": "explicit post-exit absence allowed; absence does not prove history", "present": "task_ids must equal [exec_pid]", "permission_error": "explicit error and fail closed", "snapshot_scope": "bounded post-exit observations; not live execution history", "full_proc_scan": False}:
        raise ProtocolError("v2 /proc task contract drift")
    expected_raw = {"schema": BASE.RAW_STACK_SCHEMA, "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER", "period": BASE.PERIOD, "stack_cap": BASE.STACK_CAP, "bounds": "0 <= dyn_size <= captured_stack_size <= 32768", "missing_or_short": "missing fields reject; bounded short captured_stack_size is retained in its distribution", "raw_scope": "all PERF_RECORD_SAMPLE records without misc filtering", "raw_lost_field": RAW_LOST_FIELD, "raw_lost_scope": "sum of PERF_RECORD_LOST lost fields; PERF_RECORD_LOST_SAMPLES is not added", "task_stream_field": "task_stream"}
    if plan.get("raw_stack_contract") != expected_raw:
        raise ProtocolError("v2 raw contract drift")
    if plan.get("command_receipt_contract") != {"environment": {"LD_PRELOAD": "absent", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}, "cwd": "record inherited runner cwd before launch", "no_secret_environment": True}:
        raise ProtocolError("v2 command receipt contract drift")
    _check_helper_seal(plan)
    if plan_path is not None and plan.get("plan_path") != str(plan_path.resolve()):
        raise ProtocolError("v2 plan self path mismatch")


def load_plan(path: Path) -> dict[str, Any]:
    plan = load_json(path)
    check_plan_bindings(plan, path)
    return plan


def check_go(go: dict[str, Any], plan: dict[str, Any], plan_path: Path, go_path: Path) -> None:
    if go.get("schema") != GO_SCHEMA or go.get("authorized") is not True or go.get("status") != "authorized":
        raise ProtocolError("v2 root GO is not authorized")
    if go.get("plan") != file_binding(plan_path) or (go.get("quiescence") or {}).get("status") != "pass":
        raise ProtocolError("v2 GO plan/quiescence binding mismatch")
    if go.get("helper_seal") != plan.get("helper_seal") or go.get("binary") != plan.get("binary") or go.get("fixture") != plan.get("fixture"):
        raise ProtocolError("v2 GO custody mismatch")
    if go.get("record_mode") != plan.get("perf", {}).get("record_mode") or go.get("record_memory") != plan.get("perf", {}).get("record_memory"):
        raise ProtocolError("v2 GO perf mode mismatch")
    _check_helper_seal(plan)


def install_runner_patches() -> None:
    """Install v5 command, task stream, cwd, and loss checks exactly once."""
    if getattr(BASE_RUNNER, _PATCH_STATE_ATTR, None) is not None:
        # A second invocation must not wrap an already wrapped scanner/writer.
        BASE_RUNNER.child_record_argv = native_record_argv
        return
    # Restore unwrapped retained functions first.  This also makes the v5
    # installer safe in a focused test process that imported an older draft.
    BASE_RUNNER.p.write_json_exclusive = ORIGINAL_WRITE_JSON_EXCLUSIVE
    BASE_RUNNER.raw_stack.scan_dump = ORIGINAL_SCAN_DUMP
    original_write = ORIGINAL_WRITE_JSON_EXCLUSIVE
    original_scan = ORIGINAL_SCAN_DUMP
    terminal_cache: dict[Path, dict[str, Any]] = {}

    def write(path: Path, value: Any) -> None:
        resolved = Path(path).resolve()
        payload = value
        if resolved.name.endswith(".command.json"):
            payload = dict(value)
            payload["cwd"] = str(Path.cwd())
            payload["cwd_policy"] = "inherited runner cwd; recorded before launch"
        elif resolved.name.endswith(".terminal.json"):
            payload = dict(value)
            data = resolved.with_name(resolved.name.replace(".terminal.json", ".perf.data"))
            if data.is_file() and payload.get("returncode") == 0:
                payload[V1.TERMINAL_LOST_FIELD] = V1.collect_terminal_lost_samples(data)
            else:
                payload[V1.TERMINAL_LOST_FIELD] = {"record_type": "PERF_RECORD_LOST_SAMPLES", "lost_samples": None, "source": "perf report --stdio terminal summary unavailable", "argv": [BASE.PERF_PATH, "report", "--stdio", "--percent-limit", "100", "-i", str(data)]}
            terminal_cache[resolved] = payload
        elif resolved.name.endswith(".packet.json"):
            payload = dict(value)
            terminal_path = resolved.with_name(resolved.name.replace(".packet.json", ".terminal.json"))
            if terminal_path in terminal_cache:
                payload["terminal"] = terminal_cache[terminal_path]
            if payload.get("status") == "complete":
                artifacts = payload.get("artifacts") or {}
                raw_row, dump_row, launch_row = artifacts.get("raw_stack") or {}, artifacts.get("mmap_dump") or {}, artifacts.get("launch") or {}
                header_row = artifacts.get("header") or {}
                if not all(isinstance(row.get("path"), str) for row in (raw_row, dump_row, launch_row, header_row)):
                    raise ProtocolError("v2 complete packet missing task/loss artifacts")
                raw_summary = load_json(raw_row["path"])
                launch = load_json(launch_row["path"])
                task = task_audit_from_files(dump_path=Path(dump_row["path"]), header_path=Path(header_row["path"]), expected_pid=int(launch["exec_pid"]), binary_name=Path(BASE.BINARY_PATH).name)
                if raw_summary.get("task_stream") != task["stream"]:
                    raise ProtocolError("raw task stream changed between scan and packet admission")
                payload["task_audit"] = task
                V1.require_zero_loss(raw_summary, payload.get("terminal") or {})
        original_write(resolved, payload)

    def scan(*args: Any, **kwargs: Any) -> dict[str, Any]:
        result = original_scan(*args, **kwargs)
        dump_path = Path(args[0]) if args else Path(kwargs["dump_path"])
        text = dump_path.read_text(encoding="utf-8", errors="strict")
        result.update(V1.parse_raw_lost_sum(text))
        binary = Path(kwargs["binary"] if "binary" in kwargs else args[2])
        expected_pid = int(kwargs["expected_pid"] if "expected_pid" in kwargs else args[3])
        stream = parse_task_stream(text, expected_pid=expected_pid, binary_name=binary.name)
        if stream["status"] != "pass":
            raise ProtocolError(f"raw task/lifecycle audit failed: {stream['errors']}")
        result["task_stream"] = stream
        return result

    BASE_RUNNER.p.write_json_exclusive = write
    BASE_RUNNER.raw_stack.scan_dump = scan
    BASE_RUNNER.child_record_argv = native_record_argv
    setattr(BASE_RUNNER, _PATCH_STATE_ATTR, {"original_write": original_write, "original_scan": original_scan, "original_child": ORIGINAL_CHILD_RECORD_ARGV})


def clear_runner_patches() -> None:
    """Restore retained runner callables for isolated admission tests."""
    BASE_RUNNER.p.write_json_exclusive = ORIGINAL_WRITE_JSON_EXCLUSIVE
    BASE_RUNNER.raw_stack.scan_dump = ORIGINAL_SCAN_DUMP
    BASE_RUNNER.child_record_argv = ORIGINAL_CHILD_RECORD_ARGV
    if hasattr(BASE_RUNNER, _PATCH_STATE_ATTR):
        delattr(BASE_RUNNER, _PATCH_STATE_ATTR)


def self_test() -> dict[str, Any]:
    plan_stub = {"binary": {"path": BASE.BINARY_PATH}, "fixture": {"path": BASE.FIXTURE_PATH}}
    command = native_record_argv(plan_stub, ROOT / "capture" / "child-01")
    if command[command.index("record") + 1] != "--per-thread" or command[command.index("--call-graph") - 2:command.index("--call-graph")] != MEMORY_TOKENS:
        raise ProtocolError("v2 argv delta self-test failed")
    if any(token in command for token in ("-a", "--all-cpus", "-C", "--cpu", "--inherit")):
        raise ProtocolError("v2 forbidden argv token self-test failed")
    header = "# event : name = instructions:u, task = 1, comm = 1, enable_on_exec = 1, comm_exec = 1, read_format = ID|LOST, sample_stack_user = 32768\n"
    if parse_header_contract(header)["inherit"] != 0:
        raise ProtocolError("v2 header self-test failed")
    raw = """
0 0 0x10 [0x28]: PERF_RECORD_THREAD_MAP nr: 1 thread: 77
0 0 0x20 [0x10]: PERF_RECORD_CPU_MAP: -1
0 0 0x30 [0x38]: PERF_RECORD_COMM: perf-exec:77/77
0 0 0x40 [0x38]: PERF_RECORD_COMM exec: python3.14:77/77
0 0 0x50 [0x38]: PERF_RECORD_COMM exec: doc_retention_p:77/77
31 1 0x60 [0x28]: PERF_RECORD_EXIT(77:77):(76:76)
"""
    parsed = parse_task_stream(raw, expected_pid=77, binary_name="doc_retention_probe_0942_cfb_profile")
    if parsed["status"] != "pass":
        raise ProtocolError(f"v2 task self-test failed: {parsed['errors']}")
    bad = raw.replace("doc_retention_p:77/77", "doc_retention_p:78/78")
    if parse_task_stream(bad, expected_pid=77, binary_name="doc_retention_probe_0942_cfb_profile")["status"] == "pass":
        raise ProtocolError("v2 unexpected PID mutation accepted")
    try:
        parse_header_contract(header.replace("enable_on_exec = 1", "enable_on_exec = 1, inherit = 1"))
    except ProtocolError:
        pass
    else:
        raise ProtocolError("v2 inherit=1 header mutation accepted")
    forked = raw.replace("PERF_RECORD_EXIT(77:77)", "PERF_RECORD_FORK(77:77):(77:77)\nPERF_RECORD_EXIT(77:77)")
    if parse_task_stream(forked, expected_pid=77, binary_name="doc_retention_probe_0942_cfb_profile")["status"] == "pass":
        raise ProtocolError("v2 FORK lifecycle mutation accepted")
    observed = proc_task_snapshot(999999999, phase="after_packet_decode")
    _validate_recorded_proc_snapshot(observed, expected_pid=999999999, phase="after_packet_decode")
    permission = dict(observed, status="error", task_ids=None, reason="permission_denied", observation="snapshot_error", error_type="PermissionError")
    try:
        _validate_recorded_proc_snapshot(permission, expected_pid=999999999, phase="after_packet_decode")
    except ProtocolError:
        pass
    else:
        raise ProtocolError("permission-error snapshot mutation accepted")
    # Exercise the retained runner's real run_child entry point up to (but not
    # through) subprocess launch.  Popen is the only mocked operation; this
    # catches command-builder recursion that parser-only tests cannot see.
    class LaunchSentinel(Exception):
        pass

    install_runner_patches()
    install_runner_patches()  # idempotence: must not double-wrap raw scan/write
    original_popen = BASE_RUNNER.subprocess.Popen
    observed_argv: list[str] = []

    def sentinel_popen(argv: list[str], **kwargs: Any) -> None:
        observed_argv.extend(argv)
        raise LaunchSentinel("Popen intercepted before native launch")

    BASE_RUNNER.subprocess.Popen = sentinel_popen
    try:
        with tempfile.TemporaryDirectory(prefix="litchi-0974-v5-run-child-") as temp:
            capture_root = Path(temp)
            try:
                BASE_RUNNER.run_child({"binary": {"path": BASE.BINARY_PATH}, "fixture": {"path": BASE.FIXTURE_PATH}, "execution": {"timeout_seconds": 1}}, 1, capture_root)
            except LaunchSentinel:
                pass
            else:
                raise ProtocolError("run_child sentinel was not reached")
            expected = native_record_argv({"binary": {"path": BASE.BINARY_PATH}, "fixture": {"path": BASE.FIXTURE_PATH}}, capture_root / "child-01")
            if observed_argv != expected:
                raise ProtocolError("run_child generated argv differs from v5 command builder")
    finally:
        BASE_RUNNER.subprocess.Popen = original_popen
        clear_runner_patches()
    if BASE_RUNNER.child_record_argv is not ORIGINAL_CHILD_RECORD_ARGV or BASE_RUNNER.raw_stack.scan_dump is not ORIGINAL_SCAN_DUMP:
        raise ProtocolError("runner patch clear did not restore retained callables")
    return {"schema": "litchi.performance.0974.protocol_adapter_self_test.v5", "status": "pass", "workload_started": False, "record_mode": ["--per-thread"], "record_memory": MEMORY_TOKENS, "task_audit": "same-PID COMM/THREAD_MAP; no FORK/descendant EXIT; recorded post-exit snapshots only"}


if __name__ == "__main__":
    print(json.dumps(self_test(), indent=2, sort_keys=True))
