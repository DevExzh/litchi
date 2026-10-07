#!/usr/bin/env python3
"""Fail-closed, serial runner for the prospective 0956 native width study.

This file is deliberately external to the repository.  It only runs when an
immutable plan is marked ``authorized``, ``--execute`` is supplied, and the
caller sets ``LITCHI_PERF_0956_CAPTURE_APPROVED=1``.  Every child is fresh and
is launched once.  The process terminal is sealed before any report parsing.
Observation delays are polled; only an actual deadline or terminal/validation
failure stops the schedule.  There is no retry or replacement path.
"""
from __future__ import annotations

import argparse
import datetime as dt
import errno
import hashlib
import json
import math
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from typing import Any, Iterable

SCHEMA = "litchi.performance.0956.native-width-plan.v2"
RUNNER_SCHEMA = "litchi.performance.0956.native-width-capture.v2"
REPORT_SCHEMA = "litchi.xlsx-managed-compression-benchmark.v1"
APPROVAL_ENV = "LITCHI_PERF_0956_CAPTURE_APPROVED"
REPORT_LIMIT = 8 * 1024 * 1024
TIME_FIELDS = {
    "Maximum resident set size (kbytes)": "max_rss_kib",
    "User time (seconds)": "user_seconds",
    "System time (seconds)": "system_seconds",
    "Major (requiring I/O) page faults": "major_page_faults",
    "Minor (reclaiming a frame) page faults": "minor_page_faults",
    "Voluntary context switches": "voluntary_context_switches",
    "Involuntary context switches": "involuntary_context_switches",
}


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def descriptor(path: Path, *, hash_file: bool = True) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise ValueError(f"expected regular file: {path}")
    item: dict[str, Any] = {"path": str(path.resolve()), "bytes": path.stat().st_size}
    if hash_file:
        item["sha256"] = sha256(path)
    return item


def write_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def read_json(path: Path) -> dict[str, Any]:
    with path.open("rb") as stream:
        value = json.loads(stream.read().decode("utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"JSON object required: {path}")
    return value


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def int_value(value: Any, location: str, *, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{location} must be integer >= {minimum}")
    return value


def exact_case_ids() -> tuple[dict[str, Any], ...]:
    common_medium = {
        "source_bytes": 4226429,
        "source_sha256": "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036",
    }
    common_dense = {
        "source_bytes": 4251863,
        "source_sha256": "893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a",
    }
    medium_limits = {
        "1": {"memory": 2956761, "input_bytes": 67108864, "output_bytes": 8519112, "objects": 1000000, "depth": 256, "work": 1000000000, "execution_workers": 1, "execution_in_flight_tasks": 1, "execution_in_flight_bytes": 2956761, "execution_min_parallel_bytes": 0, "execution_min_task_bytes": 0, "root_workers": 1, "root_io_concurrency": 1, "root_cpu_tasks": 5},
        "2": {"memory": 2956761, "input_bytes": 67108864, "output_bytes": 8519112, "objects": 1000000, "depth": 256, "work": 1000000000, "execution_workers": 2, "execution_in_flight_tasks": 2, "execution_in_flight_bytes": 2956761, "execution_min_parallel_bytes": 0, "execution_min_task_bytes": 0, "root_workers": 2, "root_io_concurrency": 1, "root_cpu_tasks": 5},
        "4": {"memory": 2956761, "input_bytes": 67108864, "output_bytes": 8519112, "objects": 1000000, "depth": 256, "work": 1000000000, "execution_workers": 4, "execution_in_flight_tasks": 4, "execution_in_flight_bytes": 2956761, "execution_min_parallel_bytes": 0, "execution_min_task_bytes": 0, "root_workers": 4, "root_io_concurrency": 1, "root_cpu_tasks": 5},
    }
    dense_limits = {
        str(width): {"memory": 3196436, "input_bytes": 67108864, "output_bytes": 8570408, "objects": 1000000, "depth": 256, "work": 1000000000, "execution_workers": width, "execution_in_flight_tasks": width, "execution_in_flight_bytes": 3196436, "execution_min_parallel_bytes": 0, "execution_min_task_bytes": 0, "root_workers": width, "root_io_concurrency": 1, "root_cpu_tasks": 5}
        for width in (1, 2, 4)
    }
    edit_limits = {
        str(width): {"memory": 2751297, "input_bytes": 67108864, "output_bytes": 8518496, "objects": 1000000, "depth": 256, "work": 1000000000, "execution_workers": width, "execution_in_flight_tasks": width, "execution_in_flight_bytes": 2751297, "execution_min_parallel_bytes": 0, "execution_min_task_bytes": 0, "root_workers": width, "root_io_concurrency": 1, "root_cpu_tasks": 2}
        for width in (1, 2, 4)
    }
    return (
        {**common_medium, "id": "medium-one-percent", "shape": "medium", "workload": "one-percent", "permits": {"1": 0, "2": 2, "4": 4}, "cpu_tasks": 5, "output_bytes": 4226788, "output_sha256": "fc742a7ad139e2ce2e75fabfff8d2a33408dd5eeb10997717a2e2656d8423e6c", "semantic_sha256": "f705d4200e8bbe840abb16ff7d17da887cbe1d0fb53a649c69279451ff5d829c", "context_limits": medium_limits},
        {**common_dense, "id": "dense-sparse-one-percent", "shape": "dense-sparse", "workload": "one-percent", "permits": {"1": 0, "2": 0, "4": 0}, "cpu_tasks": 5, "output_bytes": 4252436, "output_sha256": "07eee3f980949f031776087ed61f06125edd82cb29394a608a3d8d5a4b28374d", "semantic_sha256": "d8bbafac4c31288eb48a773df88b6b2add2a51b9ad6fb652ab1aeb801c5dfa8c", "context_limits": dense_limits},
        {**common_medium, "id": "medium-one-edit", "shape": "medium", "workload": "one-edit", "permits": {"1": 0, "2": 0, "4": 0}, "cpu_tasks": 2, "output_bytes": 4226480, "output_sha256": "9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1", "semantic_sha256": "3cd21160d4f74fa0f097ab40be08e211b3e460cea788aa2b6705a55fdece07de", "context_limits": edit_limits},
    )

def build_expected_rows(plan: dict[str, Any]) -> list[dict[str, Any]]:
    cases = plan.get("cases")
    comparisons = plan.get("comparisons")
    blocks = plan.get("blocks")
    require(cases == list(exact_case_ids()), "plan cases differ from the frozen three-case matrix")
    require(comparisons == [
        {"id": "width1-vs-width2", "serial_width": 1, "treatment_width": 2},
        {"id": "width1-vs-width4", "serial_width": 1, "treatment_width": 4},
    ], "plan comparisons differ from the frozen width pairs")
    require(blocks == [
        {"id": 0, "order": ["A", "B", "B", "A"]},
        {"id": 1, "order": ["B", "A", "A", "B"]},
        {"id": 2, "order": ["A", "B", "B", "A"]},
        {"id": 3, "order": ["B", "A", "A", "B"]},
        {"id": 4, "order": ["A", "B", "B", "A"]},
        {"id": 5, "order": ["B", "A", "A", "B"]},
    ], "plan block schedule differs from frozen ABBA/BAAB sequence")
    rows: list[dict[str, Any]] = []
    for case in cases:
        for comparison in comparisons:
            for block in blocks:
                for position, role in enumerate(block["order"]):
                    width = comparison["serial_width"] if role == "A" else comparison["treatment_width"]
                    rows.append({
                        "id": f"{case['id']}__{comparison['id']}__b{block['id']:02d}__p{position}",
                        "case_id": case["id"],
                        "shape": case["shape"],
                        "workload": case["workload"],
                        "comparison_id": comparison["id"],
                        "block": block["id"],
                        "position": position,
                        "role": role,
                        "width": width,
                        "serial_width": comparison["serial_width"],
                        "treatment_width": comparison["treatment_width"],
                    })
    return rows


def validate_plan(plan: dict[str, Any], *, for_execution: bool = False) -> list[dict[str, Any]]:
    require(plan.get("schema") == SCHEMA, "plan schema mismatch")
    require(plan.get("native_performance_claim") is True, "native paired-latency claim must be explicitly enabled")
    require(plan.get("adoption_claim") is False and plan.get("allocation_claim") is False, "adoption/allocation claims must remain disabled")
    require(plan.get("production_changes") is False, "production_changes must be false")
    host = plan.get("host")
    require(isinstance(host, dict), "host object missing")
    require(host.get("cpu_mask") == [12, 13, 14, 15], "CPU mask must be 12-15")
    execution = plan.get("execution")
    require(isinstance(execution, dict), "execution object missing")
    require(execution.get("warmup") == 20 and execution.get("samples") == 200, "warmup/sample schedule changed")
    require(execution.get("deadline_seconds") == 900, "deadline must be 900 seconds")
    require(execution.get("serial_fresh_processes") is True, "fresh serial process policy missing")
    require(execution.get("retry_or_replacement") is False, "retry/replacement policy changed")
    require(execution.get("terminal_before_parse") is True, "terminal-first policy missing")
    require(execution.get("bootstrap") == {"seed": 956, "resamples": 10000, "unit": "six block ratios", "statistic": "geometric_mean", "quantiles": [0.025, 0.975]}, "bootstrap contract changed")
    cli = plan.get("cli")
    require(isinstance(cli, dict), "CLI contract missing")
    expected_flags = ["--shape", "--workload", "--workers", "--warmup", "--samples", "--output"]
    require(cli.get("flags") == expected_flags, "CLI flags changed")
    require(cli.get("output_file") == "report.json", "report file name changed")
    require(cli.get("report_schema") == REPORT_SCHEMA, "report schema binding missing")
    rows = build_expected_rows(plan)
    declared = plan.get("schedule")
    require(isinstance(declared, list) and declared == rows, "declared schedule does not equal deterministic schedule")
    require(len(rows) == 144, "schedule must contain 144 children")
    require(len({row["id"] for row in rows}) == len(rows), "schedule child IDs are not unique")
    paths: set[str] = set()
    for row in rows:
        path = f"children/{row['id']}"
        require(path not in paths, "duplicate child path")
        paths.add(path)
    binary = plan.get("binary")
    require(isinstance(binary, dict), "binary binding missing")
    if for_execution:
        require(plan.get("status") == "authorized", "plan is not authorized")
        require(binary.get("sha256") and len(binary["sha256"]) == 64, "binary SHA-256 is not bound")
        binary_path = Path(binary.get("path", ""))
        require(binary_path.is_file() and os.access(binary_path, os.X_OK), f"binary is not executable: {binary_path}")
        require(sha256(binary_path) == binary["sha256"], "binary hash changed")
        require(os.environ.get(APPROVAL_ENV) == "1", f"{APPROVAL_ENV}=1 is required")
    return rows


def process_group_clear(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    return False


def terminate_group(pgid: int) -> bool:
    try:
        os.killpg(pgid, signal.SIGTERM)
    except ProcessLookupError:
        return True
    deadline = time.monotonic() + 5.0
    while time.monotonic() < deadline:
        if process_group_clear(pgid):
            return True
        time.sleep(0.05)
    try:
        os.killpg(pgid, signal.SIGKILL)
    except ProcessLookupError:
        return True
    deadline = time.monotonic() + 5.0
    while time.monotonic() < deadline:
        if process_group_clear(pgid):
            return True
        time.sleep(0.05)
    return process_group_clear(pgid)


def clean_env(plan: dict[str, Any]) -> dict[str, str]:
    configured = plan.get("environment")
    require(isinstance(configured, dict), "environment binding missing")
    env: dict[str, str] = {}
    for key, value in configured.items():
        require(isinstance(key, str) and isinstance(value, str), "environment entries must be strings")
        env[key] = value
    require("PATH" in env and "LC_ALL" in env and "TZ" in env, "runtime environment is incomplete")
    for forbidden in ("RUSTFLAGS", "RUSTDOCFLAGS", "CARGO_ENCODED_RUSTFLAGS", "CARGO_HOME", "RUSTUP_TOOLCHAIN"):
        require(forbidden not in env, f"forbidden inherited environment key {forbidden}")
    return env


def run_command(row: dict[str, Any], plan: dict[str, Any], packet: Path, index: int) -> dict[str, Any]:
    child = packet / "children" / row["id"]
    child.mkdir(parents=True, exist_ok=False)
    output_dir = child / "workload-output"
    binary = Path(plan["binary"]["path"])
    exe = [
        str(binary), "--shape", row["shape"], "--workload", row["workload"],
        "--workers", str(row["width"]), "--warmup", str(plan["execution"]["warmup"]),
        "--samples", str(plan["execution"]["samples"]), "--output", str(output_dir),
    ]
    time_file = child / "time-v.txt"
    stdout_file = child / "stdout.log"
    stderr_file = child / "stderr.log"
    report_file = output_dir / plan["cli"]["output_file"]
    paths = {"child": str(child), "output": str(output_dir), "report": str(report_file), "time": str(time_file), "stdout": str(stdout_file), "stderr": str(stderr_file)}
    command = ["/usr/bin/time", "-v", "-o", str(time_file), "--"] + exe
    for path in (time_file, stdout_file, stderr_file):
        require(not path.exists(), f"refusing to overwrite {path}")
    require(not output_dir.exists() and not output_dir.is_symlink(), "workload output directory already exists")
    cwd = Path(plan.get("cwd", binary.parent))
    require(cwd.is_dir() and not cwd.is_symlink(), f"benchmark cwd is unavailable: {cwd}")
    env = clean_env(plan)
    mask = set(plan["host"]["cpu_mask"])
    start_record = {
        "schema": RUNNER_SCHEMA + ".start",
        "index": index,
        "row": row,
        "argv": command,
        "cwd": str(cwd.resolve()),
        "environment": env,
        "cpu_mask_requested": sorted(mask),
        "started_utc": utc(),
        "deadline_seconds": plan["execution"]["deadline_seconds"],
        "terminal_path": str(child / "terminal.json"),
    }
    write_exclusive(child / "start.json", start_record)
    stdout = stdout_file.open("xb")
    stderr = stderr_file.open("xb")
    launch_monotonic = time.monotonic()
    forced_cleanup = False
    timed_out = False
    returncode: int | None = None
    observed_affinity: list[int] | None = None

    def set_affinity() -> None:
        os.sched_setaffinity(0, mask)

    try:
        proc = subprocess.Popen(
            command, cwd=str(cwd), env=env, stdout=stdout, stderr=stderr,
            start_new_session=True, preexec_fn=set_affinity,
        )
        # Observe the inherited mask while the /usr/bin/time process is live;
        # a requested preexec mask alone is not an affinity witness.
        for _ in range(100):
            try:
                observed_affinity = sorted(os.sched_getaffinity(proc.pid))
                break
            except (ProcessLookupError, PermissionError):
                if proc.poll() is not None:
                    break
                time.sleep(0.01)
        try:
            returncode = proc.wait(timeout=plan["execution"]["deadline_seconds"])
        except subprocess.TimeoutExpired:
            timed_out = True
            forced_cleanup = not terminate_group(proc.pid)
            returncode = proc.wait()
        if not process_group_clear(proc.pid):
            forced_cleanup = True
            if not terminate_group(proc.pid):
                forced_cleanup = True
    finally:
        stdout.close()
        stderr.close()
    terminal = {
        "schema": RUNNER_SCHEMA + ".terminal",
        "index": index,
        "row": row,
        "argv": command,
        "cwd": str(cwd.resolve()),
        "environment": env,
        "cpu_mask_requested": sorted(mask),
        "cpu_mask_observed": observed_affinity,
        "affinity_status": "verified" if observed_affinity == sorted(mask) else "unobserved_or_mismatch",
        "started_utc": start_record["started_utc"],
        "ended_utc": utc(),
        "elapsed_seconds": time.monotonic() - launch_monotonic,
        "returncode": returncode,
        "deadline_exceeded": timed_out,
        "forced_cleanup": forced_cleanup,
        "process_group_clear": not forced_cleanup,
        "status": "completed" if returncode == 0 and not timed_out and not forced_cleanup and observed_affinity == sorted(mask) else "failed",
        "artifacts": {},
    }
    # This is sealed before touching report.json.  The descriptors below are
    # post-terminal evidence and may be absent on a process failure.
    write_exclusive(child / "terminal.json", terminal)
    for label, path in (("stdout", stdout_file), ("stderr", stderr_file), ("time", time_file)):
        if path.is_file():
            terminal["artifacts"][label] = descriptor(path)
    if report_file.is_file() and not report_file.is_symlink():
        terminal["artifacts"]["report"] = descriptor(report_file)
    write_exclusive(child / "post-terminal-artifacts.json", terminal["artifacts"])
    if terminal["status"] != "completed":
        write_exclusive(child / "validation.json", {"schema": RUNNER_SCHEMA + ".validation", "status": "failed", "error": "process terminal failure", "terminal": descriptor(child / "terminal.json")})
        return terminal
    try:
        require(report_file.stat().st_size < REPORT_LIMIT, f"report is at least {REPORT_LIMIT} bytes")
        report = read_json(report_file)
        validate_report(report, row, plan)
        validation = {"schema": RUNNER_SCHEMA + ".validation", "status": "passed", "row": row, "report": descriptor(report_file), "terminal": descriptor(child / "terminal.json")}
    except Exception as exc:
        validation = {"schema": RUNNER_SCHEMA + ".validation", "status": "failed", "row": row, "report": terminal["artifacts"].get("report"), "terminal": descriptor(child / "terminal.json"), "error": str(exc)}
    write_exclusive(child / "validation.json", validation)
    terminal["validation_status"] = validation["status"]
    # Do not replace terminal.json: the first terminal receipt is immutable.
    return terminal


def validate_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any]) -> None:
    require(report.get("schema") == REPORT_SCHEMA, "benchmark report schema changed")
    require(report.get("status") == "completed", "benchmark report is not completed")
    for key, expected in (("shape", row["shape"]), ("workload", row["workload"]), ("workers", row["width"]), ("warmup_iterations", plan["execution"]["warmup"]), ("measured_samples", plan["execution"]["samples"])):
        require(report.get(key) == expected, f"report.{key} does not match schedule")
    case = report.get("case")
    require(isinstance(case, dict), "report.case is missing")
    for key, expected in (("id", f"{row['shape']}-{row['workload']}-workers{row['width']}"), ("shape", row["shape"]), ("workload", row["workload"]), ("workers", row["width"]), ("warmup_iterations", plan["execution"]["warmup"]), ("measured_samples", plan["execution"]["samples"])):
        require(case.get(key) == expected, f"report.case.{key} does not match schedule")
    case_cfg = next(item for item in plan["cases"] if item["id"] == row["case_id"])
    require(case.get("admission_expected_worker_permits") == case_cfg["permits"][str(row["width"])], "admission expected width changed")
    require(case.get("admission_observed_worker_permits") == case_cfg["permits"][str(row["width"])], "configured worker width was narrowed")
    require(case.get("admission_cpu_tasks_used") == case_cfg["cpu_tasks"], "CpuTasks admission changed")
    require(case.get("expected_cpu_tasks") == case_cfg["cpu_tasks"], "case expected CpuTasks changed")
    require(case.get("source_bytes") == case_cfg["source_bytes"] and case.get("output_bytes") == case_cfg["output_bytes"], "case corpus byte lengths changed")
    require(case.get("source_sha256") == case_cfg["source_sha256"], "case source oracle changed")
    require(case.get("output_sha256") == case_cfg["output_sha256"], "case output oracle changed")
    require(case.get("semantic_sha256") == case_cfg["semantic_sha256"], "case semantic oracle changed")
    require(case.get("context_limits") == case_cfg["context_limits"][str(row["width"])], "finite context limits changed")
    require(case.get("preflight_actual_commit_inverse_verified") is True, "actual managed-commit inverse preflight is not verified")
    samples = case.get("samples")
    require(isinstance(samples, list) and len(samples) == plan["execution"]["samples"], "case.samples length changed")
    phases = ("ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")
    for expected_index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"sample {expected_index} is not an object")
        require(sample.get("sample") == expected_index, f"sample index changed at {expected_index}")
        values = [int_value(sample.get(phase), f"sample[{expected_index}].{phase}") for phase in phases]
        require(sum(values[:4]) <= values[4], f"phase sum exceeds total at sample {expected_index}")
        require(sample.get("fixture_inverse_restored_exact_source") is True, f"fixture inverse check failed at sample {expected_index}")
        require(isinstance(sample.get("output_sha256"), str) and sample["output_sha256"] == case["output_sha256"], f"sample output hash changed at {expected_index}")
        require(isinstance(sample.get("semantic_sha256"), str) and sample["semantic_sha256"] == case.get("semantic_sha256"), f"sample semantic hash changed at {expected_index}")
        require(type(sample.get("output_bytes")) is int and sample["output_bytes"] == case.get("output_bytes"), f"sample output size changed at {expected_index}")
        require(type(sample.get("output_bytes_used")) is int and sample["output_bytes_used"] == sample["output_bytes"], f"sample output accounting changed at {expected_index}")
        require(type(sample.get("cpu_tasks_used")) is int and sample["cpu_tasks_used"] == case_cfg["cpu_tasks"], f"sample CpuTasks changed at {expected_index}")
        require(type(sample.get("sink_write_calls")) is int and sample["sink_write_calls"] > 0, f"sink write count missing at {expected_index}")
        require(type(sample.get("sink_largest_write")) is int and 0 < sample["sink_largest_write"] <= 65536, f"sink write bound changed at {expected_index}")
        resources = sample.get("resources_after_drop")
        require(isinstance(resources, dict) and all(resources.get(key) == 0 for key in ("memory", "workers", "io_concurrency", "objects")), f"post-drop resources not zero at {expected_index}")
    require(isinstance(report.get("phase_scope"), str) and "publication" in report["phase_scope"], "phase scope missing")


def parse_time(path: Path) -> dict[str, float | int]:
    result: dict[str, float | int] = {}
    require(path.is_file(), f"time -v output missing: {path}")
    for line in path.read_text(errors="replace").splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip()
        field = TIME_FIELDS.get(key)
        if field is None:
            continue
        value = value.strip()
        try:
            if field in ("user_seconds", "system_seconds"):
                result[field] = float(value.split()[0])
            else:
                result[field] = int(value.split()[0])
        except (ValueError, IndexError) as exc:
            raise ValueError(f"invalid /usr/bin/time field {key}: {value!r}") from exc
    required = set(TIME_FIELDS.values())
    require(required <= result.keys(), f"missing /usr/bin/time fields: {sorted(required - result.keys())}")
    return result


def run(plan_path: Path, packet: Path, *, execute: bool) -> int:
    plan = read_json(plan_path)
    rows = validate_plan(plan, for_execution=execute)
    if not execute:
        print(json.dumps({"status": "plan_valid", "rows": len(rows), "children": 144}, sort_keys=True))
        return 0
    require(not packet.exists(), f"refusing existing packet: {packet}")
    packet.mkdir(parents=True)
    (packet / "children").mkdir(exist_ok=False)
    write_exclusive(packet / "capture-start.json", {"schema": RUNNER_SCHEMA + ".outer-start", "plan": descriptor(plan_path), "runner": descriptor(Path(__file__)), "started_utc": utc(), "row_count": len(rows)})
    terminal_status: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        terminal = run_command(row, plan, packet, index)
        terminal_status.append(terminal)
        if terminal.get("status") != "completed" or terminal.get("validation_status") != "passed":
            break
    status = "completed" if len(terminal_status) == len(rows) and all(item.get("validation_status") == "passed" for item in terminal_status) else "failed"
    final = {"schema": RUNNER_SCHEMA + ".outer-terminal", "status": status, "ended_utc": utc(), "children_started": len(terminal_status), "children_expected": len(rows), "failed_index": None if status == "completed" else len(terminal_status) - 1}
    write_exclusive(packet / "capture-terminal.json", final)
    return 0 if status == "completed" else 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args()
    try:
        return run(args.plan, args.packet, execute=args.execute)
    except Exception as exc:
        print(f"capture_0956_v1: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
