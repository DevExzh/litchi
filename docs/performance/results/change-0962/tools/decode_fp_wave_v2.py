#!/usr/bin/env python3
"""Bounded offline decoder for the retained 0962 FP wave.

This controller never starts the benchmark and never builds anything.  It
requires the capture lane's validated terminal receipt before launching even
one ``perf script`` process.  Each capture gets independent inline and
no-inline exports, terminal receipts, and a replay through the frozen
``analyze_native_cpu_v1.py`` helper.  ``--all`` runs the six captures in
parallel; ``--index`` is useful when six external workers own one child each.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import datetime
import hashlib
import json
import os
import signal
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0962")
DEFAULT_EXECUTION_PLAN = ROOT / "fp-wave-execution-v1.json"
DEFAULT_CAPTURE_ROOT = ROOT / "fp-wave-v1"
FROZEN_ANALYZER_SHA256 = "f5f7a387f99ba3a587eeb8f27c8f422a9fa696306c2ae47724952b861910bd3d"
DEFAULT_DEADLINE_SECONDS = 900
SCRIPT_FIELDS = "comm,pid,tid,time,event,period,ip,sym,symoff,dso,srcline"
SCHEMA = "litchi.performance.0962.fp-wave-decoder.v2"


class DecodeError(RuntimeError):
    """A decoder admission, terminal, or conservation failure."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise DecodeError(message)


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path, label: str = "file") -> dict[str, Any]:
    path = path.expanduser()
    require(path.is_absolute(), f"{label} must be absolute: {path}")
    require(path.exists() and path.is_file(), f"{label} is not a regular file: {path}")
    require(not path.is_symlink(), f"{label} must not be a symlink: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_file(path)}


def write_create(path: Path, value: dict[str, Any]) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def read_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise DecodeError(f"could not read {label}: {path}: {exc}") from exc
    require(isinstance(value, dict), f"{label} must be a JSON object: {path}")
    return value


def bound_environment() -> tuple[dict[str, str], str]:
    env = os.environ.copy()
    env.update({"LANG": "C", "LC_ALL": "C", "TZ": "UTC"})
    visible_keys = ("PATH", "LANG", "LC_ALL", "TZ", "PERF_BUILDID_DIR", "RUST_BACKTRACE")
    visible = {key: env[key] for key in visible_keys if key in env}
    material = "\0".join(f"{key}={env[key]}" for key in sorted(env))
    return visible, hashlib.sha256(material.encode("utf-8")).hexdigest()


def process_group_clear(pid: int, grace_seconds: float = 5.0) -> bool:
    """Allow perf's short-lived helper children to leave the group."""
    deadline = time.monotonic() + grace_seconds
    while True:
        try:
            os.killpg(pid, 0)
        except ProcessLookupError:
            return True
        except PermissionError:
            return False
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.25)


def terminate_process(process: subprocess.Popen[Any]) -> None:
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        process.wait(timeout=10)


def run_bounded(
    command: list[str],
    *,
    cwd: Path,
    env: dict[str, str],
    stdout_path: Path,
    stderr_path: Path,
    terminal_path: Path,
    deadline_seconds: int,
) -> dict[str, Any]:
    require(not stdout_path.exists() and not stderr_path.exists(), "decoder export already exists")
    start = now()
    failure: str | None = None
    timed_out = False
    with stdout_path.open("x", encoding="utf-8") as stdout, stderr_path.open("x", encoding="utf-8") as stderr:
        try:
            process = subprocess.Popen(
                command,
                cwd=str(cwd),
                env=env,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
            )
        except OSError as exc:
            failure = repr(exc)
            raise DecodeError(f"could not start bounded decoder: {exc}") from exc
        try:
            returncode = process.wait(timeout=deadline_seconds)
        except subprocess.TimeoutExpired:
            timed_out = True
            terminate_process(process)
            returncode = process.returncode
        except BaseException as exc:  # pragma: no cover - defensive terminal path
            failure = repr(exc)
            terminate_process(process)
            returncode = process.returncode
    terminal = {
        "schema": SCHEMA + ".terminal",
        "started_utc": start,
        "finished_utc": now(),
        "argv": command,
        "cwd": str(cwd),
        "returncode": returncode,
        "timeout": timed_out,
        "deadline_seconds": deadline_seconds,
        "pid": process.pid,
        "terminal_before_parse": True,
        "failure": failure,
        "process_group_clear": process_group_clear(process.pid),
    }
    write_create(terminal_path, terminal)
    require(returncode == 0 and not timed_out and failure is None, f"bounded decoder failed: {terminal}")
    require(terminal["process_group_clear"], f"decoder process group remains: {terminal}")
    return terminal


def validate_capture_lane(
    execution_plan_path: Path,
    execution_plan: dict[str, Any],
    capture_root: Path,
) -> tuple[dict[str, Any], dict[str, Any], list[str]]:
    terminal_path = capture_root / "capture-terminal.json"
    require(terminal_path.is_file(), "capture terminal receipt is required before offline decode")
    terminal = read_json(terminal_path, "capture terminal receipt")
    require(terminal.get("status") == "capture_validated", "capture lane is not validated")
    children = terminal.get("children")
    capture_plan_path = Path(execution_plan["capture_plan"]["path"])
    capture_plan = read_json(capture_plan_path, "capture plan")
    schedule = capture_plan.get("wave_schedule")
    require(isinstance(children, list) and isinstance(schedule, list), "capture schedule is malformed")
    require(len(schedule) == 6 and len(children) == 6, "offline decoder requires exactly six captures")
    require(all(row.get("returncode") == 0 for row in children), "capture child failed")
    require(capture_plan.get("status") == "authorized", "capture plan is not authorized")
    expected_binary = capture_plan["binary"]
    binary_path = Path(expected_binary["path"])
    require(descriptor(binary_path, "bound binary") == expected_binary, "bound binary changed")
    analyzer = Path(execution_plan["analyzer"]["path"])
    analyzer_desc = descriptor(analyzer, "frozen analyzer")
    require(analyzer_desc["sha256"] == FROZEN_ANALYZER_SHA256, "frozen analyzer hash changed")
    require(analyzer_desc == execution_plan["analyzer"], "execution plan analyzer descriptor changed")
    require(execution_plan_path.is_file(), "execution plan disappeared")
    cases: list[str] = []
    for index, case in enumerate(schedule):
        child_dir = capture_root / f"{index:02d}-{case}"
        child_terminal = child_dir / "terminal.json"
        validation = child_dir / "validation.json"
        perf_data = child_dir / "perf.data"
        require(child_terminal.is_file(), f"missing capture terminal for child {index}")
        require(validation.is_file(), f"missing capture validation for child {index}")
        child_value = read_json(child_terminal, f"child {index} terminal")
        require(
            child_value.get("returncode") == 0
            and child_value.get("timeout") is False
            and child_value.get("process_group_clear") is True,
            f"capture child {index} terminal gate failed",
        )
        require(read_json(validation, f"child {index} validation").get("status") == "passed", f"child {index} validation failed")
        descriptor(perf_data, f"child {index} perf.data")
        cases.append(str(child_dir))
    return terminal, capture_plan, cases


def decode_one(
    index: int,
    case_dir: Path,
    *,
    analyzer: Path,
    binary: Path,
    deadline_seconds: int,
    env: dict[str, str],
    env_hash: str,
    execution_plan_path: Path,
    capture_terminal: dict[str, Any],
) -> dict[str, Any]:
    decode_dir = case_dir / "offline-decode-v2"
    require(not decode_dir.exists(), f"decoder output already exists: {decode_dir}")
    decode_dir.mkdir()
    perf_data = case_dir / "perf.data"
    old_decode_dir = case_dir / "offline-decode-v1"
    old_inline_script = old_decode_dir / "perf-script-inline.txt"
    old_inline_stderr = old_decode_dir / "perf-script-inline.stderr"
    old_inline_terminal = old_decode_dir / "inline-terminal.json"
    inline_late_clearance: dict[str, Any] | None = None
    inline_term: dict[str, Any]
    if old_inline_script.is_file() and old_inline_stderr.is_file() and old_inline_terminal.is_file():
        old_terminal = read_json(old_inline_terminal, f"child {index} prior inline terminal")
        require(old_terminal.get("returncode") == 0 and old_terminal.get("timeout") is False, f"child {index} prior inline export failed")
        before = descriptor(old_inline_script, f"child {index} prior inline export")
        require(process_group_clear(int(old_terminal["pid"])), f"child {index} prior inline process group remains")
        after = descriptor(old_inline_script, f"child {index} prior inline export")
        require(before == after, f"child {index} prior inline export changed")
        inline_late_clearance = {
            "schema": SCHEMA + ".late-clearance",
            "observed_utc": now(),
            "source_terminal": descriptor(old_inline_terminal, f"child {index} prior inline terminal"),
            "source_terminal_value": old_terminal,
            "source_script": after,
            "source_stderr": descriptor(old_inline_stderr, f"child {index} prior inline stderr"),
            "late_process_group_clear": True,
            "reused_without_rerunning_perf_script": True,
        }
        late_path = old_decode_dir / "inline-terminal-late-clearance-v2.json"
        write_create(late_path, inline_late_clearance)
        inline_script = old_inline_script
        inline_stderr = old_inline_stderr
        inline_term = old_terminal
        inline_command = list(old_terminal["argv"])
    else:
        inline_script = decode_dir / "perf-script-inline.txt"
        inline_stderr = decode_dir / "perf-script-inline.stderr"
        inline_terminal = decode_dir / "inline-terminal.json"
        inline_term = {}
        inline_command = []
    physical_script = decode_dir / "perf-script-physical.txt"
    physical_stderr = decode_dir / "perf-script-physical.stderr"
    physical_terminal = decode_dir / "physical-terminal.json"
    common = [
        "/usr/bin/perf",
        "script",
        "--show-lost-events",
        "--demangle",
        "--full-source-path",
        "-F",
        SCRIPT_FIELDS,
        "-i",
        str(perf_data),
    ]
    if not inline_command:
        inline_command = common[:4] + ["--inline"] + common[4:]
    physical_command = common[:4] + ["--no-inline"] + common[4:]
    if not inline_term:
        inline_term = run_bounded(
            inline_command,
            cwd=ROOT,
            env=env,
            stdout_path=inline_script,
            stderr_path=inline_stderr,
            terminal_path=inline_terminal,
            deadline_seconds=deadline_seconds,
        )
    physical_term = run_bounded(
        physical_command,
        cwd=ROOT,
        env=env,
        stdout_path=physical_script,
        stderr_path=physical_stderr,
        terminal_path=physical_terminal,
        deadline_seconds=deadline_seconds,
    )
    output = decode_dir / "analyzer.json"
    decoder_stdout = decode_dir / "analyzer.stdout"
    decoder_stderr = decode_dir / "analyzer.stderr"
    decoder_terminal = decode_dir / "decoder-terminal.json"
    command = [
        "/usr/bin/python3",
        str(analyzer),
        "--script",
        str(inline_script),
        "--physical-script",
        str(physical_script),
        "--binary",
        str(binary),
        "--output",
        str(output),
    ]
    decoder_term = run_bounded(
        command,
        cwd=ROOT,
        env=env,
        stdout_path=decoder_stdout,
        stderr_path=decoder_stderr,
        terminal_path=decoder_terminal,
        deadline_seconds=deadline_seconds,
    )
    result = read_json(output, f"child {index} analyzer output")
    require(result.get("status") == "pass", f"child {index} analyzer status is not pass")
    gate = result.get("qualification_gate", {})
    pairing = result.get("pairing", {})
    conservation = result.get("conservation", {})
    require(gate.get("admitted") is True, f"child {index} admission gate failed")
    require(gate.get("lost_records_zero") is True, f"child {index} lost-record gate failed")
    require(gate.get("qualified_samples", 0) >= 100, f"child {index} qualified sample gate failed")
    require(gate.get("qualified_period", 0) >= 100000, f"child {index} qualified period gate failed")
    require(pairing.get("exact_pairing_pass") is True, f"child {index} exact pairing gate failed")
    require(conservation.get("inline_sample_count_pass") is True, f"child {index} sample conservation failed")
    require(conservation.get("inline_period_sum_pass") is True, f"child {index} period conservation failed")
    require(result.get("binary_binding", {}).get("sha256") == descriptor(binary, "bound binary")["sha256"], f"child {index} binary binding changed")
    control = {
        "schema": SCHEMA,
        "status": "passed",
        "index": index,
        "case_directory": str(case_dir),
        "execution_plan": descriptor(execution_plan_path, "execution plan"),
        "capture_terminal": capture_terminal,
        "capture_perf_data": descriptor(perf_data, "perf.data"),
        "analyzer": descriptor(analyzer, "frozen analyzer"),
        "binary": descriptor(binary, "bound binary"),
        "environment": {"visible": {key: env[key] for key in ("PATH", "LANG", "LC_ALL", "TZ") if key in env}, "sha256": env_hash},
        "deadline_seconds": deadline_seconds,
        "exports": {
            "inline": {
                "command": inline_command,
                "terminal": inline_term,
                "script": descriptor(inline_script),
                "stderr": descriptor(inline_stderr),
                "reused_from_prior_controller": inline_late_clearance is not None,
                "late_clearance": inline_late_clearance,
            },
            "physical": {"command": physical_command, "terminal": physical_term, "script": descriptor(physical_script), "stderr": descriptor(physical_stderr)},
        },
        "decoder": {"command": command, "terminal": decoder_term, "output": descriptor(output), "stdout": descriptor(decoder_stdout), "stderr": descriptor(decoder_stderr)},
        "qualification_gate": gate,
        "pairing": pairing,
        "conservation": conservation,
    }
    write_create(decode_dir / "decode-control.json", control)
    return control


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execution-plan", type=Path, default=DEFAULT_EXECUTION_PLAN)
    parser.add_argument("--capture-root", type=Path, default=DEFAULT_CAPTURE_ROOT)
    parser.add_argument("--deadline-seconds", type=int, default=DEFAULT_DEADLINE_SECONDS)
    parser.add_argument("--index", type=int, action="append", help="decode one or more validated child indexes")
    parser.add_argument("--plan", action="store_true", help="validate and print the six-child plan without launching decoders")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        require(args.deadline_seconds > 0, "deadline must be positive")
        execution_plan_path = args.execution_plan.expanduser().absolute()
        capture_root = args.capture_root.expanduser().absolute()
        execution_plan = read_json(execution_plan_path, "execution plan")
        capture_terminal, capture_plan, case_paths = validate_capture_lane(execution_plan_path, execution_plan, capture_root)
        if args.plan:
            print(json.dumps({"schema": SCHEMA, "status": "plan_only", "children": case_paths, "helper_sha256": FROZEN_ANALYZER_SHA256}, sort_keys=True))
            return 0
        indexes = sorted(set(args.index if args.index is not None else range(6)))
        require(indexes and all(0 <= index < 6 for index in indexes), "child index must be 0 through 5")
        analyzer = Path(execution_plan["analyzer"]["path"])
        binary = Path(capture_plan["binary"]["path"])
        env = os.environ.copy()
        env.update({"LANG": "C", "LC_ALL": "C", "TZ": "UTC"})
        _, env_hash = bound_environment()
        with concurrent.futures.ThreadPoolExecutor(max_workers=len(indexes)) as pool:
            futures = [
                pool.submit(
                    decode_one,
                    index,
                    Path(case_paths[index]),
                    analyzer=analyzer,
                    binary=binary,
                    deadline_seconds=args.deadline_seconds,
                    env=env,
                    env_hash=env_hash,
                    execution_plan_path=execution_plan_path,
                    capture_terminal=capture_terminal,
                )
                for index in indexes
            ]
            rows = [future.result() for future in futures]
        print(json.dumps({"schema": SCHEMA, "status": "passed", "children": [row["index"] for row in rows]}, sort_keys=True))
        return 0
    except (DecodeError, OSError, UnicodeError, ValueError, subprocess.SubprocessError) as exc:
        print(json.dumps({"schema": SCHEMA, "status": "failed", "error": str(exc)}, sort_keys=True), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
