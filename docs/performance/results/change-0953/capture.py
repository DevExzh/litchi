#!/usr/bin/env python3
"""Create-only first-child-gated capture runner for the 0953 FP diagnostic (reviewed decoder contract).

The runner does not build or interpret profiles.  It verifies the final 0953
plan and the complete frozen 0950 source census, records one serial child with
``perf --call-graph fp``, and exports samples and MMAP2/task events from the
same sealed perf.data.  The sample export omits ``--debug-mode``; the MMAP2/task export uses the known-good 0952 ``--debug-mode`` path because perf emits sample headers otherwise.  Every receipt is exclusive;
any failure is retained and stops the schedule.  Root authorization and the
explicit environment approval are required before a workload is launched.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import signal
import stat
import subprocess
import sys
import time
from typing import Any

PACKET = Path(__file__).resolve().parents[1]
HERE = Path(__file__).resolve()
PLAN_DEFAULT = PACKET / "plan-0953-v1.json"
P0950 = Path("/home/zhuhe/litchi-performance-0950")
SOURCE_ROOT = P0950 / "candidate-source-v1"
SOURCE_MANIFEST = P0950 / "candidate-source-manifest-v1.json"
FREEZE_RECEIPT = P0950 / "receipts" / "candidate-freeze-v1.json"
CAPTURE_ROOT = PACKET / "capture"
ANALYSIS_ROOT = PACKET / "analysis"
FIRST_CHILD_ADMISSION = CAPTURE_ROOT / "first-child-admission.json"
PERF = Path("/usr/bin/perf")
TASKSET = Path("/usr/bin/taskset")
NM = Path("/usr/bin/x86_64-linux-gnu-nm")
READELF = Path("/usr/bin/x86_64-linux-gnu-readelf")
CPU = 12
FREQUENCY = 997
EVENT = "cycles:u"
CALL_GRAPH = "fp"
CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
SAMPLES = 200
WARMUP = 0
DEADLINE_SECONDS = 900
APPROVAL_ENV = "LITCHI_PERF_0953_CAPTURE_APPROVED"
APPROVAL_VALUE = "1"
SOURCE_SCHEMA = "litchi.performance.0950.candidate-source-manifest.v1"
PLAN_SCHEMA = "litchi.performance.0953.native-publish-sampling-plan.v1"
ANALYSIS_SCHEMA = "litchi.performance.0953.native-admission.v3"
ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LANG": "C",
    "LC_ALL": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
}
SCHEDULE = [
    {"id": "native-fp-publish-medium-r0", "index": 0, "repeat": 0, "shape": "medium"},
    {"id": "native-fp-publish-dense-sparse-r0", "index": 1, "repeat": 0, "shape": "dense-sparse"},
    {"id": "native-fp-publish-dense-sparse-r1", "index": 2, "repeat": 1, "shape": "dense-sparse"},
    {"id": "native-fp-publish-medium-r1", "index": 3, "repeat": 1, "shape": "medium"},
]


class CaptureError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def _no_symlink_components(path: Path) -> None:
    absolute = path if path.is_absolute() else (Path.cwd() / path)
    cursor = Path(absolute.anchor)
    for part in absolute.parts[1:] if absolute.anchor else absolute.parts:
        cursor /= part
        require(not cursor.is_symlink(), f"symlink path component is not allowed: {cursor}")


def descriptor(path: Path) -> dict[str, Any]:
    path = path.absolute()
    _no_symlink_components(path)
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    mode = stat.S_IMODE(path.stat().st_mode)
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha256_path(path),
        "mode": mode,
        "executable": bool(mode & stat.S_IXUSR),
    }


def _unique_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    _no_symlink_components(path)
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=_unique_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise CaptureError(f"invalid JSON {path}: {error}") from error


def write_json_exclusive(path: Path, value: Any) -> None:
    _no_symlink_components(path.parent)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8", newline="") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, ensure_ascii=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise CaptureError(f"refusing to overwrite retained artifact: {path}") from error


def resolve_descriptor(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    raw = value.get("path")
    require(isinstance(raw, str) and raw and "\x00" not in raw, f"{label}: path invalid")
    path = Path(raw).absolute()
    actual = descriptor(path)
    for key in ("bytes", "sha256", "mode", "executable"):
        require(actual[key] == value.get(key), f"{label}: {key} differs")
    return path


def run_readonly(argv: list[str], *, cwd: Path = SOURCE_ROOT, timeout: int = 60) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(argv, cwd=str(cwd), env=dict(ENVIRONMENT),
                              stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                              stderr=subprocess.PIPE, text=True, check=False, timeout=timeout)
    except (OSError, subprocess.SubprocessError) as error:
        raise CaptureError(f"read-only command failed to launch: {argv}: {error}") from error


def verify_source_census(plan: dict[str, Any]) -> dict[str, Any]:
    source = plan.get("source_binding")
    require(isinstance(source, dict), "source_binding is missing")
    manifest_spec = source.get("manifest")
    freeze_spec = source.get("freeze_receipt")
    manifest = resolve_descriptor(manifest_spec, "source manifest")
    freeze = resolve_descriptor(freeze_spec, "source freeze receipt")
    require(manifest == SOURCE_MANIFEST.absolute(), "source manifest path differs")
    require(freeze == FREEZE_RECEIPT.absolute(), "source freeze receipt path differs")
    value = load_json(manifest)
    require(value.get("schema") == SOURCE_SCHEMA, "source manifest schema differs")
    require(value.get("source_root") == SOURCE_ROOT.name, "source manifest root differs")
    entries = value.get("files")
    require(isinstance(entries, list), "source manifest file list is missing")
    expected: dict[str, dict[str, Any]] = {}
    for entry in entries:
        require(isinstance(entry, dict), "source manifest row is not an object")
        relative = entry.get("path")
        require(isinstance(relative, str) and relative and not Path(relative).is_absolute(),
                "source manifest path is invalid")
        relative_path = Path(relative)
        require(".." not in relative_path.parts and "\\" not in relative,
                f"source manifest path escapes source root: {relative}")
        require(relative not in expected, f"duplicate source manifest path: {relative}")
        expected[relative] = entry
    require(len(expected) == value.get("file_count"), "source manifest file count differs")
    total_bytes = 0
    for relative, entry in expected.items():
        path = SOURCE_ROOT / relative
        _no_symlink_components(path)
        require(path.is_file() and not path.is_symlink(), f"missing or symlink source file: {relative}")
        actual = descriptor(path)
        require(actual["bytes"] == entry.get("bytes"), f"source bytes differ: {relative}")
        require(actual["sha256"] == entry.get("sha256"), f"source digest differs: {relative}")
        total_bytes += actual["bytes"]
    actual_paths: set[str] = set()
    for directory, directories, files in os.walk(SOURCE_ROOT, topdown=True, followlinks=False):
        directory_path = Path(directory)
        for name in directories:
            require(not (directory_path / name).is_symlink(), f"source directory is a symlink: {directory_path / name}")
        for name in files:
            path = directory_path / name
            require(not path.is_symlink() and path.is_file(), f"source entry is not a regular file: {path}")
            actual_paths.add(path.relative_to(SOURCE_ROOT).as_posix())
    require(actual_paths == set(expected), "source file set differs from complete manifest")
    require(total_bytes == value.get("bytes"), "source byte total differs")
    return {"manifest": descriptor(manifest), "freeze_receipt": descriptor(freeze),
            "file_count": len(expected), "bytes": total_bytes}


def schedule_from_plan(plan: dict[str, Any]) -> list[dict[str, Any]]:
    schedule = plan.get("schedule")
    require(schedule == SCHEDULE, "capture schedule differs")
    require(plan.get("cpu") == CPU and plan.get("samples") == SAMPLES
            and plan.get("warmup") == WARMUP and plan.get("deadline_seconds") == DEADLINE_SECONDS,
            "capture sample/CPU/deadline contract differs")
    require(plan.get("retry") is False and plan.get("serial") is True,
            "capture must be serial and non-retrying")
    return schedule


def _expand_template(template: Any, *, child: Path, binary: Path,
                     shape: str, label: str) -> list[str]:
    require(isinstance(template, list) and template,
            f"{label} argv template is missing")
    replacements = {
        "$PACKET": str(PACKET.absolute()),
        "$CHILD": str(child.absolute()),
        "$BINARY": str(binary.absolute()),
        "$SHAPE": shape,
    }
    result: list[str] = []
    for item in template:
        require(isinstance(item, str), f"{label} argv template contains a non-string")
        value = item
        for marker, replacement in replacements.items():
            value = value.replace(marker, replacement)
        result.append(value)
    return result


def verify_command_templates(plan: dict[str, Any], binary: Path) -> None:
    """Compare every plan template with the exact argv constructors below.

    The comparison uses a non-existent synthetic child path, so this check is
    purely static and cannot launch a workload or create capture artifacts.
    """
    row = SCHEDULE[0]
    child = PACKET / "capture" / "__0953_v5_contract_child"
    expected_record = record_argv(binary, child, row)
    record = plan.get("record_contract")
    require(isinstance(record, dict), "record_contract is missing")
    require(_expand_template(record.get("argv_template"), child=child,
                             binary=binary, shape=row["shape"], label="record")
            == expected_record,
            "record argv template differs from the runner constructor")
    require(record.get("call_graph") == CALL_GRAPH
            and record.get("call_graph_mode") == CALL_GRAPH,
            "record contract frame-pointer binding differs")
    decoders = plan.get("decoder_contract")
    require(isinstance(decoders, dict), "decoder_contract is missing")
    for kind, key, debug in (("sample", "primary_stack_decoder", False),
                             ("mmap", "mapping_decoder", True)):
        contract = decoders.get(key)
        require(isinstance(contract, dict), f"{kind} decoder contract is missing")
        actual = decoder_argv(child, kind)
        expected = _expand_template(contract.get("argv_template"), child=child,
                                    binary=binary, shape=row["shape"], label=kind)
        require(expected == actual,
                f"{kind} decoder argv template differs from the runner constructor")
        require(contract.get("debug_mode") is debug,
                f"{kind} decoder debug-mode binding differs")


def _approval_descriptor(approval: dict[str, Any], label: str) -> Path:
    """Resolve a GO descriptor, accepting the two reviewed receipt shapes."""
    descriptor_fields = {"bytes", "sha256", "mode", "executable"}
    if descriptor_fields.issubset(approval):
        return resolve_descriptor(approval, label)
    for key in ("descriptor", "receipt"):
        nested = approval.get(key)
        if isinstance(nested, dict) and descriptor_fields.issubset(nested):
            return resolve_descriptor(nested, f"{label}.{key}")
    raise CaptureError(f"{label}: complete file descriptor is missing")


def verify_plan(plan_path: Path, *, require_go: bool = True) -> tuple[dict[str, Any], Path, dict[str, Any]]:
    plan = load_json(plan_path)
    require(isinstance(plan, dict) and plan.get("schema") == PLAN_SCHEMA, "0953 plan schema differs")
    require(plan.get("status") not in {"failed", "rejected"}, "plan is not admissible")
    require(plan.get("no_build_or_workload_until_root_go") is True,
            "plan does not retain the root-GO lock")
    runner = resolve_descriptor(plan.get("runner"), "plan runner")
    require(runner == HERE.absolute(), "plan runner path differs")
    binary = resolve_descriptor(plan.get("binary"), "plan diagnostic binary")
    require(binary.name == "litchi-perf-baseline" and binary != (P0950 / "target-candidate-native-v1/release/litchi-perf-baseline").absolute(),
            "plan binary is not the separate FP diagnostic executable")
    require(binary.stat().st_mode & stat.S_IXUSR, "diagnostic binary is not executable")
    source_census = verify_source_census(plan)
    require(plan.get("case") == CASE and plan.get("owner")
            == "litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream",
            "workload owner/case contract differs")
    require(plan.get("measured_parent")
            == "litchi_perf_baseline::run_xlsx_cell_values_edit_save",
            "workload parent contract differs")
    require(plan.get("analysis_schema", ANALYSIS_SCHEMA) == ANALYSIS_SCHEMA,
            "plan analysis schema differs from native-admission v3")
    record = plan.get("record_contract", {})
    require(record.get("call_graph") == CALL_GRAPH or record.get("call_graph_mode") == CALL_GRAPH,
            "record contract does not select frame-pointer callchains")
    argv = record.get("argv_template", [])
    if argv:
        joined = " ".join(str(x) for x in argv)
        require("--call-graph fp" in joined and "--debug-mode" not in joined,
                "record argv does not bind fp without debug mode")
        require("dwarf" not in joined, "record argv contains forbidden DWARF callgraph")
    # These values are used by every actual process invocation.  Bind them in
    # the final plan before any child is launched.
    capture_environment = plan.get("capture_environment", plan.get("environment"))
    require(capture_environment == ENVIRONMENT,
            "capture environment differs from the approved exact environment")
    capture_cwd = plan.get("capture_cwd", plan.get("cwd"))
    require(capture_cwd == str(SOURCE_ROOT),
            "capture cwd differs from the approved source-root cwd")
    schedule_from_plan(plan)
    verify_command_templates(plan, binary)
    analyzer_value = plan.get("analyzer")
    require(isinstance(analyzer_value, dict),
            "plan analyzer descriptor is missing")
    analyzer = resolve_descriptor(analyzer_value, "plan analyzer")
    require(analyzer.name == "analyze_0953_v3.py",
            "plan analyzer is not the reviewed native-admission v3 reader")
    approval = plan.get("capture_authorization") or plan.get("root_go")
    require(isinstance(approval, dict),
            "root capture authorization receipt is missing")
    if require_go:
        require(approval.get("status") == "pass",
                "root capture authorization receipt is missing or not passing")
        _approval_descriptor(approval, "root capture authorization")
    elif approval.get("status") == "pass":
        _approval_descriptor(approval, "root capture authorization")
    tools = plan.get("tools", {})
    for name, expected in (("perf", PERF), ("taskset", TASKSET)):
        if name in tools:
            require(resolve_descriptor(tools[name], f"tool {name}") == expected.absolute(),
                    f"tool {name} path differs")
    require(run_readonly([str(PERF), "--version"]).returncode == 0, "perf version probe failed")
    require(run_readonly([str(TASKSET), "--cpu-list", str(CPU), "true"]).returncode == 0,
            "CPU affinity probe failed")
    return plan, binary, {"source_census": source_census, "approval": approval,
                          "analyzer": analyzer}


def child_row(plan: dict[str, Any], index: int) -> dict[str, Any]:
    for row in plan["schedule"]:
        if row["index"] == index:
            return row
    raise CaptureError(f"unknown schedule index: {index}")


def record_argv(binary: Path, child: Path, row: dict[str, Any]) -> list[str]:
    argv = [str(TASKSET), "--cpu-list", str(CPU), str(PERF), "record",
            "--output", str(child / "perf.data"), "--freq", str(FREQUENCY),
            "--event", EVENT, "--call-graph", CALL_GRAPH, "--period", "--timestamp",
            "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--strict-freq",
            "--no-inherit", "--user-callchains", "--", str(binary),
            "--warmup", str(WARMUP), "--samples", str(SAMPLES), "--case", CASE,
            "--xlsx-cell-crud-shape", str(row["shape"]), "--json", str(child / "report.json")]
    require(argv.count("--call-graph") == 1 and argv[argv.index("--call-graph") + 1] == CALL_GRAPH,
            "record command callgraph is not exactly fp")
    require("--debug-mode" not in argv and all("dwarf" not in item for item in argv),
            "record command contains forbidden decoder/callgraph option")
    return argv


def decoder_argv(child: Path, kind: str) -> list[str]:
    base = [str(PERF), "script", "--input", str(child / "perf.data")]
    if kind == "sample":
        extra = ["--show-lost-events", "--ns", "--no-inline", "--demangle",
                 "--fields", "comm,pid,tid,time,event,period,ip,sym,symoff,dso"]
    elif kind == "mmap":
        extra = ["--debug-mode", "--show-mmap-events", "--show-task-events", "--show-lost-events",
                 "--ns", "--hide-call-graph", "--no-inline", "--demangle", "--fields",
                 "comm,pid,tid,time,event,period,ip,sym,symoff,dso"]
    else:
        raise CaptureError(f"unknown decoder kind: {kind}")
    argv = base + extra
    if kind == "sample":
        require("--debug-mode" not in argv, "sample decoder must not use --debug-mode")
    else:
        require("--debug-mode" in argv, "mmap decoder must use the reviewed debug-mode export path")
    if kind == "sample":
        require("--show-mmap-events" not in argv and "--show-task-events" not in argv,
                "sample decoder unexpectedly includes mapping records")
    else:
        require("--show-mmap-events" in argv and "--show-task-events" in argv
                and "--hide-call-graph" in argv, "mmap decoder contract differs")
    return argv


def verify_actual_process_contract(plan: dict[str, Any], binary: Path,
                                   row: dict[str, Any], child: Path,
                                   *, kind: str | None = None) -> list[str]:
    """Bind the argv/env/cwd of the imminent process to the frozen plan."""
    require(plan.get("capture_environment", plan.get("environment")) == ENVIRONMENT,
            "process environment differs from frozen capture environment")
    require(plan.get("capture_cwd", plan.get("cwd")) == str(SOURCE_ROOT),
            "process cwd differs from frozen capture cwd")
    if kind is None:
        actual = record_argv(binary, child, row)
        expected = _expand_template(plan["record_contract"]["argv_template"],
                                    child=child, binary=binary, shape=row["shape"],
                                    label="record")
    else:
        actual = decoder_argv(child, kind)
        contract_key = "primary_stack_decoder" if kind == "sample" else "mapping_decoder"
        expected = _expand_template(plan["decoder_contract"][contract_key]["argv_template"],
                                    child=child, binary=binary, shape=row["shape"],
                                    label=kind)
    require(actual == expected,
            f"actual {kind or 'record'} argv differs from frozen plan template")
    return actual


MMAP_RECORD = re.compile(
    r"^\s*\S+\s+\d+/\d+\s+(?P<time>[0-9]+(?:\.[0-9]+)?)\s*:\s+"
    r"PERF_RECORD_MMAP2\s+(?P<pid>\d+)/(?P<tid>\d+)\s*:\s+"
    r"\[(?P<start>0x[0-9a-fA-F]+)\((?P<size>0x[0-9a-fA-F]+)\)\s+@\s+"
    r"(?P<offset>(?:0x)?[0-9a-fA-F]+)\s+(?P<meta>.*?)\]:\s+"
    r"(?P<perm>\S+)\s+(?P<dso>.*)$"
)


def binary_build_id(binary: Path) -> str:
    result = run_readonly([str(READELF), "-nW", str(binary)], cwd=SOURCE_ROOT, timeout=60)
    require(result.returncode == 0, "readelf build-ID inspection failed")
    ids = re.findall(r"Build ID:\s*([0-9a-fA-F]+)", result.stdout)
    require(len(ids) == 1 and ids[0], "diagnostic binary does not have one GNU build ID")
    return ids[0].lower()


def validate_mmap_export(events: Path, binary: Path, build_id: str) -> dict[str, Any]:
    """Admit only a mapping export containing the exact new executable DSO.

    PT_LOAD compatibility and time/IP load-bias selection remain the reducer's
    responsibility.  This capture-side check prevents an empty, malformed, or
    ordinary-binary MMAP2 export from entering that reducer.
    """
    target = str(binary.absolute())
    bound: list[dict[str, Any]] = []
    unsupported = 0
    for number, line in enumerate(events.read_text(encoding="utf-8", errors="strict").splitlines(), 1):
        if not line.strip():
            continue
        match = MMAP_RECORD.match(line)
        if match is not None:
            dso = match.group("dso").strip()
            identifiers = re.findall(r"<([0-9a-fA-F]+)>", match.group("meta"))
            row = {"line": number, "dso": dso, "perm": match.group("perm"),
                   "build_id": identifiers[0].lower() if len(identifiers) == 1 else None,
                   "size": int(match.group("size"), 16)}
            require(row["size"] > 0, f"invalid MMAP2 size at line {number}")
            if dso == target or Path(dso).name == binary.name:
                require(row["build_id"] == build_id,
                        f"bound MMAP2 build ID differs at line {number}")
                bound.append(row)
            continue
        if "PERF_RECORD_" in line:
            # Non-MMAP records are retained for the reducer; only malformed
            # MMAP2 records are rejected here.
            if "PERF_RECORD_MMAP2" in line:
                raise CaptureError(f"malformed MMAP2 record at line {number}")
            continue
        unsupported += 1
    require(unsupported == 0, "mapping export contains unsupported text")
    require(bound and any("x" in row["perm"] for row in bound),
            "mapping export has no executable mapping for the new binary")
    return {"target_path": target, "build_id": build_id,
            "bound_rows": len(bound),
            "executable_bound_rows": sum("x" in row["perm"] for row in bound),
            "unsupported_records": unsupported}


def process_group_exists(pgid: int | None) -> bool:
    if pgid is None:
        return True
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def clear_process_group(pgid: int | None) -> bool:
    if pgid is None or not process_group_exists(pgid):
        return True
    for signum, wait_seconds in ((signal.SIGTERM, 2.0), (signal.SIGKILL, 2.0)):
        try:
            os.killpg(pgid, signum)
        except ProcessLookupError:
            return True
        deadline = time.monotonic() + wait_seconds
        while process_group_exists(pgid) and time.monotonic() < deadline:
            time.sleep(0.05)
        if not process_group_exists(pgid):
            return True
    return not process_group_exists(pgid)


def run_process(argv: list[str], cwd: Path, stdout_path: Path, stderr_path: Path,
                *, timeout: int = DEADLINE_SECONDS) -> dict[str, Any]:
    started = utc_now()
    monotonic = time.monotonic()
    returncode: int | None = None
    process_group_id: int | None = None
    timed_out = False
    launch_error: str | None = None
    with stdout_path.open("x", encoding="utf-8") as stdout, stderr_path.open("x", encoding="utf-8") as stderr:
        try:
            process = subprocess.Popen(argv, cwd=str(cwd), env=dict(ENVIRONMENT),
                                       stdin=subprocess.DEVNULL, stdout=stdout, stderr=stderr,
                                       start_new_session=True)
            process_group_id = process.pid
            try:
                returncode = process.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                timed_out = True
                try:
                    os.killpg(process.pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    try:
                        os.killpg(process.pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                    process.wait()
                returncode = process.returncode
        except (OSError, subprocess.SubprocessError) as error:
            launch_error = str(error)
            returncode = 127
    return {"argv": argv, "started_utc": started, "terminal_utc": utc_now(),
            "elapsed_seconds": time.monotonic() - monotonic, "returncode": returncode,
            "timed_out": timed_out, "launch_error": launch_error,
            "process_group_id": process_group_id,
            "process_group_clear": clear_process_group(process_group_id),
            "terminal_before_parse": True}


def terminal_pass(value: dict[str, Any], label: str) -> None:
    require(value.get("returncode") == 0 and value.get("timed_out") is False
            and value.get("launch_error") is None and value.get("process_group_clear") is True
            and value.get("terminal_before_parse") is True,
            f"{label} terminal failed")


def require_prior_child(plan: dict[str, Any], index: int) -> None:
    row = child_row(plan, index - 1)
    child = CAPTURE_ROOT / f"{index - 1:02d}-{row['shape']}"
    receipt_path = child / "receipt.json"
    require(receipt_path.is_file() and not receipt_path.is_symlink(),
            "serial predecessor receipt is missing")
    receipt = load_json(receipt_path)
    require(receipt.get("status") == "raw_captured" and receipt.get("retry") is False,
            "serial predecessor was not admitted raw capture")
    for name in ("record-terminal.json",):
        terminal_pass(load_json(child / name), f"serial predecessor {name}")
    for kind in ("sample", "mmap"):
        export = ANALYSIS_ROOT / kind / f"{index - 1:02d}-{row['shape']}"
        value = load_json(export / "receipt.json")
        require(value.get("status") == "pass" and value.get("workload_launched") is False,
                f"serial predecessor {kind} decoder did not pass")
        terminal_pass(load_json(export / "terminal.json"), f"serial predecessor {kind} decoder")


def _equal_descriptor(given: Any, expected: dict[str, Any], label: str) -> Path:
    """Require a descriptor to resolve to exactly the expected artifact."""
    path = resolve_descriptor(given, label)
    expected_path = Path(expected["path"]).absolute()
    require(path == expected_path, f"{label}: path differs")
    actual = descriptor(expected_path)
    for key in ("bytes", "sha256", "mode", "executable"):
        require(given.get(key) == actual[key], f"{label}: {key} differs")
    return path


def _packet_analysis_path(path: Path, label: str) -> Path:
    """Keep the analyzer result in this packet's analysis namespace."""
    path = path.absolute()
    try:
        relative = path.relative_to(PACKET.absolute())
    except ValueError as error:
        raise CaptureError(f"{label}: path escapes 0953 packet") from error
    require(relative.parts and relative.parts[0] == "analysis",
            f"{label}: path is outside the packet analysis namespace")
    require(".." not in relative.parts and not path.is_symlink(),
            f"{label}: path is not a regular packet artifact")
    return path


def _require_first_child_admission_value(value: dict[str, Any], plan: dict[str, Any],
                                         plan_path: Path, binary: Path) -> None:
    """Validate the descriptor and gate contract written by the root/analyzer.

    This function deliberately consumes an already loaded JSON object so the
    offline controls can exercise a mutated admission without touching the
    sealed capture tree.
    """
    require(value.get("schema") == "litchi.performance.0953.first-child-admission.v1"
            and value.get("status") == "pass" and value.get("child_index") == 0,
            "first-child admission has not passed for child zero")
    require(value.get("raw_validation") == "pass"
            and value.get("semantic_validation") == "pass",
            "first-child admission lacks raw/semantic pass")
    require(value.get("retry") is False and value.get("replacement") is False,
            "first-child admission permits retry or replacement")
    expected_plan = descriptor(plan_path)
    expected_runner = descriptor(HERE)
    expected_binary = descriptor(binary)
    _equal_descriptor(value.get("plan"), expected_plan, "first-child admission plan")
    _equal_descriptor(value.get("runner"), expected_runner, "first-child admission runner")
    _equal_descriptor(value.get("binary"), expected_binary, "first-child admission binary")
    analyzer_spec = plan.get("analyzer")
    require(isinstance(analyzer_spec, dict),
            "final plan analyzer descriptor is missing")
    analyzer_path = resolve_descriptor(analyzer_spec, "plan analyzer")
    _equal_descriptor(value.get("analyzer"), descriptor(analyzer_path),
                      "first-child admission analyzer")

    row = child_row(plan, 0)
    child = CAPTURE_ROOT / f"00-{row['shape']}"
    child_receipt_path = child / "receipt.json"
    require(child_receipt_path.is_file() and not child_receipt_path.is_symlink(),
            "first-child receipt is missing")
    child_receipt_desc = descriptor(child_receipt_path)
    _equal_descriptor(value.get("child_receipt"), child_receipt_desc,
                      "first-child admission child receipt")
    child_receipt = load_json(child_receipt_path)
    require(child_receipt.get("status") == "raw_captured"
            and child_receipt.get("row") == row
            and child_receipt.get("retry") is False,
            "first-child receipt is not an admitted non-retry capture")
    _equal_descriptor(child_receipt.get("plan"), expected_plan,
                      "first-child receipt plan")
    _equal_descriptor(child_receipt.get("runner"), expected_runner,
                      "first-child receipt runner")
    _equal_descriptor(child_receipt.get("binary"), expected_binary,
                      "first-child receipt binary")

    analysis_path_value = value.get("analysis_result")
    require(isinstance(analysis_path_value, dict),
            "first-child admission analysis_result descriptor is missing")
    analysis_path = _packet_analysis_path(
        resolve_descriptor(analysis_path_value, "first-child admission analysis_result"),
        "first-child admission analysis_result")
    analysis_desc = descriptor(analysis_path)
    _equal_descriptor(analysis_path_value, analysis_desc,
                      "first-child admission analysis_result")
    analysis = load_json(analysis_path)
    require(analysis.get("schema") == ANALYSIS_SCHEMA
            and analysis.get("status") == "pass",
            "first-child native analysis is not a passing v3 result")
    _equal_descriptor(analysis.get("plan"), expected_plan,
                      "native analysis plan")
    _equal_descriptor(analysis.get("runner"), expected_runner,
                      "native analysis runner")
    _equal_descriptor(analysis.get("binary"), expected_binary,
                      "native analysis binary")
    _equal_descriptor(analysis.get("child"), child_receipt_desc,
                      "native analysis child receipt")

    semantic = analysis.get("semantic")
    require(isinstance(semantic, dict) and semantic.get("status") == "pass"
            and semantic.get("samples") == SAMPLES
            and semantic.get("warmup") == WARMUP
            and semantic.get("shape") == row["shape"],
            "native analysis semantic gate is not the approved 200-sample result")
    profile = analysis.get("profile")
    require(isinstance(profile, dict), "native analysis profile is missing")
    all_data = profile.get("all")
    all_samples = all_data.get("samples") if isinstance(all_data, dict) else None
    all_period = all_data.get("period") if isinstance(all_data, dict) else None
    owner = profile.get("owner_qualified")
    unknown = profile.get("unknown_or_unqualified")
    gate = profile.get("gate")
    require(isinstance(all_samples, int) and all_samples > 0
            and isinstance(all_period, int) and all_period > 0
            and isinstance(owner, dict) and isinstance(unknown, dict)
            and isinstance(gate, dict),
            "native analysis sample profile is missing")
    raw_gate = plan.get("first_child_admission", {}).get("raw_gate", {})
    require(raw_gate.get("owner_samples_min") == 100
            and raw_gate.get("fully_resolved_owner_samples_min") == 100
            and raw_gate.get("owner_period_min") == 100000
            and raw_gate.get("lost_records") == 0,
            "first-child plan gate is not the frozen 100/100000 gate")
    require(isinstance(owner.get("samples"), int)
            and 0 < owner["samples"] <= all_samples
            and owner["samples"] >= raw_gate["owner_samples_min"]
            and isinstance(owner.get("period"), int)
            and 0 < owner["period"] <= all_period
            and owner["period"] >= raw_gate["owner_period_min"],
            "native analysis owner gate is below the approved 100/100000 minimum")
    require(isinstance(unknown.get("samples"), int)
            and isinstance(unknown.get("period"), int)
            and owner["samples"] + unknown["samples"] == all_samples
            and owner["period"] + unknown["period"] == all_period,
            "native analysis sample/period partition is not conserved")
    require(gate.get("lost_records_zero") is True
            and gate.get("fully_resolved_owner_samples_min") is True
            and gate.get("owner_period_min") is True
            and profile.get("admitted") is True,
            "native analysis raw gate did not pass")
    # The admission itself repeats the numeric gate so a later child cannot
    # be released by a result file whose summary and admission disagree.
    require(value.get("minimum_owner_samples") == 100
            and value.get("minimum_owner_period") == 100000
            and isinstance(value.get("owner_samples"), int)
            and value["owner_samples"] == owner["samples"]
            and isinstance(value.get("owner_period"), int)
            and value["owner_period"] == owner["period"],
            "first-child admission does not bind the 100/100000 owner gates")


def require_first_child_admission(plan: dict[str, Any], plan_path: Path,
                                  binary: Path) -> None:
    require(FIRST_CHILD_ADMISSION.is_file() and not FIRST_CHILD_ADMISSION.is_symlink(),
            "first-child admission is missing")
    value = load_json(FIRST_CHILD_ADMISSION)
    require(isinstance(value, dict), "first-child admission is not an object")
    _require_first_child_admission_value(value, plan, plan_path, binary)


def export_decoder(plan: dict[str, Any], plan_path: Path, row: dict[str, Any], child: Path,
                   child_start: dict[str, Any], kind: str) -> dict[str, Any]:
    output = ANALYSIS_ROOT / kind / child.name
    require(not output.exists(), f"refusing to overwrite decoder output: {output}")
    output.mkdir(parents=True)
    binary = Path(plan["binary"]["path"]).absolute()
    argv = verify_actual_process_contract(plan, binary, row, child, kind=kind)
    build_id = binary_build_id(binary)
    start = {"schema": f"litchi.performance.0953.{kind}-export-start.v1",
             "row": row, "plan": descriptor(plan_path), "runner": descriptor(HERE),
             "child_start": descriptor(child / "start.json"), "perf_data": descriptor(child / "perf.data"),
             "binary": plan["binary"], "build_id": build_id, "argv": argv, "cwd": str(SOURCE_ROOT),
             "environment": ENVIRONMENT, "workload_launched": False, "started_utc": utc_now()}
    write_json_exclusive(output / "start.json", start)
    terminal = run_process(argv, SOURCE_ROOT, output / "events.txt", output / "stderr.txt")
    write_json_exclusive(output / "terminal.json", {
        "schema": f"litchi.performance.0953.{kind}-export-terminal.v1", **terminal})
    success = (terminal.get("returncode") == 0 and terminal.get("timed_out") is False
               and terminal.get("launch_error") is None and terminal.get("process_group_clear") is True)
    require((output / "events.txt").is_file(), f"{kind} event output missing")
    size = (output / "events.txt").stat().st_size
    mapping: dict[str, Any] | None = None
    validation_error: str | None = None
    if success and size > 0 and kind == "mmap":
        try:
            mapping = validate_mmap_export(output / "events.txt", binary, build_id)
        except (CaptureError, OSError, UnicodeError) as error:
            success = False
            validation_error = str(error)
    receipt = {"schema": f"litchi.performance.0953.{kind}-export-receipt.v1",
               "status": "pass" if success and size > 0 else "failed", "row": row,
               "start": descriptor(output / "start.json"),
               "terminal": descriptor(output / "terminal.json"),
               "events": descriptor(output / "events.txt"),
               "stderr": descriptor(output / "stderr.txt"),
               "perf_data": descriptor(child / "perf.data"), "binary": plan["binary"],
               "build_id": build_id, "mapping_validation": mapping,
               "validation_error": validation_error, "argv": argv,
               "terminal_before_parse": True, "workload_launched": False}
    write_json_exclusive(output / "receipt.json", receipt)
    return receipt


def run_child(plan: dict[str, Any], plan_path: Path, index: int) -> dict[str, Any]:
    row = child_row(plan, index)
    child = CAPTURE_ROOT / f"{index:02d}-{row['shape']}"
    binary = Path(plan["binary"]["path"]).absolute()
    require(not child.exists(), f"refusing to overwrite child: {child}")
    require(not ANALYSIS_ROOT.exists() or not (ANALYSIS_ROOT / "sample" / child.name).exists(),
            "refusing to overwrite sample decoder output")
    require(not ANALYSIS_ROOT.exists() or not (ANALYSIS_ROOT / "mmap" / child.name).exists(),
            "refusing to overwrite mmap decoder output")
    if index > 0:
        require_first_child_admission(plan, plan_path, binary)
        require_prior_child(plan, index)
    child.mkdir(parents=True)
    record = verify_actual_process_contract(plan, binary, row, child)
    start = {"schema": "litchi.performance.0953.child-start.v1", "row": row,
             "plan": descriptor(plan_path), "runner": descriptor(HERE),
             "binary": plan["binary"], "source_root": str(SOURCE_ROOT), "cwd": str(SOURCE_ROOT),
             "environment": ENVIRONMENT, "record_argv": record, "started_utc": utc_now(),
             "retry": False, "workload_launched": True}
    write_json_exclusive(child / "start.json", start)
    command = {"schema": "litchi.performance.0953.child-command.v1", "row": row,
               "plan": descriptor(plan_path), "runner": descriptor(HERE),
               "binary": plan["binary"], "source_root": str(SOURCE_ROOT), "cwd": str(SOURCE_ROOT),
               "environment": ENVIRONMENT, "record_argv": record, "terminal_before_parse": True,
               "retry": False}
    write_json_exclusive(child / "command.json", command)
    record_term = run_process(record, SOURCE_ROOT, child / "record.stdout", child / "record.stderr")
    write_json_exclusive(child / "record-terminal.json", {
        "schema": "litchi.performance.0953.record-terminal.v1", **record_term})
    status = "record_failed"
    sample_receipt: dict[str, Any] | None = None
    mmap_receipt: dict[str, Any] | None = None
    if record_term.get("returncode") == 0 and record_term.get("timed_out") is False \
            and record_term.get("launch_error") is None and record_term.get("process_group_clear") is True:
        require((child / "perf.data").is_file() and not (child / "perf.data").is_symlink(),
                "perf record passed without perf.data")
        require((child / "report.json").is_file() and not (child / "report.json").is_symlink(),
                "perf record passed without semantic report")
        sample_receipt = export_decoder(plan, plan_path, row, child, start, "sample")
        mmap_receipt = export_decoder(plan, plan_path, row, child, start, "mmap")
        if sample_receipt.get("status") == "pass" and mmap_receipt.get("status") == "pass":
            status = "raw_captured"
        else:
            status = "decode_failed"
    artifacts: list[dict[str, Any]] = []
    for path in sorted(child.iterdir()):
        require(path.name in {"start.json", "command.json", "record.stdout", "record.stderr",
                              "record-terminal.json", "report.json", "perf.data", "receipt.json"},
                f"unexpected child artifact: {path.name}")
        if path.is_file() and path.name != "receipt.json":
            artifacts.append(descriptor(path))
    receipt = {"schema": "litchi.performance.0953.child-receipt.v1", "status": status,
               "row": row, "plan": descriptor(plan_path), "runner": descriptor(HERE),
               "binary": plan["binary"], "source_root": str(SOURCE_ROOT), "cwd": str(SOURCE_ROOT),
               "environment": ENVIRONMENT, "record_terminal": descriptor(child / "record-terminal.json"),
               "report": descriptor(child / "report.json") if (child / "report.json").is_file() else None,
               "perf_data": descriptor(child / "perf.data") if (child / "perf.data").is_file() else None,
               "sample_export": sample_receipt, "mmap_export": mmap_receipt,
               "artifacts": artifacts, "terminal_before_parse": True, "retry": False,
               "workload_launched": True}
    write_json_exclusive(child / "receipt.json", receipt)
    print(json.dumps({"status": status, "index": index, "child": str(child)}, sort_keys=True))
    require(status == "raw_captured", "capture or decoder failed; retained artifacts stop schedule")
    return receipt


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("check-plan", "capture"))
    parser.add_argument("--plan", type=Path, default=PLAN_DEFAULT)
    parser.add_argument("--index", type=int)
    args = parser.parse_args()
    try:
        plan_path = args.plan.absolute()
        plan, _binary, _metadata = verify_plan(plan_path)
        if args.action == "check-plan":
            print(json.dumps({"status": "pass", "plan": descriptor(plan_path),
                              "runner": descriptor(HERE), "schedule": plan["schedule"]},
                             sort_keys=True))
            return 0
        require(args.index in range(4), "capture requires --index 0..3")
        require(os.environ.get(APPROVAL_ENV) == APPROVAL_VALUE,
                f"{APPROVAL_ENV}=1 is required for capture")
        receipt = run_child(plan, plan_path, int(args.index))
        return 0 if receipt.get("status") == "raw_captured" else 2
    except (CaptureError, OSError, ValueError, TypeError) as error:
        print(f"capture_0953_v5: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
