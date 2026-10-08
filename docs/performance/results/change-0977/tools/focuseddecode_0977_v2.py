#!/usr/bin/env python3
"""Bounded offline decoder for one retained 0977 perf child.

The capture runner owns the native command, terminal receipt, launch receipt,
and raw ``perf.data``.  This module only consumes those immutable inputs and
produces decoder artifacts plus a bounded child packet.  It deliberately
loads the sealed 0973 parser/raw scanner and 0974 task/loss predicates without
installing any runner patches.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, Mapping


ROOT = Path("/home/zhuhe/litchi-performance-0977")
ADAPTER_PATH = Path("/home/zhuhe/litchi-performance-0974/tools/protocol_adapter_0974_v5.py")
ADAPTER_SHA256 = "e3725cbc027d2a063f32b926e3fb4024d640aab2a2fbd7493ed5325de9ea2ac6"
PARSER_PATH = Path("/home/zhuhe/litchi-performance-0973/tools-v2/protocol_common_0973_v2.py")
PARSER_SHA256 = "56b509fde01bbd66702ac8cd835da0d09d74d332935926078c6545093ed52bf4"
RAW_SCANNER_PATH = Path("/home/zhuhe/litchi-performance-0973/tools-v2/raw_stack_0973_v2.py")
RAW_SCANNER_SHA256 = "3a8e36ebc5f21f4f0b8293cafa675d722f2264fb1cfd6c7f48271b4f0a5056d2"

FINAL_SHIM_SOURCE = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.c")
FINAL_SHIM_SOURCE_SHA256 = "a30536989f1fcc3e28c5bbe5876b9f01159841ba0a6d6c3aa22706cdcfd23d1a"
FINAL_SHIM_MAP = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.map")
FINAL_SHIM_MAP_SHA256 = "173dd1902d46a378e713cd794ede6c4fdac2a184124e8aefecc0c2fda7494883"
FINAL_SHIM_SO = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.so")
FINAL_SHIM_SO_SHA256 = "2acdc5c06ac672dca0e5a34bf204e0a45451b4bfe2d31c304f4c791f5b2a0a12"

BASELINE_BINARY_SHA256 = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
BASELINE_BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
BASELINE_BINARY_NAME = "doc_retention_probe_0942_cfb_profile"
PERF_PATH = Path("/usr/bin/perf")
# Fresh 0977 plan binding after the 2026-10-08 distro perf upgrade.  The
# old 0976/0977-v1 decoder snapshot retains the prior hash in its immutable
# receipts; this fresh helper admits only the explicitly requalified binary.
PERF_SHA256 = "1371ae5cb9fa59632d1d0e34b877d7251df4029810ba67dc56e946545e87b531"
DEFAULT_CPU = 31
DEFAULT_PERIOD = 1_000_000
DEFAULT_STACK_CAP = 32_768
EXPECTED_DELTA = 0x1000
TARGET_LOST_RE = re.compile(r"PERF_RECORD_LOST:\s+.*?\blost:(\d+)\b")
TERMINAL_LOST_RE = re.compile(r"^#\s*Total Lost Samples:\s*(\d+)\s*$", re.MULTILINE)
FORK_RE = re.compile(r"PERF_RECORD_(?:FORK|CLONE)\((\d+):(\d+)\):\((\d+):(\d+)\)")
SAMPLE_HEADER_RE = re.compile(
    r"^\s*(?:(?P<comm>\S+)\s+)?(?P<pid>\d+)(?:/(?P<tid>\d+))?"
    r"(?:\s+\[(?P<cpu>\d+)\])?\s+(?P<time>\d+(?:\.\d+)?):\s+"
    r"(?P<period>\d+)\s+(?P<event>[^\n]+):\s*$"
)
FRAME_RE = re.compile(r"^\s*(?P<ip>[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+\((?P<dso>.*)\)\s*$")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _file_binding(path: Path) -> dict[str, Any]:
    path = path.absolute()
    if not path.is_file():
        raise DecodeFailure("missing_artifact", f"missing file: {path}", {"path": str(path)})
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": _sha256(path)}


class DecodeFailure(RuntimeError):
    """Typed fail-closed result for one child."""

    def __init__(self, code: str, message: str, details: Mapping[str, Any] | None = None):
        super().__init__(f"{code}: {message}")
        self.code = code
        self.message = message
        self.details = dict(details or {})

    def as_dict(self) -> dict[str, Any]:
        return {"status": "failed", "code": self.code, "message": self.message, "details": self.details}


def _fail(code: str, message: str, **details: Any) -> None:
    raise DecodeFailure(code, message, details)


def _load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        _fail("immutable_import", f"cannot load retained module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as exc:  # noqa: BLE001 - convert to typed decode failure
        _fail("immutable_import", f"retained module import failed: {path}", error=str(exc))
    return module


def _verify_immutable_inputs() -> None:
    for label, path, expected in (
        ("0974_adapter", ADAPTER_PATH, ADAPTER_SHA256),
        ("0973_parser", PARSER_PATH, PARSER_SHA256),
        ("0973_raw_scanner", RAW_SCANNER_PATH, RAW_SCANNER_SHA256),
        ("final_shim_source", FINAL_SHIM_SOURCE, FINAL_SHIM_SOURCE_SHA256),
        ("final_shim_map", FINAL_SHIM_MAP, FINAL_SHIM_MAP_SHA256),
        ("final_shim_so", FINAL_SHIM_SO, FINAL_SHIM_SO_SHA256),
    ):
        if not path.is_file():
            _fail("immutable_input_missing", f"{label} is missing", path=str(path))
        actual = _sha256(path)
        if actual != expected:
            _fail("immutable_input_changed", f"{label} hash changed", path=str(path), actual=actual, expected=expected)


_verify_immutable_inputs()
ADAPTER = _load_module("focuseddecode_0977_adapter", ADAPTER_PATH)
# ADAPTER imports the sealed parser/raw scanner itself and exposes the exact
# immutable modules.  No install_runner_patches call is made here.
PARSER = ADAPTER.BASE
RAW_SCANNER = ADAPTER.BASE_RAW_STACK


def _load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        _fail("invalid_json", f"cannot load JSON: {path}", error=str(exc))
    if not isinstance(value, dict):
        _fail("invalid_json", f"JSON root is not an object: {path}")
    return value


def _write_json_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
    except FileExistsError:
        _fail("output_exists", "refusing to overwrite decoder output", path=str(path))


def _write_command(path: Path, argv: list[str], env: Mapping[str, str]) -> dict[str, Any]:
    value = {
        "schema": "litchi.performance.0977.decoder_command.v1",
        "argv": argv,
        "environment": dict(env),
        "cwd": str(Path.cwd()),
        "cwd_policy": "inherited decoder cwd; recorded before offline launch",
        "native_workload": False,
    }
    _write_json_exclusive(path, value)
    return value


def _binding_from_plan(row: Any, label: str, *, expected_path: Path | None = None, expected_sha: str | None = None) -> Path:
    if not isinstance(row, dict) or not isinstance(row.get("path"), str):
        _fail("plan_binding", f"missing {label} binding")
    path = Path(row["path"]).absolute()
    if expected_path is not None and path != expected_path.absolute():
        _fail("plan_binding", f"{label} path is not the sealed final input", path=str(path), expected=str(expected_path.absolute()))
    if not path.is_file():
        _fail("plan_binding", f"{label} path is missing", path=str(path))
    actual = _file_binding(path)
    wanted = expected_sha or row.get("sha256")
    if actual["sha256"] != wanted:
        _fail("plan_binding", f"{label} hash mismatch", actual=actual, expected=wanted)
    if row.get("bytes") is not None and row["bytes"] != actual["bytes"]:
        _fail("plan_binding", f"{label} byte count mismatch", actual=actual["bytes"], expected=row["bytes"])
    return path


def _find_plan_binding(plan: Mapping[str, Any], names: tuple[str, ...]) -> Any:
    containers: list[Any] = [plan]
    for key in ("helpers", "dependencies", "shim", "decoder", "helper_seal", "helpers_seal"):
        value = plan.get(key)
        if isinstance(value, Mapping):
            containers.append(value)
    for container in containers:
        for name in names:
            row = container.get(name) if isinstance(container, Mapping) else None
            if isinstance(row, Mapping) and isinstance(row.get("path"), str):
                return row
    return None


def _validate_plan(plan: Mapping[str, Any]) -> tuple[Path, Path, int, int, int, str, str]:
    binary_row = plan.get("binary")
    if not isinstance(binary_row, Mapping) or not isinstance(binary_row.get("path"), str):
        _fail("plan_binding", "plan binary binding is missing")
    binary = Path(str(binary_row["path"])).absolute()
    actual_binary = _file_binding(binary)
    if actual_binary["sha256"] != BASELINE_BINARY_SHA256:
        _fail("baseline_binary", "final_v2 shim only admits the retained baseline binary", actual=actual_binary["sha256"], expected=BASELINE_BINARY_SHA256)
    if binary.name != BASELINE_BINARY_NAME:
        _fail("baseline_binary", "target basename is not the final_v2 baseline", actual=binary.name, expected=BASELINE_BINARY_NAME)
    build_id = str(binary_row.get("build_id") or plan.get("binary_build_id") or BASELINE_BUILD_ID).lower()
    if build_id != BASELINE_BUILD_ID:
        _fail("baseline_build_id", "target build ID is not the final_v2 baseline", actual=build_id, expected=BASELINE_BUILD_ID)

    source = _binding_from_plan(_find_plan_binding(plan, ("shim_source", "decoder_source", "source")), "shim_source", expected_path=FINAL_SHIM_SOURCE, expected_sha=FINAL_SHIM_SOURCE_SHA256)
    version_map = _binding_from_plan(_find_plan_binding(plan, ("shim_map", "decoder_map", "version_script", "map")), "shim_map", expected_path=FINAL_SHIM_MAP, expected_sha=FINAL_SHIM_MAP_SHA256)
    shim = _binding_from_plan(_find_plan_binding(plan, ("shim_so", "decoder_so", "shared_object", "so")), "shim_so", expected_path=FINAL_SHIM_SO, expected_sha=FINAL_SHIM_SO_SHA256)
    # Keep these variables live in the validation result so a caller can bind
    # all three custody inputs without re-reading the plan.
    _ = source, version_map

    perf = plan.get("perf") if isinstance(plan.get("perf"), Mapping) else {}
    perf_path = Path(str(perf.get("path", PERF_PATH))).absolute()
    if perf_path != PERF_PATH.absolute() or _sha256(perf_path) != PERF_SHA256:
        _fail("perf_binding", "offline perf tool is not the sealed executable", path=str(perf_path))
    period = int(perf.get("period", DEFAULT_PERIOD))
    cpu = int(perf.get("cpu", (plan.get("matrix") or {}).get("cpu", DEFAULT_CPU)))
    stack_cap = int(perf.get("stack_cap", 32_768))
    if period <= 0 or cpu < 0 or stack_cap <= 0:
        _fail("perf_contract", "period, CPU, or stack cap is invalid")
    marker_row = plan.get("marker") if isinstance(plan.get("marker"), Mapping) else {}
    marker = str(marker_row.get("symbol", PARSER.MARKER))
    marker_dso = str(marker_row.get("dso_path", binary))
    if marker_dso != str(binary):
        _fail("marker_binding", "marker DSO differs from exact target path", marker_dso=marker_dso, binary=str(binary))
    if str(marker_row.get("build_id", BASELINE_BUILD_ID)).lower() != BASELINE_BUILD_ID:
        _fail("marker_binding", "marker build ID differs from final_v2 baseline")
    return binary, shim, period, cpu, stack_cap, marker, marker_dso


def _input_path(child_dir: Path, suffix: str) -> Path:
    path = child_dir / f"{child_dir.name}.{suffix}"
    if not path.is_file():
        _fail("missing_capture_artifact", "runner artifact is missing", path=str(path))
    return path


def _output_path(child_dir: Path, suffix: str) -> Path:
    return child_dir / f"{child_dir.name}.{suffix}"


def _validate_launch(launch: Mapping[str, Any], child_dir: Path, binary: Path) -> int:
    try:
        pid = int(launch["exec_pid"])
        mirrored = int(launch.get("pid", pid))
    except (KeyError, TypeError, ValueError):
        _fail("launch_contract", "launch receipt lacks an integer exec_pid/pid")
    if pid <= 0 or mirrored != pid:
        _fail("launch_contract", "launch PID identity mismatch", exec_pid=pid, pid=mirrored)
    packet_id = str(launch.get("packet_id", child_dir.name))
    if packet_id != child_dir.name:
        _fail("launch_contract", "launch packet ID differs from child directory", packet_id=packet_id, child=child_dir.name)
    path = launch.get("exec_path")
    if path is not None and str(Path(str(path)).absolute()) != str(binary):
        _fail("launch_contract", "launch executable differs from plan binary", launch_path=path, binary=str(binary))
    if launch.get("execve") is False:
        _fail("launch_contract", "launch receipt says execve failed")
    return pid


def _terminal_loss(terminal: Mapping[str, Any], child_dir: Path) -> int:
    if terminal.get("returncode") not in (None, 0) or terminal.get("timed_out") is True or terminal.get("launch_error") is True:
        _fail("terminal_contract", "terminal receipt records a failed capture", returncode=terminal.get("returncode"), timed_out=terminal.get("timed_out"), launch_error=terminal.get("launch_error"))
    if terminal.get("native_ld_preload") not in (None, "absent"):
        _fail("terminal_contract", "native capture had LD_PRELOAD", native_ld_preload=terminal.get("native_ld_preload"))
    if terminal.get("terminal_report_returncode") not in (None, 0):
        _fail("terminal_contract", "terminal perf report failed", returncode=terminal.get("terminal_report_returncode"))
    summary = terminal.get("lost_samples_summary")
    if not isinstance(summary, Mapping):
        summary = terminal.get("terminal") if isinstance(terminal.get("terminal"), Mapping) else None
        summary = summary.get("lost_samples_summary") if isinstance(summary, Mapping) else None
    values: list[int] = []
    if isinstance(summary, Mapping) and summary.get("record_type") == "PERF_RECORD_LOST_SAMPLES":
        value = summary.get("lost_samples")
        if isinstance(value, int) and value >= 0:
            values.append(value)
    text_sources: list[str] = []
    for key in ("terminal_stdout", "terminal_report", "report_stdout", "stdout"):
        value = terminal.get(key)
        if isinstance(value, str):
            candidate = Path(value)
            if "\n" not in value and candidate.is_file():
                text_sources.append(candidate.read_text(encoding="utf-8", errors="strict"))
            else:
                text_sources.append(value)
    for text in text_sources:
        matches = TERMINAL_LOST_RE.findall(text)
        if len(matches) != 1:
            _fail("terminal_loss", "terminal report must contain exactly one Total Lost Samples line", matches=len(matches), child=str(child_dir))
        values.append(int(matches[0]))
    if not values:
        _fail("terminal_loss", "terminal LOST_SAMPLES value is missing", child=str(child_dir))
    if len(set(values)) != 1:
        _fail("terminal_loss", "terminal LOST_SAMPLES representations disagree", values=values)
    return values[0]


def _run_to_file(argv: list[str], env: Mapping[str, str], output: Path, stderr: Path, *, timeout: int = 600) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        out_handle = output.open("xb")
        err_handle = stderr.open("xb")
    except FileExistsError as exc:
        _fail("output_exists", "refusing to overwrite offline decoder output", path=str(exc.filename))
    try:
        with out_handle, err_handle:
            completed = subprocess.run(argv, stdin=subprocess.DEVNULL, stdout=out_handle, stderr=err_handle, env=dict(env), timeout=timeout, check=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        _fail("offline_command", "offline perf command did not complete", argv=argv, error=str(exc))
    if completed.returncode != 0:
        tail = stderr.read_text(encoding="utf-8", errors="replace")[-2000:] if stderr.is_file() else ""
        _fail("offline_command", "offline perf command failed", argv=argv, returncode=completed.returncode, stderr_tail=tail)


def _validate_evlist(text: str, *, expected_period: int, expected_stack_cap: int) -> dict[str, Any]:
    """Validate ``perf evlist -v`` as the authoritative event contract.

    ``perf report --header-only`` on the upgraded profiler can emit a known
    ``Invalid HEADER_EVENT_DESC`` warning and omit its event descriptor.  The
    evlist-v record is parsed independently and therefore carries the actual
    period/sample/task attributes used for admission; the header warning is
    retained as descriptive stderr evidence by the caller.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) != 1:
        _fail("evlist_contract", "perf evlist -v must contain exactly one event line", nonempty_lines=len(lines))
    line = lines[0]
    if not line.startswith("instructions:u:"):
        _fail("evlist_contract", "event is not exact instructions:u", line=line[:300])
    required = (
        "type: 4 (cpu)", "size: 144", "config: 0xc0 (instructions)",
        f"{{ sample_period, sample_freq }}: {expected_period}",
        "sample_type: IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
        "read_format: ID|LOST", "disabled: 1", "exclude_kernel: 1", "exclude_hv: 1",
        "mmap: 1", "comm: 1", "enable_on_exec: 1", "task: 1", "mmap_data: 1",
        "sample_id_all: 1", "exclude_callchain_kernel: 1", "exclude_callchain_user: 1",
        "mmap2: 1", "comm_exec: 1", "ksymbol: 1", "bpf_event: 1", "build_id: 1",
        "sample_regs_user: 0xff0fff", f"sample_stack_user: {expected_stack_cap}",
    )
    # ``in`` is too permissive here: it accepts 10,000,000 as 1,000,000 and
    # 32,7680 as 32,768, and it also accepts a required value followed by an
    # extra flag.  The evlist output is a comma-delimited record (apart from
    # the comma inside the sample-period key), so require an attribute-token
    # boundary on both sides.  The first attribute follows the event's colon;
    # subsequent attributes follow a comma.
    def exact_attribute(token: str) -> bool:
        pattern = rf"(?:^|,\s*|:\s*){re.escape(token)}(?=,\s*|$)"
        return re.search(pattern, line) is not None

    missing = [token for token in required if not exact_attribute(token)]
    if missing:
        _fail("evlist_contract", "perf evlist -v is missing an exact event attribute", missing=missing, line=line[:500])

    # Reject duplicate known keys.  This keeps a line with a valid prefix plus
    # a contradictory second value from passing the exact-token checks.
    known_keys = (
        "type", "size", "config", "{ sample_period, sample_freq }", "sample_type",
        "read_format", "disabled", "exclude_kernel", "exclude_hv", "mmap", "comm",
        "enable_on_exec", "task", "mmap_data", "sample_id_all",
        "exclude_callchain_kernel", "exclude_callchain_user", "mmap2", "comm_exec",
        "ksymbol", "bpf_event", "build_id", "sample_regs_user", "sample_stack_user",
        "inherit",
    )
    key_counts: dict[str, int] = {}
    for key in known_keys:
        key_pattern = rf"(?:^|,\s*|:\s*){re.escape(key)}\s*:"
        count = len(re.findall(key_pattern, line))
        if count > 1:
            _fail("evlist_contract", "perf evlist -v repeats an event attribute", attribute=key, count=count)
        if count:
            key_counts[key] = count

    inherit = re.search(r"(?:^|,\s*|:\s*)inherit:\s*(\d+)(?=,\s*|$)", line)
    if inherit is not None and inherit.group(1) != "0":
        _fail("evlist_contract", "perf evlist -v reports inherit enabled", inherit=inherit.group(1))
    return {
        "schema": "litchi.performance.0977.evlist_contract.v1",
        "status": "pass",
        "event": "instructions:u",
        "period": expected_period,
        "stack_cap": expected_stack_cap,
        "inherit": int(inherit.group(1)) if inherit else 0,
        "inherit_field_absent": inherit is None,
        "line": line,
        "required_attributes": list(required),
        "authority": "perf evlist -v; report header descriptor is retained separately and is not admission authority",
    }


def _header_metadata(path: Path) -> dict[str, Any]:
    """Record the report-header warning without treating it as evlist data."""
    text = path.read_text(encoding="utf-8", errors="replace") if path.is_file() else ""
    warnings = [line.strip() for line in text.splitlines() if "Invalid HEADER_EVENT_DESC" in line]
    return {
        "schema": "litchi.performance.0977.perf_header_metadata.v1",
        "stderr": _file_binding(path),
        "known_invalid_header_event_desc_warning": bool(warnings),
        "warnings": warnings,
        "authority": "evlist-v; header warning is descriptive and not silently ignored",
    }


def _base_env() -> dict[str, str]:
    env = os.environ.copy()
    env["LC_ALL"] = "C"
    env["LANG"] = "C"
    env["TZ"] = "UTC"
    env.pop("LD_PRELOAD", None)
    for key in list(env):
        if key.startswith("LITCHI_0973_") or key.startswith("LITCHI_0976_"):
            env.pop(key, None)
    return env


def _decoder_env(binary: Path, shim: Path, layout: Mapping[str, Any], log_path: Path) -> dict[str, str]:
    env = _base_env()
    env.update({
        "LITCHI_0973_SHIM_MODE": "corrected",
        "LITCHI_0973_SHIM_LOG": str(log_path),
        "LITCHI_0973_TARGET_PATH": str(binary),
        "LITCHI_0973_TARGET_NAME": binary.name,
        "LITCHI_0973_CORRECTED_BASE": str(layout["corrected_base"]),
        "LITCHI_0973_PERF_EXPECTED_START": str(layout["perf_expected_start"]),
        "LITCHI_0976_TARGET_BUILD_ID": BASELINE_BUILD_ID,
        "LITCHI_0976_TARGET_SHA256": BASELINE_BINARY_SHA256,
        "LITCHI_0976_EXPECTED_DELTA": hex(EXPECTED_DELTA),
        "LD_PRELOAD": str(shim),
    })
    return env


def _parse_mmap_dump(path: Path, binary: Path, build_id: str, expected_pid: int) -> tuple[list[dict[str, Any]], list[dict[str, int]]]:
    rows: list[dict[str, Any]] = []
    forks: list[dict[str, int]] = []
    mmap_re = PARSER._MMAP2
    with path.open("r", encoding="utf-8", errors="strict") as stream:
        for line_number, line in enumerate(stream, 1):
            fork = FORK_RE.search(line)
            if fork is not None:
                forks.append({"child_pid": int(fork.group(1)), "child_tid": int(fork.group(2)), "parent_pid": int(fork.group(3)), "parent_tid": int(fork.group(4))})
            match = mmap_re.search(line)
            if match is None:
                continue
            row_path = match.group("path")
            if row_path != str(binary):
                continue
            build_match = re.search(r"<([0-9a-fA-F]+)>", match.group("meta"))
            row = {
                "line": line_number,
                "pid": int(match.group("pid")),
                "start": int(match.group("start"), 16),
                "length": int(match.group("length"), 16),
                "offset": int(match.group("offset"), 0),
                "meta": match.group("meta"),
                "build_id": build_match.group(1).lower() if build_match else None,
                "permissions": match.group("perm"),
                "path": row_path,
            }
            if row["pid"] != expected_pid:
                _fail("layout_contract", "target MMAP2 PID differs from launch PID", pid=row["pid"], expected=expected_pid)
            if row["build_id"] != build_id.lower():
                _fail("layout_contract", "target MMAP2 BuildID differs from retained baseline", actual=row["build_id"], expected=build_id.lower())
            rows.append(row)
    if not rows:
        _fail("layout_contract", "MMAP2 dump has no exact target rows", path=str(path))
    return rows, forks


def _derive_layout(perf_data: Path, mmap_dump: Path, binary: Path, expected_pid: int, build_id: str) -> dict[str, Any]:
    elf = PARSER._parse_elf_loads(binary)
    rows, forks = _parse_mmap_dump(mmap_dump, binary, build_id, expected_pid)
    selected: list[dict[str, Any]] = []
    for load in elf["pt_loads"]:
        candidates = [row for row in rows if row["offset"] == load["offset_page"] and PARSER._mapping_covers_load(load, row)]
        if not candidates:
            _fail("layout_contract", "missing target MMAP2 row for PT_LOAD", offset=hex(load["offset_page"]))
        biases = {row["start"] - load["vaddr_page"] for row in candidates}
        if len(biases) != 1:
            _fail("layout_contract", "PT_LOAD/MMAP2 rows imply inconsistent load biases", offset=hex(load["offset_page"]))
        selected.append({"load": load, "mmap": candidates[0]})
    biases = {item["mmap"]["start"] - item["load"]["vaddr_page"] for item in selected}
    if len(biases) != 1:
        _fail("layout_contract", "target PT_LOAD rows do not imply one load bias")
    corrected_base = next(iter(biases))
    exec_loads = [item for item in selected if item["load"]["flags"] & 1]
    if len(exec_loads) != 1:
        _fail("layout_contract", "expected exactly one executable PT_LOAD")
    executable = exec_loads[0]
    executable_maps = [row for row in rows if "x" in row["permissions"][:3]]
    if len(executable_maps) != 1 or executable_maps[0] != executable["mmap"]:
        _fail("layout_contract", "executable PT_LOAD does not match one target RX map")
    perf_expected_start = executable["mmap"]["start"] - executable["mmap"]["offset"]
    if perf_expected_start <= corrected_base or perf_expected_start - corrected_base != EXPECTED_DELTA:
        _fail("layout_contract", "derived perf/comparator base gap is not 0x1000", corrected_base=hex(corrected_base), perf_expected_start=hex(perf_expected_start))
    return {
        "schema": "litchi.performance.0973.target_layout.v2",
        "binary": _file_binding(binary),
        "perf_data": _file_binding(perf_data),
        "target_mmap2": [item["mmap"] for item in selected],
        "executable_mapping": executable["mmap"],
        "corrected_base": hex(corrected_base),
        "perf_expected_start": hex(perf_expected_start),
        "elf": elf,
        "process_tree": {"fork_records": forks, "unexpected_children": len(forks)},
        "formulas": {
            "corrected_base": "selected_mmap.start - page_down(PT_LOAD.p_vaddr); equal for every PT_LOAD",
            "perf_expected_start": "executable_selected_mmap.start - executable_selected_mmap.offset",
        },
        "perf_mmap_dump": {
            "argv": [str(PERF_PATH), "script", "-D", "-i", str(perf_data)],
            "returncode": 0,
            "sha256": _sha256(mmap_dump),
            "bytes": mmap_dump.stat().st_size,
            "streamed_once": True,
        },
    }


def _scan_raw_meta(path: Path, expected_pid: int, binary: Path) -> dict[str, Any]:
    raw_sum = 0
    raw_records = 0
    task_lines: list[str] = []
    tokens = ("PERF_RECORD_FORK", "PERF_RECORD_CLONE", "PERF_RECORD_EXIT", "PERF_RECORD_THREAD_MAP", "PERF_RECORD_CPU_MAP", "PERF_RECORD_COMM")
    with path.open("r", encoding="utf-8", errors="strict") as stream:
        for line in stream:
            match = TARGET_LOST_RE.search(line)
            if match is not None:
                raw_sum += int(match.group(1))
                raw_records += 1
            if any(token in line for token in tokens):
                task_lines.append(line.rstrip("\n"))
    try:
        task = ADAPTER.parse_task_stream("\n".join(task_lines), expected_pid=expected_pid, binary_name=binary.name)
    except Exception as exc:  # noqa: BLE001
        _fail("task_stream", "retained 0974 task parser rejected raw lifecycle stream", error=str(exc))
    if task.get("status") != "pass":
        _fail("task_stream", "raw lifecycle stream failed", errors=task.get("errors"))
    return {
        "raw_lost_sum": raw_sum,
        "raw_lost_record_count": raw_records,
        "raw_lost_scope": "sum PERF_RECORD_LOST fields only; terminal LOST_SAMPLES is separate",
        "task_stream": task,
    }


def _stream_perf_script(path: Path, *, marker: str, marker_dso: str, expected_pid: int, expected_cpu: int, expected_period: int) -> dict[str, Any]:
    """Parse one physical perf-script stream in O(N) space bounded by one stack."""
    header_re = SAMPLE_HEADER_RE
    frame_re = FRAME_RE
    current: dict[str, Any] | None = None
    sample_count = 0
    total_period = 0
    marker_count = 0
    marker_period = 0
    unresolved = 0
    explicit_truncation = 0
    empty = 0
    lost = 0
    physical: dict[str, int] = {}
    inclusive: dict[str, int] = {}
    deepest: dict[str, int] = {}
    root_owner: dict[str, int] = {}
    proxy: dict[str, int] = {}
    unknown_period = 0
    unknown_samples = 0
    marker_no_resolved = 0

    def owner(frame: Mapping[str, Any]) -> str:
        return f"{frame['dso']}::{frame['symbol_base']}"

    def finish() -> None:
        nonlocal current, sample_count, total_period, marker_count, marker_period
        nonlocal unresolved, empty, unknown_period, unknown_samples, marker_no_resolved
        if current is None:
            return
        frames = current["frames"]
        period = current["period"]
        sample_count += 1
        total_period += period
        unresolved += sum(1 for frame in frames if not frame["resolved"])
        if not frames:
            empty += 1
        is_marker = any(frame["symbol_base"] == marker and frame["dso"] == marker_dso for frame in frames)
        if is_marker:
            marker_count += 1
            marker_period += period
            resolved = [frame for frame in frames if frame["resolved"]]
            if not frames or not frames[0]["resolved"]:
                unknown_period += period
                unknown_samples += 1
            else:
                key = owner(frames[0])
                physical[key] = physical.get(key, 0) + period
            if resolved:
                first = owner(resolved[0])
                last = owner(resolved[-1])
                proxy[first] = proxy.get(first, 0) + period
                deepest[first] = deepest.get(first, 0) + period
                root_owner[last] = root_owner.get(last, 0) + period
                for key in {owner(frame) for frame in resolved}:
                    inclusive[key] = inclusive.get(key, 0) + period
            else:
                marker_no_resolved += 1
        current = None

    try:
        stream = path.open("r", encoding="utf-8", errors="strict")
    except OSError as exc:
        _fail("perf_script", "cannot open decoded perf script", path=str(path), error=str(exc))
    with stream:
        for line_number, raw in enumerate(stream, 1):
            line = raw.rstrip("\n")
            stripped = line.strip()
            if not stripped:
                finish()
                continue
            low = stripped.lower()
            lost_match = re.search(r"(?:lost|dropped)[^0-9]*(\d+)", low)
            if lost_match is not None and ("lost" in low or "drop" in low):
                lost += int(lost_match.group(1))
                continue
            if stripped in {"...", "[stack truncated]", "[truncated]"} or "stack truncated" in low:
                explicit_truncation += 1
                continue
            header = header_re.match(line)
            if header is not None:
                finish()
                event = header.group("event").strip()
                if event != "instructions:u":
                    _fail("perf_script", "unexpected perf event", line=line_number, event=event)
                pid = int(header.group("pid"))
                tid = header.group("tid")
                cpu_text = header.group("cpu")
                period = int(header.group("period"))
                if pid != expected_pid or tid is not None and int(tid) != expected_pid:
                    _fail("perf_script", "sample PID/TID escaped expected launch", line=line_number, pid=pid, tid=tid, expected=expected_pid)
                if cpu_text is None or int(cpu_text) != expected_cpu:
                    _fail("perf_script", "sample CPU escaped expected affinity", line=line_number, cpu=cpu_text, expected=expected_cpu)
                if period != expected_period or period <= 0:
                    _fail("perf_script", "sample period changed", line=line_number, period=period, expected=expected_period)
                current = {"period": period, "frames": []}
                continue
            if current is not None and line[:1].isspace():
                frame = frame_re.match(line)
                if frame is None:
                    _fail("perf_script", "malformed physical frame", line=line_number, text=line[:300])
                symbol = frame.group("symbol").strip()
                base, _, offset = symbol.partition("+")
                current["frames"].append({
                    "ip": frame.group("ip").lower(),
                    "symbol": symbol,
                    "symbol_base": base.strip(),
                    "offset": offset or None,
                    "dso": frame.group("dso").strip(),
                    "resolved": base.strip() not in {"[unknown]", "unknown", ""} and frame.group("dso").strip() not in {"[unknown]", ""},
                })
                continue
            _fail("perf_script", "malformed perf-script line", line=line_number, text=line[:300])
    finish()
    if sample_count == 0 and lost == 0:
        _fail("perf_script", "decoded script has no samples or loss record", path=str(path))
    classified = sum(physical.values()) + unknown_period
    if classified != marker_period:
        _fail("owner_conservation", "physical marker-period conservation failed", classified=classified, marker_period=marker_period)
    return {
        "sample_count": sample_count,
        "total_period": total_period,
        "marker": marker,
        "marker_sample_count": marker_count,
        "marker_period": marker_period,
        "marker_period_share": marker_period / total_period if total_period else None,
        "unresolved_frames": unresolved,
        "explicit_stack_truncated_count": explicit_truncation,
        "empty_callchain_count": empty,
        "lost_events": lost,
        "inclusive_marker_ancestor_period": inclusive,
        "deepest_resolved_proxy_period": proxy,
        "physical_leaf_owner_period": physical,
        "physical_leaf_unknown_period": unknown_period,
        "physical_leaf_unknown_samples": unknown_samples,
        "marker_unknown_physical_leaf_samples": unknown_samples,
        "marker_no_resolved_frame_samples": marker_no_resolved,
        "physical_leaf_scope": "marker-containing samples only; frame-0 physical leaf or explicit unknown",
        "physical_leaf_classified_period": classified,
        "physical_leaf_conservation": {"marker_period": marker_period, "classified_period": classified, "passed": True, "scope": "marker-containing samples only"},
    }


def _stream_shim_log(path: Path, layout: Mapping[str, Any], binary: Path, *, expected_build_id: str) -> dict[str, Any]:
    counts = {"rows": 0, "reports": 0, "infos": 0, "target_reports": 0, "target_infos": 0, "cache_stored": 0, "reused_existing": 0, "adjusted": 0, "fallback": 0, "result_module_null": 0}
    corrected_base = str(layout["corrected_base"])
    expected_start = str(layout["perf_expected_start"])
    if not path.is_file():
        _fail("shim_log", "corrected shim log is missing", path=str(path))
    with path.open("r", encoding="utf-8", errors="strict") as stream:
        for line_number, line in enumerate(stream, 1):
            if not line.strip():
                continue
            try:
                row = json.loads(line)
            except json.JSONDecodeError as exc:
                _fail("shim_log", "shim JSONL row is malformed", line=line_number, error=str(exc))
            if not isinstance(row, Mapping):
                _fail("shim_log", "shim JSONL row is not an object", line=line_number)
            schema = str(row.get("schema", ""))
            counts["rows"] += 1
            if schema == "litchi.performance.0976.dwfl_report.final":
                counts["reports"] += 1
                exact = row.get("path_match") is True and row.get("name_match") is True
                partial = row.get("path_match") is True or row.get("name_match") is True
                if row.get("cache_stored") is True:
                    counts["cache_stored"] += 1
                if row.get("reused_existing") is True:
                    counts["reused_existing"] += 1
                if row.get("result_module") in (None, "0x0", "(nil)"):
                    # A null result on an unrelated DSO is an ordinary DWFL
                    # lookup outcome.  It is a failure only for an exact
                    # admitted target row, where the final-v2 identity gate
                    # promises a retained module.
                    if exact:
                        counts["result_module_null"] += 1
                if exact:
                    counts["target_reports"] += 1
                    if row.get("admission") is not True:
                        # ``forwarded_base != original_base`` is the expected
                        # correction on every target row.  Fallback means an
                        # exact target row was not admitted, not that its base
                        # changed.
                        counts["fallback"] += 1
                    required = {
                        "config_valid": True,
                        "admission": True,
                        "caller_is_perf": True,
                        "caller_file": "/usr/bin/perf",
                        "fd": -1,
                        "add_p_vaddr": False,
                        "forwarded_base": corrected_base,
                        "identity_valid": True,
                    }
                    for key, value in required.items():
                        if row.get(key) != value:
                            _fail("shim_log", "exact target report failed final_v2 identity/admission check", line=line_number, field=key, actual=row.get(key), expected=value)
                    if row.get("name") != binary.name or row.get("file_name") != str(binary):
                        _fail("shim_log", "exact target report path/name differs from binary", line=line_number)
                    if row.get("original_base") not in {expected_start, corrected_base}:
                        _fail("shim_log", "exact target report has an unexpected input base", line=line_number, original=row.get("original_base"))
                    if row.get("cache_stored") is True:
                        if row.get("reused_existing") is True:
                            _fail("shim_log", "cache store and reuse overlap on one target row", line=line_number)
                    elif row.get("reused_existing") is not True:
                        _fail("shim_log", "exact target report is neither first cache store nor reuse", line=line_number)
                elif partial:
                    if row.get("forwarded_base") != row.get("original_base") or row.get("admission") is True or row.get("cache_stored") is True or row.get("reused_existing") is True:
                        _fail("shim_log", "partial target report was modified", line=line_number)
            elif schema == "litchi.performance.0976.dwfl_info.final":
                counts["infos"] += 1
                target = row.get("target_match") is True
                if target:
                    counts["target_infos"] += 1
                if row.get("start_adjusted") is True:
                    counts["adjusted"] += 1
                    if not target or row.get("caller_is_perf") is not True or row.get("caller_file") != "/usr/bin/perf" or row.get("before_start") != corrected_base or row.get("after_start") != expected_start or row.get("config_valid") is not True or row.get("current_path_valid") is not True or row.get("current_range_valid") is not True:
                        _fail("shim_log", "invalid final_v2 info adjustment", line=line_number)
                elif target and row.get("before_start") != row.get("after_start"):
                    _fail("shim_log", "target info changed without adjustment flag", line=line_number)
            else:
                _fail("shim_log", "unexpected final_v2 shim log schema", line=line_number, schema=schema)
    if counts["rows"] == 0 or counts["target_reports"] == 0:
        _fail("shim_log", "final_v2 shim log has no exact target report rows", counts=counts)
    if counts["cache_stored"] != 1 or counts["reused_existing"] != counts["target_reports"] - 1:
        _fail("shim_log", "final_v2 cache lifecycle is not one store followed by reuse", counts=counts)
    if counts["fallback"] != 0 or counts["result_module_null"] != 0:
        _fail("shim_log", "final_v2 target reports contain fallback or null modules", counts=counts)
    return counts


def _artifact_map(child_dir: Path) -> dict[str, Path]:
    stem = child_dir.name
    result = {
        "perf_data": child_dir / f"{stem}.perf.data",
        "command": child_dir / f"{stem}.command.json",
        "launch": child_dir / f"{stem}.launch.json",
        "terminal": child_dir / f"{stem}.terminal.json",
        "terminal_report": child_dir / f"{stem}.terminal-report.txt",
        "report": child_dir / f"{stem}.report.json",
        "perf_header": _output_path(child_dir, "perf-header.txt"),
        "perf_evlist": _output_path(child_dir, "perf-evlist-v.txt"),
        "perf_evlist_stderr": _output_path(child_dir, "perf-evlist-v.stderr"),
        "perf_buildid": _output_path(child_dir, "perf-buildid.txt"),
        "perf_mmap_dump": _output_path(child_dir, "perf-mmap-dump.txt"),
        "perf_mmap_stderr": _output_path(child_dir, "perf-mmap-dump.stderr"),
        "terminal_report_stderr": _output_path(child_dir, "terminal-report.stderr"),
        "target_layout": _output_path(child_dir, "target-layout.json"),
        # The reducer accepts both the descriptive decoder name and the
        # concise runner alias.  They intentionally bind the same immutable
        # MMAP2/layout files rather than creating duplicate evidence.
        "layout": _output_path(child_dir, "target-layout.json"),
        "raw_stack": _output_path(child_dir, "raw-stack.json"),
        "raw_task": _output_path(child_dir, "perf-mmap-dump.txt"),
        "corrected_script": _output_path(child_dir, "corrected.perf-script-no-inline.txt"),
        "corrected_stderr": _output_path(child_dir, "corrected.perf.stderr"),
        "corrected_log": _output_path(child_dir, "corrected.dwfl.jsonl"),
        "corrected_command": _output_path(child_dir, "corrected.command.json"),
    }
    return result


def _ensure_terminal_report(
    terminal: Mapping[str, Any], child_dir: Path, perf_data: Path, paths: Mapping[str, Path], env: Mapping[str, str]
) -> dict[str, Any]:
    """Make the decoder's exact LOST_SAMPLES input explicit and immutable.

    The normal runner has already captured ``terminal-report.txt``.  A
    terminal-only caller is still supported: the decoder performs this one
    offline ``perf report`` before applying the loss gate, recording the same
    bounded path locally.  The packet keeps the original terminal receipt;
    this copy is only the decoder's parse input.
    """
    result = dict(terminal)
    supplied = result.get("terminal_report")
    supplied_path = Path(str(supplied)).absolute() if isinstance(supplied, str) else None
    if supplied_path is not None and supplied_path.is_file():
        return result
    report_path = paths["terminal_report"]
    if not report_path.is_file():
        _run_to_file(
            # Current perf's report subcommand emits the exact
            # ``Total Lost Samples`` trailer without the script-only
            # ``--show-lost-events`` option (which report rejects).
            [str(PERF_PATH), "report", "--stdio", "--percent-limit", "100", "-i", str(perf_data)],
            env, report_path, paths["terminal_report_stderr"], timeout=300,
        )
    result["terminal_report"] = str(report_path.resolve())
    result["terminal_report_returncode"] = 0
    return result


def _bindings(paths: Mapping[str, Path], *, required_inputs: set[str]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for name, path in paths.items():
        if name in required_inputs or path.is_file():
            result[name] = _file_binding(path)
    return result


def _control_paths(child_dir: Path) -> dict[str, Path]:
    """Return the bounded control inputs/outputs without target-layout work."""
    stem = child_dir.name
    return {
        "perf_data": child_dir / f"{stem}.perf.data",
        "command": child_dir / f"{stem}.command.json",
        "terminal": child_dir / f"{stem}.terminal.json",
        "terminal_report": child_dir / f"{stem}.terminal-report.txt",
        "terminal_report_stderr": _output_path(child_dir, "terminal-report.stderr"),
        "control_target": child_dir / f"{stem}.control-target.json",
        "perf_header": _output_path(child_dir, "perf-header.txt"),
        "perf_header_stderr": _output_path(child_dir, "perf-header.stderr"),
        "perf_evlist": _output_path(child_dir, "perf-evlist-v.txt"),
        "perf_evlist_stderr": _output_path(child_dir, "perf-evlist-v.stderr"),
        "perf_mmap_dump": _output_path(child_dir, "perf-mmap-dump.txt"),
        "perf_mmap_stderr": _output_path(child_dir, "perf-mmap-dump.stderr"),
        "raw_stack": _output_path(child_dir, "raw-stack.json"),
        "control_summary": _output_path(child_dir, "control-summary.json"),
    }


def _validate_control_inputs(
    plan: Mapping[str, Any], child_dir: Path, terminal: Mapping[str, Any], control: Mapping[str, Any]
) -> tuple[Path, int, int, int]:
    if terminal.get("returncode") not in (None, 0) or terminal.get("timed_out") is True or terminal.get("launch_error") is True:
        _fail("control_terminal", "control terminal receipt records a failed capture")
    if terminal.get("native_ld_preload") not in (None, "absent"):
        _fail("control_terminal", "control capture had LD_PRELOAD", value=terminal.get("native_ld_preload"))
    if control.get("status") != "pass" or control.get("control_only") is not True or control.get("native_workload_started") is not False:
        _fail("control_receipt", "control target receipt is not a passing no-workload receipt")
    if control.get("target_cpu") != DEFAULT_CPU or control.get("recorder_cpu") != 30:
        _fail("control_receipt", "control target/recorder CPU contract drift")
    target = control.get("target")
    executable = control.get("target_executable")
    if not isinstance(target, Mapping) or type(target.get("pid")) is not int or target["pid"] <= 0:
        _fail("control_receipt", "control target PID is missing")
    if not isinstance(executable, Mapping) or not isinstance(executable.get("path"), str) or not isinstance(executable.get("sha256"), str):
        _fail("control_receipt", "control target executable binding is missing")
    python_row = plan.get("python")
    if not isinstance(python_row, Mapping) or not isinstance(python_row.get("path"), str):
        _fail("control_plan", "plan Python binding is missing")
    python_path = Path(str(python_row["path"])).absolute()
    if Path(str(executable["path"])).absolute() != python_path:
        _fail("control_receipt", "control executable differs from bound Python", actual=executable.get("path"), expected=str(python_path))
    actual = _file_binding(python_path)
    if actual["sha256"] != str(python_row.get("sha256", actual["sha256"])) or actual["sha256"] != executable["sha256"]:
        _fail("control_receipt", "control Python executable hash differs from plan", actual=actual["sha256"], expected=executable.get("sha256"))
    perf = plan.get("perf") if isinstance(plan.get("perf"), Mapping) else {}
    perf_path = Path(str(perf.get("path", PERF_PATH))).absolute()
    if perf_path != PERF_PATH.absolute() or _sha256(perf_path) != PERF_SHA256:
        _fail("control_plan", "control offline perf tool is not sealed")
    period = int(perf.get("period", DEFAULT_PERIOD))
    stack_cap = int(perf.get("stack_cap", DEFAULT_STACK_CAP))
    if period != DEFAULT_PERIOD or stack_cap != DEFAULT_STACK_CAP:
        _fail("control_plan", "control period/stack contract drift")
    return python_path, int(target["pid"]), period, stack_cap


def validate_control_child(
    plan: Mapping[str, Any], child_dir: Path, terminal: Mapping[str, Any], control_receipt: Mapping[str, Any]
) -> dict[str, Any]:
    """Validate the harmless control through the same offline raw boundary.

    This path intentionally does not derive the Office ELF layout, inspect the
    final DWFL shim, or apply the marker gate.  It only proves the recorder's
    header, single-PID raw sample fields, task lifecycle, and two independent
    loss representations for the control process.
    """
    child_dir = Path(child_dir).absolute()
    if not child_dir.is_dir():
        _fail("control_input", "control child directory is missing", child=str(child_dir))
    python_path, expected_pid, period, stack_cap = _validate_control_inputs(plan, child_dir, terminal, control_receipt)
    paths = _control_paths(child_dir)
    for name in ("perf_data", "command", "terminal", "control_target"):
        if not paths[name].is_file():
            _fail("control_input", "control artifact is missing", name=name, path=str(paths[name]))
    command = _load_json(paths["command"])
    if command.get("argv") != terminal.get("command"):
        _fail("control_command", "terminal command differs from immutable command receipt")
    argv = command.get("argv")
    if not isinstance(argv, list) or "--per-thread" not in argv or "-m" not in argv or "16M" not in argv or "--call-graph" not in argv or "dwarf,32768" not in argv:
        _fail("control_command", "control command sampling contract is incomplete")
    if any(token in argv for token in ("-a", "--all-cpus", "--inherit")):
        _fail("control_command", "control command contains forbidden system-wide/inherit mode")
    terminal_for_loss = _ensure_terminal_report(terminal, child_dir, paths["perf_data"], paths, _base_env())
    terminal_loss = _terminal_loss(terminal_for_loss, child_dir)
    terminal_text = paths["terminal_report"].read_text(encoding="utf-8", errors="strict")
    terminal_matches = TERMINAL_LOST_RE.findall(terminal_text)
    if len(terminal_matches) != 1 or int(terminal_matches[0]) != terminal_loss:
        _fail("control_terminal", "terminal report and receipt loss values disagree")

    base_env = _base_env()
    _run_to_file(
        [str(PERF_PATH), "evlist", "-v", "-i", str(paths["perf_data"])],
        base_env, paths["perf_evlist"], paths["perf_evlist_stderr"],
    )
    try:
        evlist_contract = _validate_evlist(
            paths["perf_evlist"].read_text(encoding="utf-8", errors="strict"),
            expected_period=period, expected_stack_cap=stack_cap,
        )
    except DecodeFailure:
        raise
    except Exception as exc:  # noqa: BLE001
        _fail("control_evlist", "control perf evlist contract failed", error=str(exc))
    _run_to_file(
        [str(PERF_PATH), "report", "--header-only", "--stdio", "-i", str(paths["perf_data"])],
        base_env, paths["perf_header"], paths["perf_header_stderr"],
    )
    header_metadata = _header_metadata(paths["perf_header_stderr"])
    try:
        header_contract = ADAPTER.parse_header_contract(paths["perf_header"].read_text(encoding="utf-8", errors="strict"))
    except Exception as exc:  # noqa: BLE001
        # The current perf upgrade can omit HEADER_EVENT_DESC and report the
        # exact known warning below.  ``evlist -v`` has already passed the
        # authoritative event contract, so retain this header failure as
        # descriptive provenance instead of treating the missing duplicate
        # descriptor as a second admission source.
        if not header_metadata.get("known_invalid_header_event_desc_warning"):
            _fail("control_header", "control perf header failed retained contract", error=str(exc))
        header_contract = {
            "schema": "litchi.performance.0977.header_contract.v1",
            "status": "unavailable",
            "authority": "evlist-v",
            "error": str(exc),
        }
    _run_to_file(
        [str(PERF_PATH), "script", "-D", "-i", str(paths["perf_data"])],
        base_env, paths["perf_mmap_dump"], paths["perf_mmap_stderr"],
    )
    raw_meta = _scan_raw_meta(paths["perf_mmap_dump"], expected_pid, python_path)
    try:
        raw_summary = RAW_SCANNER.scan_dump(
            paths["perf_mmap_dump"], perf_data=paths["perf_data"], binary=python_path,
            expected_pid=expected_pid, expected_cpu=DEFAULT_CPU, expected_period=period,
            stack_cap=stack_cap, layout=None,
        )
    except Exception as exc:  # noqa: BLE001
        _fail("control_raw_stack", "control raw sample scan failed", error=str(exc))
    raw_summary.update(raw_meta)
    if (
        raw_summary.get("sample_count") != raw_summary.get("sample_line_count")
        or raw_summary.get("sample_count", 0) <= 0
        or raw_summary.get("period_valid") is not True
        or raw_summary.get("pid_tid_cpu_valid") is not True
        or raw_summary.get("bounds_valid") is not True
        or raw_summary.get("violations") != []
    ):
        _fail("control_raw_stack", "control raw sample gate failed", summary={key: raw_summary.get(key) for key in ("sample_count", "sample_line_count", "period_valid", "pid_tid_cpu_valid", "bounds_valid", "violations")})
    if raw_summary.get("task_stream", {}).get("status") != "pass":
        _fail("control_task", "control lifecycle task stream failed", errors=raw_summary.get("task_stream", {}).get("errors"))
    try:
        ADAPTER.V1.require_zero_loss(
            {"raw_lost_sum": raw_summary.get("raw_lost_sum")},
            {"lost_samples_summary": {"record_type": "PERF_RECORD_LOST_SAMPLES", "lost_samples": terminal_loss}},
        )
    except Exception as exc:  # noqa: BLE001
        _fail("control_loss", "control raw/terminal loss gate failed", error=str(exc), raw_lost_sum=raw_summary.get("raw_lost_sum"), terminal_lost_samples=terminal_loss)
    _write_json_exclusive(paths["raw_stack"], raw_summary)
    summary = {
        "sample_count": raw_summary["sample_count"],
        "raw_lost_sum": int(raw_summary["raw_lost_sum"]),
        "terminal_lost_samples": int(terminal_loss),
        "cpu_gate": bool(raw_summary["pid_tid_cpu_valid"]),
        "task_gate": raw_summary["task_stream"].get("status") == "pass",
        "period": period,
        "scope": "harmless control only; no Office marker/ELF/DWFL attribution",
    }
    _write_json_exclusive(paths["control_summary"], summary)
    artifacts = _bindings(paths, required_inputs={"perf_data", "command", "terminal", "terminal_report", "control_target"})
    artifacts["raw_task"] = _file_binding(paths["perf_mmap_dump"])
    artifacts["raw-task.txt"] = artifacts["raw_task"]
    artifacts["terminal-summary.json"] = {"path": str(paths["control_summary"]), "bytes": paths["control_summary"].stat().st_size, "sha256": _sha256(paths["control_summary"])}
    return {
        "schema": "litchi.performance.0977.control_packet.v1",
        "status": "complete",
        "packet_id": child_dir.name,
        "control_only": True,
        "native_workload_started": False,
        "summary": summary,
        "raw_stack": raw_summary,
        "task_audit": raw_summary["task_stream"],
        "header_contract": header_contract,
        "evlist_contract": evlist_contract,
        "header_metadata": header_metadata,
        "control_receipt": dict(control_receipt),
        "terminal": dict(terminal),
        "command": command,
        "artifacts": artifacts,
        "gate": {"passed": True, "reasons": [], **summary},
    }


def decode_child(plan: Mapping[str, Any], child_dir: Path, terminal: Mapping[str, Any], launch: Mapping[str, Any]) -> dict[str, Any]:
    """Decode one child offline, or raise :class:`DecodeFailure`.

    The function never starts a workload or a native capture.  It invokes only
    ``perf`` subcommands against an already-retained perf.data file.
    """
    child_dir = Path(child_dir).absolute()
    if not child_dir.is_dir():
        _fail("child_input", "child directory is missing", child=str(child_dir))
    binary, shim, period, cpu, stack_cap, marker, marker_dso = _validate_plan(plan)
    expected_pid = _validate_launch(launch, child_dir, binary)
    paths = _artifact_map(child_dir)
    for name in ("perf_data", "command", "launch", "terminal", "report"):
        if not paths[name].is_file():
            _fail("missing_capture_artifact", "required runner artifact is missing", name=name, path=str(paths[name]))
    terminal_for_loss = _ensure_terminal_report(terminal, child_dir, paths["perf_data"], paths, _base_env())
    terminal_loss = _terminal_loss(terminal_for_loss, child_dir)
    report = _load_json(paths["report"])
    try:
        PARSER.validate_report(report)
    except Exception as exc:  # noqa: BLE001
        _fail("report_contract", "native measurement report failed retained validator", error=str(exc))

    base_env = _base_env()
    evlist_argv = [str(PERF_PATH), "evlist", "-v", "-i", str(paths["perf_data"])]
    _run_to_file(evlist_argv, base_env, paths["perf_evlist"], paths["perf_evlist_stderr"])
    try:
        evlist_contract = _validate_evlist(paths["perf_evlist"].read_text(encoding="utf-8", errors="strict"), expected_period=period, expected_stack_cap=stack_cap)
    except DecodeFailure:
        raise
    except Exception as exc:  # noqa: BLE001
        _fail("evlist_contract", "offline perf evlist contract failed", error=str(exc))
    header_argv = [str(PERF_PATH), "report", "--header-only", "--stdio", "-i", str(paths["perf_data"])]
    buildid_argv = [str(PERF_PATH), "buildid-list", "-i", str(paths["perf_data"])]
    mmap_argv = [str(PERF_PATH), "script", "-D", "-i", str(paths["perf_data"])]
    _run_to_file(header_argv, base_env, paths["perf_header"], _output_path(child_dir, "perf-header.stderr"))
    header_metadata = _header_metadata(_output_path(child_dir, "perf-header.stderr"))
    _run_to_file(buildid_argv, base_env, paths["perf_buildid"], _output_path(child_dir, "perf-buildid.stderr"))
    buildid_text = paths["perf_buildid"].read_text(encoding="utf-8", errors="strict")
    try:
        PARSER.check_buildid_export(buildid_text, binary_path=str(binary), build_id=BASELINE_BUILD_ID)
    except Exception as exc:  # noqa: BLE001
        _fail("buildid_export", "perf buildid-list lacks exact baseline path/build ID", error=str(exc))
    _run_to_file(mmap_argv, base_env, paths["perf_mmap_dump"], paths["perf_mmap_stderr"])
    layout = _derive_layout(paths["perf_data"], paths["perf_mmap_dump"], binary, expected_pid, BASELINE_BUILD_ID)
    _write_json_exclusive(paths["target_layout"], layout)

    raw_meta = _scan_raw_meta(paths["perf_mmap_dump"], expected_pid, binary)
    raw_summary = RAW_SCANNER.scan_dump(paths["perf_mmap_dump"], perf_data=paths["perf_data"], binary=binary, expected_pid=expected_pid, expected_cpu=cpu, expected_period=period, stack_cap=stack_cap, layout=layout)
    raw_summary.update(raw_meta)
    if raw_summary.get("sample_count") != raw_summary.get("sample_line_count") or raw_summary.get("period_valid") is not True or raw_summary.get("pid_tid_cpu_valid") is not True or raw_summary.get("bounds_valid") is not True or raw_summary.get("violations") != []:
        _fail("raw_stack", "retained raw STACK_USER scanner rejected child", summary={key: raw_summary.get(key) for key in ("sample_count", "sample_line_count", "period_valid", "pid_tid_cpu_valid", "bounds_valid", "violations")})
    try:
        ADAPTER.V1.require_zero_loss({"raw_lost_sum": raw_summary["raw_lost_sum"]}, {"lost_samples_summary": {"record_type": "PERF_RECORD_LOST_SAMPLES", "lost_samples": terminal_loss}})
    except Exception as exc:  # noqa: BLE001
        _fail("loss_gate", "raw and terminal loss gates did not both pass; corrected decode skipped", raw_lost_sum=raw_summary.get("raw_lost_sum"), terminal_lost_samples=terminal_loss, error=str(exc))
    _write_json_exclusive(paths["raw_stack"], raw_summary)

    log_path = paths["corrected_log"]
    corrected_argv = [str(PERF_PATH), "script", "--show-lost-events", "--no-inline", "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso", "-i", str(paths["perf_data"])]
    decoder_env = _decoder_env(binary, shim, layout, log_path)
    command_env = {key: decoder_env[key] for key in ("LC_ALL", "LANG", "TZ", "LD_PRELOAD", "LITCHI_0973_SHIM_MODE", "LITCHI_0973_SHIM_LOG", "LITCHI_0973_TARGET_PATH", "LITCHI_0973_TARGET_NAME", "LITCHI_0973_CORRECTED_BASE", "LITCHI_0973_PERF_EXPECTED_START", "LITCHI_0976_TARGET_BUILD_ID", "LITCHI_0976_TARGET_SHA256", "LITCHI_0976_EXPECTED_DELTA")}
    command = _write_command(paths["corrected_command"], corrected_argv, command_env)
    _run_to_file(corrected_argv, decoder_env, paths["corrected_script"], paths["corrected_stderr"])
    corrected = _stream_perf_script(paths["corrected_script"], marker=marker, marker_dso=marker_dso, expected_pid=expected_pid, expected_cpu=cpu, expected_period=period)
    if corrected["sample_count"] != raw_summary["sample_count"] or corrected["total_period"] != raw_summary["sample_count"] * period:
        _fail("perf_conservation", "corrected decoded sample/period count disagrees with raw sample ledger", corrected_sample_count=corrected["sample_count"], raw_sample_count=raw_summary["sample_count"], corrected_period=corrected["total_period"], expected_period=raw_summary["sample_count"] * period)
    shim_stats = _stream_shim_log(log_path, layout, binary, expected_build_id=BASELINE_BUILD_ID)
    # These provenance values are explicitly allowed by the reducer.  All
    # owner/count fields remain the strict parser-equivalent keys above; the
    # reducer independently reparses the script and checks these bindings.
    corrected.update({
        "raw_lost_sum": int(raw_summary["raw_lost_sum"]),
        "terminal_lost_samples": int(terminal_loss),
        "decoder_schema": "litchi.performance.0977.corrected_summary.v1",
        "decoder_version": "focuseddecode-0977-v1",
        "corrected_script_sha256": _sha256(paths["corrected_script"]),
        "corrected_log_sha256": _sha256(log_path),
        "summary_scope": "corrected final-v2 no-inline stream; samples omitted; marker-containing physical leaf buckets",
    })
    marker_gate = plan.get("child_gate") if isinstance(plan.get("child_gate"), Mapping) else {}
    minimum_marker = int(marker_gate.get("min_marker_samples", 100))
    reasons: list[str] = []
    if corrected["marker_sample_count"] < minimum_marker:
        reasons.append("marker_sample_count_lt_minimum")
    if corrected["marker_period_share"] is None or corrected["marker_period_share"] <= 0:
        reasons.append("marker_period_share_nonpositive")
    if corrected["lost_events"] != 0:
        reasons.append("perf_script_lost_events_nonzero")
    if corrected["empty_callchain_count"] != 0:
        reasons.append("empty_callchain_count_nonzero")
    if corrected["physical_leaf_conservation"].get("passed") is not True:
        reasons.append("physical_leaf_conservation_failed")
    if reasons:
        _fail("child_gate", "corrected child quality gate failed", reasons=reasons, corrected={key: corrected.get(key) for key in ("sample_count", "marker_sample_count", "marker_period", "marker_period_share", "lost_events", "empty_callchain_count")})

    artifacts = _bindings(paths, required_inputs={"perf_data", "command", "launch", "terminal", "terminal_report", "report"})
    artifacts["perf_header_stderr"] = _file_binding(_output_path(child_dir, "perf-header.stderr"))
    artifacts["perf_evlist_stderr"] = _file_binding(paths["perf_evlist_stderr"])
    artifacts["perf_buildid_stderr"] = _file_binding(_output_path(child_dir, "perf-buildid.stderr"))
    artifacts["corrected_log"] = _file_binding(log_path)
    packet = {
        "schema": "litchi.performance.0977.child_packet.v1",
        "status": "complete",
        "packet_id": child_dir.name,
        "child_number": int(re.search(r"(\d+)$", child_dir.name).group(1)) if re.search(r"(\d+)$", child_dir.name) else None,
        "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "launch": dict(launch),
        "terminal": dict(terminal),
        "report": report,
        "layout": layout,
        "shim_config": {
            "target_path": str(binary),
            "target_name": binary.name,
            "target_build_id": BASELINE_BUILD_ID,
            "target_sha256": BASELINE_BINARY_SHA256,
            "corrected_base": layout["corrected_base"],
            "perf_expected_start": layout["perf_expected_start"],
            "expected_delta": hex(EXPECTED_DELTA),
            "mode": "corrected",
            "shared_object": _file_binding(shim),
            "source": _file_binding(FINAL_SHIM_SOURCE),
            "version_script": _file_binding(FINAL_SHIM_MAP),
        },
        "perf_contract": {"evlist": evlist_contract, "header": header_metadata},
        "raw_stack": raw_summary,
        "task_audit": raw_summary["task_stream"],
        "artifacts": artifacts,
        # The 0977 reducer requires this exact corrected-only shape.  Raw
        # conservation and decoder provenance live in the bounded summary's
        # explicitly allowed extra fields above.
        "perf": {"corrected": corrected},
        "summary": {
            "sample_count": corrected["sample_count"],
            "marker_sample_count": corrected["marker_sample_count"],
            "marker_period": corrected["marker_period"],
            "marker_period_share": corrected["marker_period_share"],
            "empty_callchain_count": corrected["empty_callchain_count"],
            "raw_lost_sum": raw_summary["raw_lost_sum"],
            "terminal_lost_samples": terminal_loss,
            "cpu_gate": True,
            "task_gate": True,
            "scope": "runner gate projection of corrected summary; full owner maps remain under perf.corrected",
        },
        "gate": {
            "passed": True,
            "reasons": [],
            "minimum_marker_samples": minimum_marker,
            "raw_lost_sum": raw_summary["raw_lost_sum"],
            "terminal_lost_samples": terminal_loss,
            "empty_callchain_count": corrected["empty_callchain_count"],
            "unresolved_frames": corrected["unresolved_frames"],
            "explicit_stack_truncated_count": corrected["explicit_stack_truncated_count"],
            "shim_log": shim_stats,
        },
    }
    return packet


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--child", type=Path, required=True)
    args = parser.parse_args(argv)
    plan = _load_json(args.plan)
    child = args.child.absolute()
    terminal = _load_json(_input_path(child, "terminal.json"))
    launch = _load_json(_input_path(child, "launch.json"))
    try:
        packet = decode_child(plan, child, terminal, launch)
    except DecodeFailure as exc:
        print(json.dumps(exc.as_dict(), indent=2, sort_keys=True), file=sys.stderr)
        return 2
    output = child / f"{child.name}.packet.json"
    _write_json_exclusive(output, packet)
    print(json.dumps({"status": packet["status"], "packet": str(output), "packet_sha256": _sha256(output)}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
