#!/usr/bin/env python3
"""Offline native admission for the relocatable 0953 FP packet (v3 custody and raw-gate binding).

This reader is derived from the frozen v1 parser but tightens artifact custody:
child plan/runner/start/command descriptors, every retained child artifact,
the exact frame-pointer record argv, and exact decoder argv templates must all
match before semantic or stack parsing. It never starts a benchmark or perf.
Packet-local inputs resolve below ``--packet``; external immutable inputs are
explicit plan descriptors checked by digest, size, mode, and regular-file
status. A failed owner gate emits no native fraction or leaf categories.
"""
from __future__ import annotations

import argparse
import collections
import decimal
import hashlib
import importlib.util
import json
import os
import re
import stat
import struct
import sys
from pathlib import Path
from typing import Any, Iterable

DEFAULT_PACKET = Path(__file__).resolve().parents[1]
DEFAULT_PLAN_NAME = "plan-0953-v1.json"
SEMANTIC_MODULE_NAME = "semantic_0953_v1.py"
EXPECTED_SCHEMA = "litchi.performance.0953.native-publish-sampling-plan.v1"
CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
EVENT = "cycles:u"
PAGE = 4096
OWNER_MIN_SAMPLES = 100
OWNER_MIN_PERIOD = 100_000
DECIMAL = decimal.Decimal
HEX = re.compile(r"^(?:0x)?[0-9a-fA-F]+$")


SAMPLE_HEADER = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)/(?P<tid>\d+)\s+"
    r"(?P<time>[0-9]+(?:\.[0-9]+)?)\s*:\s*(?P<period>\d+)\s+(?P<event>\S+)"
)
FRAME = re.compile(
    r"^\s*(?P<ip>(?:0x)?[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+"
    r"\((?P<dso>[^)]*)\)(?:\s+\d+)?\s*$"
)
MMAP = re.compile(
    r"^\s*\S+\s+\d+/\d+\s+(?P<header_time>[0-9]+(?:\.[0-9]+)?)\s*:\s+"
    r"PERF_RECORD_MMAP2\s+(?P<pid>\d+)/(?P<tid>\d+)\s*:\s+"
    r"\[(?P<start>0x[0-9a-fA-F]+)\((?P<size>0x[0-9a-fA-F]+)\)\s+@\s+"
    r"(?P<offset>(?:0x)?[0-9a-fA-F]+)\s+(?P<meta>.*?)\]:\s+"
    r"(?P<perm>\S+)\s+(?P<dso>.*)$"
)
UNKNOWN = {"", "??", "[unknown]", "unknown"}


class AdmissionError(RuntimeError):
    pass


def require(value: bool, message: str) -> None:
    if not value:
        raise AdmissionError(message)


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    st = path.stat()
    return {"path": str(path), "bytes": st.st_size, "sha256": sha(path),
            "mode": stat.S_IMODE(st.st_mode), "executable": bool(st.st_mode & stat.S_IXUSR)}


def unique_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in out, f"duplicate JSON key: {key}")
        out[key] = value
    return out


def load_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=unique_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise AdmissionError(f"invalid JSON {path}: {exc}") from exc


def resolve_path(raw: Any, packet: Path, label: str, *, kind: str = "file") -> Path:
    require(isinstance(raw, str) and raw and "\x00" not in raw, f"{label} path is invalid")
    candidate = Path(raw)
    if not candidate.is_absolute():
        candidate = packet / candidate
        try:
            candidate.resolve().relative_to(packet.resolve())
        except ValueError as exc:
            raise AdmissionError(f"{label} relative path escapes packet: {raw}") from exc
    path = candidate.resolve()
    require(not path.is_symlink(), f"{label} is a symlink: {path}")
    if kind == "file":
        require(path.is_file(), f"{label} missing regular file: {path}")
    elif kind == "directory":
        require(path.is_dir(), f"{label} missing directory: {path}")
    else:
        raise AdmissionError(f"unknown path kind for {label}: {kind}")
    return path


def verify_plan_descriptor(value: Any, packet: Path, label: str) -> tuple[Path, dict[str, Any]]:
    require(isinstance(value, dict), f"{label} descriptor is missing")
    path = resolve_path(value.get("path"), packet, label)
    actual = descriptor(path)
    for key in ("bytes", "sha256", "mode", "executable"):
        require(actual[key] == value.get(key), f"{label} descriptor {key} differs")
    return path, actual


def import_semantic(path: Path) -> Any:
    spec = importlib.util.spec_from_file_location("semantic_0953_bound", path)
    require(spec is not None and spec.loader is not None, "semantic adapter cannot load")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_plan_for(packet: Path, plan_path: Path | None) -> tuple[dict[str, Any], Path]:
    packet = packet.resolve()
    candidate = plan_path if plan_path is not None else packet / DEFAULT_PLAN_NAME
    path = candidate if candidate.is_absolute() else packet / candidate
    path = path.resolve()
    try:
        path.relative_to(packet)
    except ValueError as exc:
        raise AdmissionError(f"plan is outside packet: {path}") from exc
    plan = load_json(path)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == EXPECTED_SCHEMA, "0953 plan schema differs")
    return plan, path


def verify_raw_gate(plan: dict[str, Any]) -> None:
    admission = plan.get("first_child_admission")
    require(isinstance(admission, dict), "first_child_admission is missing")
    raw_gate = admission.get("raw_gate")
    require(isinstance(raw_gate, dict), "first-child raw gate is missing")
    expected = {
        "owner_samples_min": OWNER_MIN_SAMPLES,
        "fully_resolved_owner_samples_min": OWNER_MIN_SAMPLES,
        "owner_period_min": OWNER_MIN_PERIOD,
        "lost_records": 0,
    }
    for key, value in expected.items():
        actual = raw_gate.get(key)
        require(type(actual) is int and actual == value,
                f"first-child raw gate {key} differs")


def verify_bindings(packet: Path, plan_path: Path | None) -> tuple[dict[str, Any], Any, dict[str, Any], dict[str, Any], Path]:
    plan, plan_file = load_plan_for(packet, plan_path)
    verify_raw_gate(plan)
    contract = plan.get("contract", plan)
    require(isinstance(contract, dict), "plan contract is missing")
    require(contract.get("samples") == 200 and contract.get("warmup") == 0,
            "0953 sample contract differs")
    require(contract.get("case") == CASE and contract.get("cpu") == 12,
            "0953 workload contract differs")
    require(contract.get("event", EVENT) == EVENT, "0953 event contract differs")
    schedule = plan.get("schedule")
    require(isinstance(schedule, list) and schedule, "0953 schedule is missing")
    shapes = contract.get("shapes") or sorted({row.get("shape") for row in schedule})
    require(isinstance(shapes, list) and shapes and all(isinstance(shape, str) for shape in shapes),
            "0953 shapes are missing")
    paths: dict[str, Path] = {}
    actuals: dict[str, dict[str, Any]] = {}
    for key in ("binary", "runner", "semantic_adapter"):
        paths[key], actuals[key] = verify_plan_descriptor(plan.get(key), packet, f"plan {key}")
    for key in ("sample_exporter", "mmap_exporter"):
        if key in plan:
            paths[key], actuals[key] = verify_plan_descriptor(plan.get(key), packet, f"plan {key}")
    require(paths["semantic_adapter"].name == SEMANTIC_MODULE_NAME,
            "plan semantic adapter filename differs")
    sem = import_semantic(paths["semantic_adapter"])
    sem_bindings, oracle, sem_contract, sem_plan_file = sem.verify_bindings(packet, plan_file)
    require(sem_plan_file.resolve() == plan_file.resolve(), "semantic plan path differs")
    require(sem_contract["samples"] == contract["samples"]
            and sem_contract["warmup"] == contract["warmup"]
            and sem_contract["case"] == contract["case"]
            and sem_contract["cpu"] == contract["cpu"], "semantic contract differs")
    require(sem_bindings["binary"]["sha256"] == actuals["binary"]["sha256"],
            "semantic and plan binary bindings differ")
    bindings = {key: actual for key, actual in actuals.items()}
    bindings["semantic"] = actuals["semantic_adapter"]
    bindings["semantic_chain"] = sem_bindings
    paths["plan"] = plan_file
    bindings["plan"] = descriptor(plan_file)
    return bindings, sem, plan, contract, plan_file


def terminal_pass(path: Path, label: str) -> dict[str, Any]:
    value = load_json(path)
    require(value.get("returncode") == 0 and value.get("timed_out") is False,
            f"{label} terminal did not pass")
    require(value.get("launch_error") is None and value.get("process_group_clear") is True
            and value.get("terminal_before_parse") is True, f"{label} terminal custody failed")
    return value


def equal_desc(left: Any, right: Any, label: str) -> None:
    require(isinstance(left, dict) and isinstance(right, dict), f"{label} descriptor missing")
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        require(left.get(key) == right.get(key), f"{label} descriptor {key} differs")


def equal_plan_desc(given: Any, expected: dict[str, Any], label: str) -> None:
    """Compare descriptor content and its resolved path to a bound descriptor."""
    require(isinstance(given, dict), f"{label} descriptor missing")
    for key in ("bytes", "sha256", "mode", "executable"):
        require(given.get(key) == expected.get(key), f"{label} descriptor {key} differs")
    require(isinstance(given.get("path"), str), f"{label} descriptor path is missing")
    require(Path(given["path"]).resolve() == Path(expected["path"]).resolve(),
            f"{label} descriptor path differs")


def expand_argv(template: Any, *, packet: Path, child: Path, output: Path,
                binary: Path, shape: str, label: str) -> list[str]:
    require(isinstance(template, list) and template, f"{label} argv template is missing")
    replacements = {
        "$PACKET": str(packet.resolve()), "$CHILD": str(child.resolve()),
        "$ANALYSIS_CHILD": str(output.resolve()), "$BINARY": str(binary.resolve()),
        "$SHAPE": shape,
    }
    result: list[str] = []
    for item in template:
        require(isinstance(item, str), f"{label} argv contains a non-string")
        value = item
        for marker, replacement in replacements.items():
            value = value.replace(marker, replacement)
        result.append(value)
    return result


def contract_for(plan: dict[str, Any], kind: str) -> dict[str, Any]:
    contracts = plan.get("decoder_contract")
    require(isinstance(contracts, dict), "decoder_contract is missing")
    aliases = (kind, "primary_stack_decoder" if kind == "sample" else "mapping_decoder")
    value = next((contracts.get(name) for name in aliases
                  if isinstance(contracts.get(name), dict)), None)
    require(isinstance(value, dict), f"{kind} decoder contract is missing")
    require(isinstance(value.get("argv_template"), list), f"{kind} argv template is missing")
    require(isinstance(value.get("debug_mode"), bool), f"{kind} debug-mode binding is missing")
    return value


def record_template(plan: dict[str, Any]) -> list[Any]:
    contract = plan.get("record_contract")
    require(isinstance(contract, dict), "record_contract is missing")
    template = contract.get("argv_template")
    require(isinstance(template, list), "record argv template is missing")
    return template


def child_context(plan: dict[str, Any], packet: Path, index: int,
                   bindings: dict[str, Any]) -> tuple[dict[str, Any], Path, dict[str, Any]]:
    rows = {row["index"]: row for row in plan.get("schedule", [])}
    require(index in rows, f"schedule index {index} missing")
    row = rows[index]
    child_name = plan.get("child_directory_template", "capture/{index:02d}-{shape}").format(
        index=index, shape=row["shape"])
    child = resolve_path(child_name, packet, f"child {index}", kind="directory")
    names = plan.get("child_artifacts", {})
    names = {
        "start": names.get("start", "start.json"),
        "command": names.get("command", "command.json"),
        "record_terminal": names.get("record_terminal", "record-terminal.json"),
        "record_stdout": names.get("record_stdout", "record.stdout"),
        "record_stderr": names.get("record_stderr", "record.stderr"),
        "report": names.get("report", "report.json"),
        "perf_data": names.get("perf_data", "perf.data"),
    }
    expected_names = set(names.values()) | {"receipt.json"}
    entries = list(child.iterdir())
    require(all(path.is_file() and not path.is_symlink() for path in entries),
            "child contains a non-regular artifact")
    actual_paths = {path.name: path for path in entries}
    require(set(actual_paths) == expected_names,
            f"child artifact set differs: expected {sorted(expected_names)}, got {sorted(actual_paths)}")
    receipt = load_json(child / "receipt.json")
    require(receipt.get("status") == "raw_captured" and receipt.get("retry") is False,
            "child raw receipt is not admitted")
    require(receipt.get("row") == row, "child schedule row differs")
    require(receipt.get("binary", {}).get("sha256") == plan["binary"]["sha256"],
            "child binary differs")
    equal_plan_desc(receipt.get("plan"), bindings["plan"], "child receipt plan")
    equal_plan_desc(receipt.get("runner"), bindings["runner"], "child receipt runner")
    equal_plan_desc(receipt.get("binary"), bindings["binary"], "child receipt binary")
    expected_argv = expand_argv(record_template(plan), packet=packet, child=child,
                                output=child, binary=Path(bindings["binary"]["path"]),
                                shape=row["shape"], label="record")
    require(expected_argv.count("--call-graph") == 1
            and expected_argv[expected_argv.index("--call-graph") + 1] == "fp",
            "record template is not exactly frame-pointer callgraph")
    require("--debug-mode" not in expected_argv
            and all("dwarf" not in item for item in expected_argv),
            "record template contains forbidden decoder/callgraph option")
    start = load_json(child / names["start"])
    command = load_json(child / names["command"])
    for label, value in (("child start", start), ("child command", command)):
        require(value.get("row") == row, f"{label} row differs")
        equal_plan_desc(value.get("plan"), bindings["plan"], f"{label} plan")
        equal_plan_desc(value.get("runner"), bindings["runner"], f"{label} runner")
        equal_plan_desc(value.get("binary"), bindings["binary"], f"{label} binary")
        require(value.get("record_argv") == expected_argv, f"{label} record argv differs")
        require(value.get("retry") is False, f"{label} retry flag differs")
    require(start.get("record_argv") == command.get("record_argv"),
            "child start/command record argv differ")
    terminal_record = terminal_pass(child / names["record_terminal"], "record")
    require(terminal_record.get("argv") == expected_argv,
            "record terminal argv differs from frozen FP template")
    # If the plan freezes a capture environment/cwd, bind those values too;
    # otherwise at least require start and command to agree exactly.
    env = plan.get("capture_environment", plan.get("environment"))
    require(isinstance(env, dict), "capture_environment is missing from the frozen plan")
    require(start.get("environment") == env and command.get("environment") == env,
            "child environment differs from frozen capture environment")
    expected_cwd = plan.get("capture_cwd", plan.get("cwd"))
    require(isinstance(expected_cwd, str) and expected_cwd,
            "capture_cwd is missing from the frozen plan")
    require(start.get("cwd") == expected_cwd and command.get("cwd") == expected_cwd,
            "child cwd differs from frozen capture cwd")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, list) and len(artifacts) == len(expected_names) - 1,
            "child artifact descriptor list is incomplete")
    seen: set[str] = set()
    for item in artifacts:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                "child artifact descriptor is malformed")
        artifact_path = Path(item["path"]).resolve()
        require(artifact_path.parent == child.resolve(), "child artifact escapes child directory")
        name = artifact_path.name
        require(name in expected_names - {"receipt.json"} and name not in seen,
                f"unexpected or duplicate child artifact: {name}")
        equal_desc(item, descriptor(artifact_path), f"child artifact {name}")
        seen.add(name)
    require(seen == expected_names - {"receipt.json"}, "child artifact descriptor coverage differs")
    for key, name in (("record_terminal", names["record_terminal"]),
                      ("report", names["report"]), ("perf_data", names["perf_data"])):
        equal_desc(receipt.get(key), descriptor(child / name), f"child receipt {key}")
    return row, child, {"receipt": receipt, "record_terminal": terminal_record,
                        "start": start, "command": command, "names": names,
                        "expected_record_argv": expected_argv}


def export_context(root: Path, packet: Path, child: Path, receipt: dict[str, Any],
                   row: dict[str, Any], kind: str, plan: dict[str, Any],
                   bindings: dict[str, Any], child_custody: dict[str, Any]) -> dict[str, Any]:
    value = load_json(root / "receipt.json")
    expected_schema = plan.get("export_receipt_schemas", {}).get(
        kind, f"litchi.performance.0953.{kind}-export-receipt.v1")
    require(value.get("schema") == expected_schema, f"{kind} export schema differs")
    require(value.get("status") == "pass" and value.get("terminal_before_parse") is True
            and value.get("workload_launched") is False, f"{kind} export did not pass")
    require(value.get("row") == row, f"{kind} export schedule row differs")
    try:
        root.resolve().relative_to((packet / "capture").resolve())
    except ValueError:
        pass
    else:
        raise AdmissionError(f"{kind} export is inside sealed capture root")
    child_perf = descriptor(child / child_custody["names"]["perf_data"])
    equal_desc(value.get("perf_data"), child_perf, f"{kind} perf.data")
    for key, name in (("start", "start.json"), ("terminal", "terminal.json"),
                      ("events", "events.txt"), ("stderr", "stderr.txt")):
        actual = descriptor(root / name)
        equal_desc(value.get(key), actual, f"{kind} {key}")
    terminal = terminal_pass(root / "terminal.json", f"{kind} export")
    start = load_json(root / "start.json")
    argv = start.get("argv")
    require(isinstance(argv, list), f"{kind} export argv missing")
    expected = expand_argv(contract_for(plan, kind)["argv_template"], packet=packet,
                           child=child, output=root, binary=Path(bindings["binary"]["path"]),
                           shape=row["shape"], label=kind)
    require(argv == expected, f"{kind} start argv differs from frozen decoder template")
    require(terminal.get("argv") == expected, f"{kind} terminal argv differs")
    equal_plan_desc(start.get("plan"), bindings["plan"], f"{kind} start plan")
    equal_plan_desc(start.get("runner"), bindings["runner"], f"{kind} start runner")
    equal_plan_desc(start.get("child_start"), descriptor(child / child_custody["names"]["start"]),
                    f"{kind} child start")
    equal_desc(start.get("perf_data"), child_perf, f"{kind} start perf.data")
    equal_plan_desc(start.get("binary"), bindings["binary"], f"{kind} start binary")
    require(start.get("environment") == plan.get("capture_environment"),
            f"{kind} export environment differs from frozen capture environment")
    require(start.get("cwd") == plan.get("capture_cwd", plan.get("cwd")),
            f"{kind} export cwd differs from frozen capture cwd")
    require(start.get("workload_launched") is False, f"{kind} export workload flag differs")
    require(("--debug-mode" in expected) == bool(contract_for(plan, kind)["debug_mode"]),
            f"{kind} debug-mode does not match frozen decoder contract")
    require("--show-lost-events" in expected and "--ns" in expected and "--no-inline" in expected,
            f"{kind} event-field contract differs")
    if kind == "mmap":
        for flag in ("--show-mmap-events", "--show-task-events", "--hide-call-graph"):
            require(flag in expected, f"mmap export missing {flag}")
    else:
        require("--show-mmap-events" not in expected and "--show-task-events" not in expected,
                "sample export has mapping records")
    events = root / "events.txt"
    require(events.stat().st_size > 0, f"{kind} event export is empty")
    return {key: descriptor(root / name) for key, name in (("receipt", "receipt.json"),
            ("start", "start.json"), ("terminal", "terminal.json"),
            ("events", "events.txt"), ("stderr", "stderr.txt"))} | {
                "terminal_value": terminal, "argv": argv}


def read_elf(binary: Path) -> tuple[list[dict[str, int]], str]:
    with binary.open("rb") as stream:
        ident = stream.read(64)
        require(len(ident) == 64 and ident[:4] == b"\x7fELF" and ident[4] == 2 and ident[5] == 1,
                "bound binary is not ELF64 little-endian")
        fields = struct.unpack_from("<HHIQQQIHHHHHH", ident, 16)
        _, _, _, _, phoff, _, _, _, phentsize, phnum, _, _, _ = fields
        require(phentsize >= 56 and phnum > 0, "ELF program header table is invalid")
        stream.seek(phoff)
        headers = [stream.read(phentsize) for _ in range(phnum)]
        loads: list[dict[str, int]] = []
        notes: list[tuple[int, int]] = []
        for raw in headers:
            require(len(raw) >= 56, "truncated ELF program header")
            p_type, p_flags, p_offset, p_vaddr, _, p_filesz, p_memsz, p_align = struct.unpack_from("<IIQQQQQQ", raw, 0)
            if p_type == 1:
                require(p_filesz <= p_memsz and p_offset % PAGE == p_vaddr % PAGE,
                        "PT_LOAD alignment/size is invalid")
                loads.append({"flags": p_flags, "offset": p_offset, "vaddr": p_vaddr,
                              "filesz": p_filesz, "memsz": p_memsz, "align": p_align})
            elif p_type == 4 and p_filesz:
                notes.append((p_offset, p_filesz))
        require(loads, "binary has no PT_LOAD segments")
        build_id = None
        for offset, size in notes:
            stream.seek(offset)
            data = stream.read(size)
            pos = 0
            while pos + 12 <= len(data):
                namesz, descsz, ntype = struct.unpack_from("<III", data, pos)
                pos += 12
                nend = pos + ((namesz + 3) // 4) * 4
                dend = nend + ((descsz + 3) // 4) * 4
                require(dend <= len(data), "truncated ELF note")
                name = data[pos:pos + namesz].rstrip(b"\0")
                desc = data[nend:nend + descsz]
                if name == b"GNU" and ntype == 3:
                    build_id = desc.hex()
                pos = dend
        require(build_id is not None, "GNU build-id is missing")
    return loads, build_id


def align_down(value: int) -> int:
    return value & ~(PAGE - 1)


def align_up(value: int) -> int:
    return (value + PAGE - 1) & ~(PAGE - 1)


def load_bias_for(offset: int, start: int, size: int, perm: str,
                  loads: list[dict[str, int]]) -> tuple[int | None, int | None, bool]:
    page_offset = align_down(offset)
    candidates = []
    for index, seg in enumerate(loads):
        lo = align_down(seg["offset"])
        hi = align_up(seg["offset"] + max(seg["filesz"], seg["memsz"]))
        if lo <= page_offset < hi and page_offset + size > lo:
            mapped_vaddr = seg["vaddr"] + (page_offset - seg["offset"])
            candidates.append((index, start - mapped_vaddr, seg))
    require(candidates, f"mapping file offset {offset:#x} does not bind to PT_LOAD")
    # ELF LOAD segments can share a page at a permission boundary, while the
    # loader may also expose a broad read-only file mapping.  Use mmap
    # permissions to select the compatible segment.  Never fall back to an
    # incompatible candidate.  Executable mappings must bind uniquely because
    # they can participate in static-IP ownership; ambiguous non-executable
    # rows remain present but cannot resolve a code frame.
    required_flags = ((4 if "r" in perm else 0)
                      | (2 if "w" in perm else 0)
                      | (1 if "x" in perm else 0))
    filtered = [item for item in candidates
                if required_flags & ~item[2]["flags"] == 0]
    executable = "x" in perm
    if not filtered:
        # Keep a non-executable loader overlay visible, but it cannot resolve
        # code and therefore must never be selected for static ownership.
        if not executable:
            return None, None, True
        raise AdmissionError(f"mapping permissions do not bind PT_LOAD at {offset:#x}")
    if executable:
        require(len(filtered) == 1,
                f"executable mapping file offset {offset:#x} binds ambiguously to PT_LOAD")
    if len(filtered) != 1:
        return None, None, True
    index, bias, _ = filtered[0]
    return index, bias, False


def parse_mmap_events(path: Path, binary: Path, loads: list[dict[str, int]], build_id: str) -> dict[str, Any]:
    target = str(binary.resolve())
    rows: list[dict[str, Any]] = []
    unsupported = 0
    for number, line in enumerate(path.read_text(encoding="utf-8", errors="strict").splitlines(), 1):
        if not line.strip():
            continue
        match = MMAP.match(line)
        if match is None:
            if "PERF_RECORD_LOST" in line or "PERF_RECORD_FORK" in line or "PERF_RECORD_EXIT" in line or "PERF_RECORD_COMM" in line:
                continue
            unsupported += 1
            continue
        meta = match.group("meta")
        ids = re.findall(r"<([0-9a-fA-F]+)>", meta)
        dso = match.group("dso").strip()
        row = {"line": number, "time": DECIMAL(match.group("header_time")),
               "pid": int(match.group("pid")), "tid": int(match.group("tid")),
               "start": int(match.group("start"), 16), "size": int(match.group("size"), 16),
               "offset": int(match.group("offset"), 16), "perm": match.group("perm"),
               "dso": dso, "build_id": ids[0] if ids else None}
        require(row["size"] > 0 and row["start"] + row["size"] > row["start"], "invalid MMAP2 range")
        row["is_bound_dso"] = dso == target or Path(dso).name == binary.name
        if row["is_bound_dso"]:
            require(row["build_id"] == build_id, f"bound DSO build-id differs at mmap line {number}")
            row["pt_load_index"], row["load_bias"], row["mapping_ambiguous"] = load_bias_for(row["offset"], row["start"], row["size"], row["perm"], loads)
        rows.append(row)
    require(unsupported == 0, f"unsupported mmap records: {unsupported}")
    require(rows, "mmap export contains no MMAP2 rows")
    bound = [row for row in rows if row["is_bound_dso"]]
    require(bound and any("x" in row["perm"] for row in bound), "no executable bound DSO mapping")
    return {"rows": rows, "bound_rows": bound, "line_count": len(rows), "unsupported": unsupported,
            "target_path": target, "build_id": build_id,
            "bound_pids": sorted({row["pid"] for row in bound}),
            "bound_build_ids": sorted({row["build_id"] for row in bound}),
            "ambiguous_bound_mappings": sum(1 for row in bound if row.get("mapping_ambiguous"))}


def map_for_ip(maps: list[dict[str, Any]], pid: int, when: decimal.Decimal, ip: int, dso: str) -> dict[str, Any] | None:
    candidates = [row for row in maps if row["pid"] == pid and row["time"] <= when
                  and row["is_bound_dso"] and row["start"] <= ip < row["start"] + row["size"]
                  and (dso == row["dso"] or Path(dso).name == Path(row["dso"]).name)]
    return max(candidates, key=lambda row: (row["time"], row["line"])) if candidates else None


def parse_frame(line: str) -> dict[str, Any] | None:
    match = FRAME.match(line)
    if match is None:
        stripped = line.strip()
        tokens = stripped.split()
        if not tokens or not HEX.fullmatch(tokens[0]):
            return None
        return {"ip": int(tokens[0], 16), "symbol": tokens[1] if len(tokens) > 1 else "[unknown]",
                "dso": "[unknown]", "unresolved": True}
    symbol = match.group("symbol").strip()
    dso = match.group("dso").strip() or "[unknown]"
    return {"ip": int(match.group("ip"), 16), "symbol": symbol, "dso": dso,
            "unresolved": symbol in UNKNOWN or dso.lower() in UNKNOWN}


def static_ip(frame: dict[str, Any], pid: int, when: decimal.Decimal, maps: list[dict[str, Any]], binary: Path) -> tuple[int | None, dict[str, Any] | None]:
    dso = frame["dso"]
    if dso not in {str(binary), binary.name} and Path(dso).name != binary.name:
        return None, None
    mapping = map_for_ip(maps, pid, when, frame["ip"], dso)
    if (mapping is None or mapping.get("mapping_ambiguous")
            or "x" not in mapping.get("perm", "")
            or mapping.get("load_bias") is None):
        return None, None
    return frame["ip"] - mapping["load_bias"], mapping


def parse_samples(path: Path, binary: Path, mmap: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    symbols = plan["static_symbols"]
    owners = [(int(item["address"], 16), int(item["end_exclusive"], 16)) for item in symbols["owner_ranges"]]
    parent = (int(symbols["parent_range"]["address"], 16), int(symbols["parent_range"]["end_exclusive"], 16))
    maps = mmap["rows"]
    all_samples = all_period = owner_samples = owner_period = 0
    lost_records = 0
    unresolved_period = 0
    unknown_period = 0
    unknown_samples = 0
    reasons: collections.Counter[str] = collections.Counter()
    frame_lengths: collections.Counter[int] = collections.Counter()
    leaves: collections.Counter[str] = collections.Counter()
    nested: collections.Counter[str] = collections.Counter()
    current: dict[str, Any] | None = None

    def finish(record: dict[str, Any]) -> None:
        nonlocal owner_samples, owner_period, unknown_samples, unknown_period, unresolved_period
        frames = record["frames"]
        frame_lengths[len(frames)] += 1
        resolved: list[tuple[int | None, dict[str, Any] | None]] = []
        for frame in frames:
            static, mapping = static_ip(frame, record["pid"], record["time"], maps, binary)
            resolved.append((static, mapping))
        owner_positions = [i for i, (static, mapping) in enumerate(resolved)
                           if static is not None and any(lo <= static < hi for lo, hi in owners)]
        reason = "owner_not_observed"
        qualified = False
        if any(frame.get("unresolved") for frame in frames) or record.get("truncated"):
            unresolved_period += record["period"]
            reason = "unresolved_or_truncated"
        if owner_positions:
            oi = owner_positions[0]
            if oi + 1 >= len(frames):
                reason = "owner_without_parent"
            else:
                parent_static, parent_mapping = resolved[oi + 1]
                if parent_static is None or not (parent[0] <= parent_static < parent[1]):
                    reason = "wrong_or_missing_parent"
                elif not frames[oi + 1].get("unresolved"):
                    qualified = True
                    reason = "owner_parent_bound"
        reasons[reason] += 1
        if qualified and not any(frame.get("unresolved") for frame in frames) and not record.get("truncated"):
            owner_samples += 1
            owner_period += record["period"]
            leaves[frames[0]["symbol"]] += record["period"]
            for frame in frames:
                symbol = frame["symbol"]
                if any(token in symbol for token in ("zlib", "deflate", "xml_minifier", "quick_xml", "soapberry_zip", "litchi_opc")):
                    nested[symbol] += record["period"]
        else:
            unknown_samples += 1
            unknown_period += record["period"]

    with path.open("r", encoding="utf-8", errors="strict") as stream:
        for line_number, line in enumerate(stream, 1):
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "PERF_RECORD_LOST" in line:
                lost_records += 1
                continue
            if stripped.startswith("..."):
                require(current is not None, f"sample truncation marker before sample at line {line_number}")
                current["truncated"] = True
                continue
            header = SAMPLE_HEADER.match(line)
            if header:
                if current is not None:
                    finish(current)
                event = header.group("event").rstrip(":")
                require(event == EVENT, f"unexpected perf event {event} at line {line_number}")
                period = int(header.group("period"))
                require(period > 0, f"nonpositive sample period at line {line_number}")
                all_samples += 1
                all_period += period
                current = {"pid": int(header.group("pid")), "tid": int(header.group("tid")),
                           "time": DECIMAL(header.group("time")), "period": period,
                           "frames": [], "truncated": False}
                continue
            require(current is not None, f"unparsed line outside sample at {line_number}: {stripped[:120]}")
            frame = parse_frame(line)
            require(frame is not None, f"unparsed frame at line {line_number}: {stripped[:120]}")
            current["frames"].append(frame)
    if current is not None:
        finish(current)
    require(all_samples > 0, "corrected sample export contains no samples")
    require(all_period == owner_period + unknown_period, "sample period partition is not disjoint")
    gate = {"lost_records_zero": lost_records == 0,
            "fully_resolved_owner_samples_min": owner_samples >= OWNER_MIN_SAMPLES,
            "owner_period_min": owner_period >= OWNER_MIN_PERIOD}
    admitted = all(gate.values())
    return {"all": {"samples": all_samples, "period": all_period},
            "owner_qualified": {"samples": owner_samples, "period": owner_period},
            "unknown_or_unqualified": {"samples": unknown_samples, "period": unknown_period},
            "unresolved_period": unresolved_period, "lost_records": lost_records,
            "frame_length_counts": dict(sorted(frame_lengths.items())),
            "classification_reasons": dict(sorted(reasons.items())),
            "exclusive_leaf_ranking": ([{"symbol": name, "period": value}
                                        for name, value in leaves.most_common()]
                                       if admitted else []),
            "nested_compression_xml_rows": ([{"symbol": name, "period": value}
                                              for name, value in nested.most_common()]
                                             if admitted else []),
            "nested_rows_overlap": True, "gate": gate,
            "admitted": admitted}


def write_exclusive(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8", newline="") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())


def run(args: argparse.Namespace) -> tuple[dict[str, Any], int]:
    packet = args.packet.resolve()
    bindings, semantic, plan, contract, plan_file = verify_bindings(packet, args.plan)
    row, child, custody = child_context(plan, packet, args.index, bindings)
    sample_root = resolve_path(str(args.sample_export), packet, "sample export", kind="directory")
    mmap_root = resolve_path(str(args.mmap_export), packet, "mmap export", kind="directory")
    sample = export_context(sample_root, packet, child, custody["receipt"], row, "sample", plan, bindings, custody)
    mmap_export = export_context(mmap_root, packet, child, custody["receipt"], row, "mmap", plan, bindings, custody)
    binary = Path(bindings["binary"]["path"])
    loads, build_id = read_elf(binary)
    mmap = parse_mmap_events(mmap_root / "events.txt", binary, loads, build_id)
    sem_result = semantic.validate_report(child / "report.json", row["shape"], packet, plan,
                                          contract, semantic._load_oracle(
                                              Path(bindings["semantic_chain"]["semantic_oracle"]["path"])),
                                          receipt_path=child / "receipt.json")
    profile = parse_samples(sample_root / "events.txt", binary, mmap, plan)
    status = "pass" if profile["admitted"] else "fail"
    result = {"schema": "litchi.performance.0953.native-admission.v3", "status": status,
              "plan": bindings["plan"], "runner": bindings["runner"],
              "binary": bindings["binary"], "semantic_adapter": bindings["semantic"],
              "semantic_chain": bindings["semantic_chain"],
              "sample_exporter": bindings.get("sample_exporter"),
              "mmap_exporter": bindings.get("mmap_exporter"), "row": row,
              "child": descriptor(child / "receipt.json"), "custody": custody,
              "sample_export": sample, "mmap_export": mmap_export,
              "elf": {"build_id": build_id, "pt_loads": loads},
              "mmap_binding": {key: value for key, value in mmap.items()
                               if key not in {"rows", "bound_rows"}},
              "semantic": sem_result, "profile": profile,
              "scope": "Descriptive cycles:u owner attribution only; no native timing, speedup, allocation, RSS, or adoption claim."}
    return result, 0 if status == "pass" else 2


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=DEFAULT_PACKET)
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--index", type=int, required=True)
    parser.add_argument("--sample-export", type=Path, required=True)
    parser.add_argument("--mmap-export", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        result, code = run(args)
        output = args.output if args.output.is_absolute() else args.packet / args.output
        write_exclusive(output.resolve(), result)
        print(json.dumps({"status": result["status"], "output": str(output.resolve()),
                          "sha256": sha(output.resolve())}, sort_keys=True))
        return code
    except (AdmissionError, AssertionError, OSError, ValueError, KeyError, ImportError, struct.error) as exc:
        failure = {"schema": "litchi.performance.0953.native-admission.v3",
                   "status": "error", "error": str(exc),
                   "scope": "No admission; failed closed before native attribution."}
        try:
            output = args.output if args.output.is_absolute() else args.packet / args.output
            write_exclusive(output.resolve(), failure)
        except Exception:
            pass
        print(f"analyze_0953_v2: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
