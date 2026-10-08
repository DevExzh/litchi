#!/usr/bin/env python3
"""Fail-closed 0973 protocol and dynamic ELF/MMAP2 layout helpers."""
from __future__ import annotations

import hashlib
import importlib.util
import json
import math
import os
import re
import struct
import subprocess
from pathlib import Path
from typing import Any

BASE_PROTOCOL_PATH = Path("/home/zhuhe/litchi-performance-0971/tools/protocol_common_0971_v2.py")
BASE_PROTOCOL_SHA = "4ff9af4c64b50896f39a6a30286d445ebe0c9364435888593ade45a11c5cab43"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


if _sha256(BASE_PROTOCOL_PATH) != BASE_PROTOCOL_SHA:
    raise RuntimeError("retained 0971 parser changed")
_spec = importlib.util.spec_from_file_location("litchi_protocol_common_0971_v2", BASE_PROTOCOL_PATH)
if _spec is None or _spec.loader is None:
    raise RuntimeError("cannot load retained 0971 parser")
_base = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_base)

ProtocolError = _base.ProtocolError
load_json = _base.load_json
write_json_exclusive = _base.write_json_exclusive
write_text_exclusive = _base.write_text_exclusive
sha256 = _base.sha256
file_binding = _base.file_binding
check_binding = _base.check_binding
parse_perf_script = _base.parse_perf_script
percentile = _base.percentile
pearson = _base.pearson
spearman = _base.spearman
finite_positive = _base.finite_positive
check_launch_metadata = _base.check_launch_metadata
check_artifact_bindings = _base.check_artifact_bindings
check_buildid_export = _base.check_buildid_export

SCHEMA = "litchi.performance.0973.instruction_profile.v2"
PLAN_SCHEMA = "litchi.performance.0973.plan.v2"
GO_SCHEMA = "litchi.performance.0973.root_go.v2"
PACKET_SCHEMA = "litchi.performance.0973.child_packet.v2"
MANIFEST_SCHEMA = "litchi.performance.0973.capture_manifest.v2"
LAYOUT_SCHEMA = "litchi.performance.0973.target_layout.v2"
HELPER_SCHEMA = "litchi.performance.0973.helper_seal.v2"
RAW_STACK_SCHEMA = "litchi.performance.0973.raw_stack.v2"
RAW_ARTIFACTS = [
    "report", "stderr", "perf_data", "launch", "command", "terminal", "layout", "mmap_dump",
    "raw_stack",
    "stock_script", "stock_stderr", "stock_no_log", "noop_script", "noop_stderr", "noop_log",
    "corrected_script", "corrected_stderr", "corrected_log", "inline", "header", "buildid",
]
MARKER = _base.MARKER
BINARY_PATH = _base.BINARY_PATH
BINARY_SHA = _base.BINARY_SHA
BINARY_BYTES = _base.BINARY_BYTES
BINARY_BUILD_ID = _base.BINARY_BUILD_ID
FIXTURE_PATH = _base.FIXTURE_PATH
FIXTURE_SHA = _base.FIXTURE_SHA
FIXTURE_BYTES = _base.FIXTURE_BYTES
CPU = 31
WARMUPS = 5
SAMPLES = 80
CALLS = WARMUPS + SAMPLES
PERIOD = 1_000_000
CALLGRAPH = "dwarf,32768"
STACK_CAP = 32_768
CHILDREN = 12
PERF_PATH = "/usr/bin/perf"
PERF_SHA = _base.PERF_SHA
PYTHON_PATH = "/usr/bin/python3.14"
TASKSET_PATH = "/usr/bin/taskset"
LAUNCHER_PATH = "/home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py"
LAUNCHER_SHA = _base.LAUNCH_SHA
VALIDATOR_PATH = _base.VALIDATOR_PATH
VALIDATOR_SHA = _base.VALIDATOR_SHA
SOURCE_MANIFEST_PATH = "/home/zhuhe/litchi-performance-0969/manifests/source-manifest-v1.json"
SOURCE_MANIFEST_SHA = _base.SOURCE_MANIFEST_SHA
SOURCE_FILE_MANIFEST_PATH = "/home/zhuhe/litchi-performance-0969/manifests/source-file-manifest-v1.json"
SOURCE_FILE_MANIFEST_SHA = _base.SOURCE_FILES_SHA
SOURCE_PROOF_PATH = "/home/zhuhe/litchi-performance-0971/receipts/source-content-proof-0971-v2.json"
SOURCE_PROOF_SHA = _base.SOURCE_PROOF_SHA
ENVIRONMENT = {"LC_ALL": "C", "LANG": "C", "TZ": "UTC"}
PAGE_SIZE = os.sysconf("SC_PAGE_SIZE")
LIBDW_PATH = "/usr/lib/x86_64-linux-gnu/libdw-0.194.so"
LIBDW_VERSION = "elfutils 0.194 (libdw1t64:amd64)"
LIBDW_SHA = "86a59c8071ab57c25daacb275cedae072de83151bd358d4ffdc7844afbd31109"
LIBDW_BYTES = 823824
LIBELF_PATH = "/usr/lib/x86_64-linux-gnu/libelf-0.194.so"
LIBELF_VERSION = "elfutils 0.194 (libelf1t64:amd64)"
LIBELF_SHA = "d1e1539cc36403b603deb0d4558711e50bc8144eec1d833ece69df81ce01b527"
LIBELF_BYTES = 154408
COMPILER_PATH = "/usr/bin/cc"
COMPILER_SHA = "864711037b900acdb7a449cf48d912cefa935b527fb60c767ccb956779e66136"
HEADER_PATH = "/home/zhuhe/litchi-performance-0972/header-v2/root/usr/include/elfutils/libdwfl.h"
HEADER_SHA = "6f9dc2c2f9ca8d691f42f78e7634c2e40750a3af134e8226bd7c5aff395a3657"
DESIGN_PATH = "/home/zhuhe/litchi-performance-0973/design-v1.md"
DESIGN_SHA = "0eba90ae698d18e48f52453791f200456923ae81f2445498b05fc4b3705f060a"
METRIC_PROPOSAL_PATH = "/home/zhuhe/litchi-performance-0973/analysis/metric-proposal-v1.md"
METRIC_PROPOSAL_SHA = "a09f6535426e5f199c5652d1eec8e4035833f5f4ac0d1b507e2312b70a1c3bcd"
DESIGN_PRECEDENCE_PATH = "/home/zhuhe/litchi-performance-0973/design-precedence-clarification-v1.md"
DESIGN_PRECEDENCE_SHA = "4075b51bc31596a5d5af34a23ffb32faa43941eacef4de9fb034e7904f613a35"
REVIEW_FINDINGS_PATH = "/home/zhuhe/litchi-performance-0973/receipts/independent-review-0973-v1.json"
REVIEW_FINDINGS_SHA = "2a31dea0857cdcfd6eb77ce4f1352d61fae09c8fde6cdb75daacd53b1fb516d1"
ASLR_CONTROL_PATH = "/home/zhuhe/litchi-performance-0973/receipts/aslr-translation-0973-final-v4.json"
SYNTHETIC_E2E_PATH = "/home/zhuhe/litchi-performance-0973/receipts/synthetic-packet-pipeline-author-v2-v4.json"
SHIM_BUILD_COMMAND = [
    "cc", "-shared", "-fPIC", "-O2", "-std=c11", "-Wall", "-Wextra", "-Werror",
    "-I/home/zhuhe/litchi-performance-0972/header-v2/root/usr/include",
    "-Wl,--version-script=/home/zhuhe/litchi-performance-0973/tools-v2/dwfl_dynamic_shim_v2.map",
    "-o", "/home/zhuhe/litchi-performance-0973/replay-dwfl-v2/dwfl_dynamic_shim_v2.so",
    "/home/zhuhe/litchi-performance-0973/tools-v2/dwfl_dynamic_shim_v2.c", "-ldl",
]

_MMAP2 = re.compile(
    r"PERF_RECORD_MMAP2\s+(?P<pid>\d+)/\d+:\s+"
    r"\[(?P<start>0x[0-9a-fA-F]+)\((?P<length>0x[0-9a-fA-F]+)\)\s+@\s+"
    r"(?P<offset>(?:0x[0-9a-fA-F]+|\d+))\s+(?P<meta>[^\]]*)\]:\s+"
    r"(?P<perm>\S+)\s+(?P<path>.*)$"
)
_SAMPLE_PID_TID = re.compile(
    r"^\s*(?:(?P<comm>\S+)\s+)?(?P<pid>\d+)"
    r"(?:/(?P<tid>\d+))?\s+\[(?P<cpu>\d+)\]\s+"
)


def check_artifact_bindings_0973(packet: dict[str, Any], packet_path: Path | None = None) -> dict[str, Path]:
    artifacts = packet.get("artifacts")
    if not isinstance(artifacts, dict) or set(artifacts) != set(RAW_ARTIFACTS):
        raise ProtocolError("0973 packet artifact set drift")
    result: dict[str, Path] = {}
    for name in RAW_ARTIFACTS:
        result[name] = check_binding(artifacts[name], f"0973 packet artifact {name}")
    if packet_path is not None:
        child_dir = packet_path.parent
        for name, path in result.items():
            if path.parent != child_dir:
                raise ProtocolError(f"0973 packet artifact escaped child directory: {name}")
    return result


def check_raw_stack_summary(summary: dict[str, Any], *, plan: dict[str, Any]) -> None:
    """Validate the bounded raw STACK_USER admission receipt.

    A synthetic packet may use the explicitly declared offline perf-boundary
    mock, but that path is never accepted for a real capture plan.
    """
    if summary.get("schema") != RAW_STACK_SCHEMA or summary.get("status") != "complete":
        raise ProtocolError("raw STACK_USER receipt schema/status mismatch")
    if summary.get("period") != PERIOD or summary.get("stack_cap") != STACK_CAP:
        raise ProtocolError("raw STACK_USER period/cap contract drift")
    synthetic = summary.get("offline_boundary_mock") is True
    declared = (plan.get("synthetic_boundary") or {}).get("enabled") is True
    if synthetic and not declared:
        raise ProtocolError("offline raw boundary mock is not declared by the plan")
    if not synthetic and declared:
        raise ProtocolError("synthetic boundary declaration cannot admit a real capture")
    if synthetic:
        if summary.get("boundary_scope") != "offline perf-boundary mock; no native sample bytes":
            raise ProtocolError("raw boundary mock scope is not explicit")
    else:
        if summary.get("sample_line_count") != summary.get("sample_count"):
            raise ProtocolError("raw sample line/record count mismatch")
        if summary.get("missing") not in ([], {}):
            raise ProtocolError("raw STACK_USER fields are missing")
        if summary.get("violations") != [] or summary.get("bounds_valid") is not True:
            raise ProtocolError("raw STACK_USER bounds or record validation failed")
        if summary.get("pid_tid_cpu_valid") is not True or summary.get("period_valid") is not True:
            raise ProtocolError("raw sample PID/TID/CPU/period validation failed")
        pf_x = summary.get("pf_x") or {}
        if not isinstance(pf_x.get("interval_count"), int) or pf_x.get("interval_count") < 1 or not isinstance(pf_x.get("sample_ip_in_interval_count"), int):
            raise ProtocolError("raw PF_X interval binding missing")
    if not isinstance(summary.get("stack_size_counts"), dict) or not isinstance(summary.get("dyn_size_counts"), dict):
        raise ProtocolError("raw STACK_USER distributions missing")


def _page_down(value: int) -> int:
    return value - (value % PAGE_SIZE)


def _hex(value: int) -> str:
    return f"0x{value:x}"


def _parse_elf_loads(binary: Path) -> dict[str, Any]:
    raw = binary.read_bytes()
    if len(raw) < 64 or raw[:4] != b"\x7fELF" or raw[4] != 2 or raw[5] != 1:
        raise ProtocolError("target must be a little-endian ELF64 binary")
    header = struct.unpack_from("<16sHHIQQQIHHHHHH", raw, 0)
    _, e_type, e_machine, _, _, e_phoff, _, _, _, e_phentsize, e_phnum, _, _, _ = header
    if e_type not in (2, 3) or e_machine != 62 or e_phentsize != 56:
        raise ProtocolError("target ELF header type/machine/program-header size mismatch")
    loads: list[dict[str, Any]] = []
    for index in range(e_phnum):
        offset = e_phoff + index * e_phentsize
        if offset + e_phentsize > len(raw):
            raise ProtocolError("target ELF program header exceeds file")
        p_type, p_flags, p_offset, p_vaddr, _, p_filesz, p_memsz, p_align = struct.unpack_from("<IIQQQQQQ", raw, offset)
        if p_type != 1:
            continue
        if p_filesz > p_memsz or p_align == 0 or p_align & (p_align - 1):
            raise ProtocolError("invalid PT_LOAD alignment/size")
        loads.append({
            "index": index,
            "flags": p_flags,
            "offset": p_offset,
            "offset_page": _page_down(p_offset),
            "vaddr": p_vaddr,
            "vaddr_page": _page_down(p_vaddr),
            "filesz": p_filesz,
            "memsz": p_memsz,
            "align": p_align,
        })
    if not loads:
        raise ProtocolError("target has no PT_LOAD segments")
    lowest = min(row["vaddr_page"] for row in loads)
    return {
        "class": "ELF64",
        "data": "little",
        "type": e_type,
        "machine": e_machine,
        "page_size": PAGE_SIZE,
        "lowest_load_vaddr_page": lowest,
        "pt_loads": loads,
    }


def _parse_mmap_dump(text: str, *, binary_path: str, build_id: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for line_number, line in enumerate(text.splitlines(), 1):
        match = _MMAP2.search(line)
        if match is None:
            continue
        path = match.group("path")
        meta = match.group("meta")
        build_match = re.search(r"<([0-9a-fA-F]+)>", meta)
        row = {
            "line": line_number,
            "pid": int(match.group("pid")),
            "start": int(match.group("start"), 16),
            "length": int(match.group("length"), 16),
            "offset": int(match.group("offset"), 0),
            "meta": meta,
            "build_id": build_match.group(1).lower() if build_match else None,
            "permissions": match.group("perm"),
            "path": path,
        }
        if path == binary_path:
            if row["build_id"] != build_id.lower():
                raise ProtocolError("target MMAP2 path has wrong or missing BuildID")
            rows.append(row)
    if not rows:
        raise ProtocolError("perf MMAP2 dump has no exact target path row")
    return rows


def _mapping_covers_load(load: dict[str, Any], row: dict[str, Any]) -> bool:
    if row["start"] % PAGE_SIZE != 0 or row["offset"] % PAGE_SIZE != 0 or row["length"] <= 0:
        return False
    expected_end = _page_down(load["vaddr"] + load["memsz"] + PAGE_SIZE - 1)
    if row["length"] < expected_end - load["vaddr_page"]:
        return False
    flags = load["flags"]
    permissions = row["permissions"]
    if flags & 4 and "r" not in permissions[:3]:
        return False
    if flags & 2 and "w" not in permissions[:3]:
        return False
    if flags & 1 and "x" not in permissions[:3]:
        return False
    return True


def _process_tree_from_dump(text: str, expected_pid: int | None) -> dict[str, Any]:
    fork_rows: list[dict[str, int]] = []
    fork_re = re.compile(r"PERF_RECORD_FORK\((\d+):(\d+)\):\((\d+):(\d+)\)")
    for line in text.splitlines():
        match = fork_re.search(line)
        if match is None:
            continue
        row = {"child_pid": int(match.group(1)), "child_tid": int(match.group(2)), "parent_pid": int(match.group(3)), "parent_tid": int(match.group(4))}
        fork_rows.append(row)
        if expected_pid is not None and (row["child_pid"] != expected_pid or row["child_tid"] != expected_pid):
            raise ProtocolError("perf process tree contains an undeclared child/TID")
    return {"fork_records": fork_rows, "unexpected_children": 0}


def parse_perf_script_strict(
    text: str,
    *,
    marker: str = MARKER,
    marker_dso: str | None = None,
    expected_pid: int,
    expected_cpu: int = CPU,
    require_event: str = "instructions:u",
    expected_period: int = PERIOD,
) -> dict[str, Any]:
    """Require every decoded sample to be the launcher PID/TID on CPU 31."""
    headers = 0
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("Lost samples:") or stripped in {"...", "[stack truncated]", "[truncated]"}:
            continue
        if line[:1].isspace():
            continue
        match = _SAMPLE_PID_TID.match(line)
        if match is None:
            if _base._HEADER.match(line) is not None:
                raise ProtocolError("perf sample header lacks explicit PID/TID format")
            # The retained parser gives the canonical malformed-line error for
            # unrelated non-header text.
            continue
        headers += 1
        pid = int(match.group("pid"))
        tid_text = match.group("tid")
        if tid_text is None or int(tid_text) != pid or pid != expected_pid:
            raise ProtocolError(f"perf sample PID/TID escaped launcher: {pid}/{tid_text} != {expected_pid}")
        if int(match.group("cpu")) != expected_cpu:
            raise ProtocolError(f"perf sample CPU escaped pinned CPU: {match.group('cpu')} != {expected_cpu}")
    if headers == 0:
        raise ProtocolError("strict perf decode has no PID/TID sample headers")
    parsed = parse_perf_script(text, marker=marker, marker_dso=marker_dso,
                               expected_pid=expected_pid, expected_cpu=expected_cpu,
                               require_event=require_event)
    for sample in parsed["samples"]:
        if sample["period"] != expected_period:
            raise ProtocolError(f"perf sample period changed: {sample['period']} != {expected_period}")
    explicit_truncation = sum(
        1 for line in text.splitlines()
        if line.strip() in {"...", "[stack truncated]", "[truncated]"}
        or "stack truncated" in line.strip().lower()
    )
    empty_callchains = sum(1 for sample in parsed["samples"] if not sample["frames"])
    parsed["explicit_stack_truncated_count"] = explicit_truncation
    parsed["empty_callchain_count"] = empty_callchains
    parsed["stack_truncated"] = explicit_truncation
    parsed.pop("exclusive_deepest_resolved_period", None)
    parsed.pop("exclusive_leaf_resolved_period", None)
    parsed.pop("exclusive_root_resolved_period", None)
    parsed.pop("inclusive_marker_ancestor_period", None)
    def owner_key(frame: dict[str, Any]) -> str:
        return f"{frame['dso']}::{frame['symbol_base']}"
    physical: dict[str, int] = {}
    proxy: dict[str, int] = {}
    inclusive: dict[str, int] = {}
    root_owner: dict[str, int] = {}
    unknown_period = 0
    unknown_samples = 0
    marker_unknown_leaf_samples = 0
    marker_no_resolved_samples = 0
    marker_sample_set = {
        index for index, sample in enumerate(parsed["samples"])
        if any(frame["symbol_base"] == marker and (marker_dso is None or frame["dso"] == marker_dso) for frame in sample["frames"])
    }
    for index, sample in enumerate(parsed["samples"]):
        # Owner-period diagnostics are scoped to marker-containing samples.
        # The report's marker_period is their denominator; counting frame 0
        # over every decoded sample would make the reducer publish a ratio
        # greater than one when non-marker samples are present.
        if index not in marker_sample_set:
            continue
        frames = sample["frames"]
        if not frames or not frames[0]["resolved"]:
            unknown_period += sample["period"]
            unknown_samples += 1
            marker_unknown_leaf_samples += 1
        else:
            owner = owner_key(frames[0])
            physical[owner] = physical.get(owner, 0) + sample["period"]
        resolved = [frame for frame in frames if frame["resolved"]]
        if resolved:
            proxy_owner = owner_key(resolved[0])
            proxy[proxy_owner] = proxy.get(proxy_owner, 0) + sample["period"]
            root_key = owner_key(resolved[-1])
            root_owner[root_key] = root_owner.get(root_key, 0) + sample["period"]
            for frame in {owner_key(frame) for frame in resolved}:
                inclusive[frame] = inclusive.get(frame, 0) + sample["period"]
        else:
            marker_no_resolved_samples += 1
    physical_classified_period = sum(physical.values()) + unknown_period
    if physical_classified_period != parsed["marker_period"]:
        raise ProtocolError(
            "physical leaf marker-period conservation failed: "
            f"{physical_classified_period} != {parsed['marker_period']}"
        )
    parsed["physical_leaf_owner_period"] = physical
    parsed["physical_leaf_unknown_period"] = unknown_period
    parsed["physical_leaf_unknown_samples"] = unknown_samples
    parsed["marker_unknown_physical_leaf_samples"] = marker_unknown_leaf_samples
    parsed["marker_no_resolved_frame_samples"] = marker_no_resolved_samples
    parsed["deepest_resolved_proxy_period"] = proxy
    parsed["inclusive_marker_ancestor_period"] = inclusive
    parsed["root_resolved_owner_period"] = root_owner
    parsed["owner_key_format"] = "<DSO>::<symbol>"
    parsed["physical_leaf_scope"] = "marker-containing samples only; frame-0 physical leaf or explicit unknown"
    parsed["physical_leaf_classified_period"] = physical_classified_period
    parsed["physical_leaf_conservation"] = {
        "marker_period": parsed["marker_period"],
        "classified_period": physical_classified_period,
        "passed": True,
        "scope": "marker-containing samples only",
    }
    parsed["marker_ancestry"] = "observed" if parsed["marker_sample_count"] else "no observed marker ancestry"
    parsed["marker_period_scope"] = "decoder period coverage; not an exact timed-region attribution"
    return parsed


def derive_target_layout(
    perf_data: Path,
    binary: Path,
    *,
    perf_path: Path = Path(PERF_PATH),
    expected_pid: int | None = None,
) -> dict[str, Any]:
    """Derive the DWFL and perf-comparator bases from one raw child's MMAP2/ELF records."""
    if not perf_data.is_file() or not binary.is_file():
        raise ProtocolError("layout inputs are missing")
    elf = _parse_elf_loads(binary)
    command = [str(perf_path), "script", "-D", "-i", str(perf_data)]
    try:
        cp = subprocess.run(command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=ENVIRONMENT, timeout=300, check=False)
    except OSError as exc:
        raise ProtocolError(f"perf MMAP2 dump launch failed: {exc}") from exc
    if cp.returncode != 0:
        raise ProtocolError(f"perf MMAP2 dump failed: {cp.stderr[-500:]}")
    rows = _parse_mmap_dump(cp.stdout, binary_path=str(binary.absolute()), build_id=BINARY_BUILD_ID)
    process_tree = _process_tree_from_dump(cp.stdout, expected_pid)
    if expected_pid is not None and any(row["pid"] != expected_pid for row in rows):
        raise ProtocolError("target MMAP2 row PID differs from launcher exec PID")
    selected: list[dict[str, Any]] = []
    for load in elf["pt_loads"]:
        candidates = [row for row in rows if row["offset"] == load["offset_page"] and _mapping_covers_load(load, row)]
        if not candidates:
            raise ProtocolError(f"missing target MMAP2 row for PT_LOAD offset {_hex(load['offset_page'])}")
        # A PT_LOAD can have a repeated mapping after lifecycle reuse. Use the
        # first mapping in the raw stream; require every matching candidate to
        # imply the same ELF bias, so ambiguity cannot silently pass.
        biases = {row["start"] - load["vaddr_page"] for row in candidates}
        if len(biases) != 1:
            raise ProtocolError(f"inconsistent target MMAP2 bias at PT_LOAD offset {_hex(load['offset_page'])}")
        row = candidates[0]
        selected.append({"load": load, "mmap": row})
    biases = {item["mmap"]["start"] - item["load"]["vaddr_page"] for item in selected}
    if len(biases) != 1:
        raise ProtocolError("PT_LOAD/MMAP2 rows do not yield one consistent ELF load bias")
    corrected_base = next(iter(biases))
    exec_loads = [item for item in selected if item["load"]["flags"] & 0x1]
    if len(exec_loads) != 1:
        raise ProtocolError("expected exactly one executable PT_LOAD mapping")
    executable = exec_loads[0]
    executable_maps = [row for row in rows if "x" in row["permissions"][:3]]
    if len(executable_maps) != 1:
        raise ProtocolError("expected exactly one target executable RX MMAP2 row")
    if executable_maps[0] != executable["mmap"]:
        raise ProtocolError("executable PT_LOAD did not select the unique target RX map")
    perf_expected_start = executable["mmap"]["start"] - executable["mmap"]["offset"]
    if perf_expected_start <= 0:
        raise ProtocolError("derived perf comparator start is invalid")
    return {
        "schema": LAYOUT_SCHEMA,
        "binary": file_binding(binary),
        "perf_data": file_binding(perf_data),
        "perf_mmap_dump": {"argv": command, "returncode": cp.returncode, "sha256": hashlib.sha256(cp.stdout.encode()).hexdigest(), "bytes": len(cp.stdout.encode())},
        "elf": elf,
        "target_mmap2": [item["mmap"] for item in selected],
        "executable_mapping": executable["mmap"],
        "process_tree": process_tree,
        "corrected_base": _hex(corrected_base),
        "perf_expected_start": _hex(perf_expected_start),
        "formulas": {
            "corrected_base": "selected_mmap.start - page_down(PT_LOAD.p_vaddr); equal for every PT_LOAD",
            "perf_expected_start": "executable_selected_mmap.start - executable_selected_mmap.offset",
        },
    }


def check_layout(layout: dict[str, Any], *, perf_data: Path, binary: Path, expected_pid: int | None = None) -> None:
    if layout.get("schema") != LAYOUT_SCHEMA:
        raise ProtocolError("layout schema mismatch")
    if layout.get("binary") != file_binding(binary) or layout.get("perf_data") != file_binding(perf_data):
        raise ProtocolError("layout input custody mismatch")
    fresh = derive_target_layout(perf_data, binary, expected_pid=expected_pid)
    for key in ("corrected_base", "perf_expected_start", "target_mmap2", "executable_mapping", "process_tree", "elf"):
        if layout.get(key) != fresh.get(key):
            raise ProtocolError(f"layout field changed on full reparse: {key}")


def _require_binding(plan: dict[str, Any], key: str) -> None:
    row = plan.get(key)
    if not isinstance(row, dict):
        raise ProtocolError(f"missing plan binding {key}")
    check_binding(row, key)


def _same_binding(row: dict[str, Any], path: str, digest: str, size: int | None = None) -> bool:
    return row.get("path") == path and row.get("sha256") == digest and (size is None or row.get("bytes") == size)


def validate_report(report: dict[str, Any]) -> None:
    # The fixed cell/correctness validator is retained, but the event period
    # and call-graph settings belong to perf, not the probe report.
    _base.validate_report(report)


def check_plan_bindings(plan: dict[str, Any], plan_path: Path | None = None) -> None:
    if plan.get("schema") != PLAN_SCHEMA or plan.get("status") != "pre_capture":
        raise ProtocolError("plan schema/status mismatch")
    if plan.get("no_build_or_workload") is not True:
        raise ProtocolError("plan must prohibit build/workload before GO")
    execution = plan.get("execution") or {}
    if execution != {"serial": True, "retry": False, "timeout_seconds": 900, "children": CHILDREN}:
        raise ProtocolError("execution contract drift")
    matrix = plan.get("matrix") or {}
    if matrix.get("children") != CHILDREN or matrix.get("warmups") != WARMUPS or matrix.get("samples") != SAMPLES or matrix.get("cpu") != CPU or matrix.get("serial_order") != list(range(1, CHILDREN + 1)):
        raise ProtocolError("matrix drift")
    if matrix.get("cells") != [{"case": "docfloat", "edits": 1, "retention": "default"}]:
        raise ProtocolError("matrix must contain only fixed Float/e1/default cell")
    perf = plan.get("perf") or {}
    if perf.get("path") != PERF_PATH or perf.get("sha256") != PERF_SHA or perf.get("event") != "instructions:u" or perf.get("period") != PERIOD or perf.get("callgraph") != CALLGRAPH or perf.get("stack_cap") != STACK_CAP:
        raise ProtocolError("perf contract drift")
    if perf.get("sample_flags") != ["--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--user-callchains"]:
        raise ProtocolError("perf sample flags drift")
    marker = plan.get("marker") or {}
    if marker != {"symbol": MARKER, "dso_path": BINARY_PATH, "build_id": BINARY_BUILD_ID, "calls": CALLS, "p50_scope": "80 measured samples only", "marker_scope": "5 warmups + 80 measured calls combined", "owner_scope": "all 12 children; missing owner membership unknown"}:
        raise ProtocolError("marker scope drift")
    pilot = plan.get("pilot_gate") or {}
    if pilot != {"child": 1, "min_marker_samples": 100, "lost_events": 0, "marker_period_share_positive": True, "on_failure": "stop_without_retry"}:
        raise ProtocolError("pilot gate drift")
    if plan.get("child_gate") != {"all_children": True, "min_marker_samples": 100, "lost_events": 0, "marker_period_share_positive": True, "on_failure": "stop_without_retry"}:
        raise ProtocolError("all-child gate drift")
    if plan.get("environment") != ENVIRONMENT:
        raise ProtocolError("environment drift")
    synthetic_boundary = (plan.get("synthetic_boundary") or {}).get("enabled") is True
    required_bindings = ("binary", "fixture", "source_manifest", "source_file_manifest", "source_proof", "proposal", "design", "metric_proposal", "adr_revalidation", "validator", "launcher", "python", "taskset", "perf", "shim_source", "shim_version_script", "helper_seal", "review_context", "review_findings", "libdw", "libelf", "compiler", "header")
    if not synthetic_boundary:
        required_bindings += ("aslr_control", "synthetic_e2e", "design_precedence")
    for key in required_bindings:
        _require_binding(plan, key)
    binary = plan["binary"]
    if not _same_binding(binary, BINARY_PATH, BINARY_SHA, BINARY_BYTES) or binary.get("build_id") != BINARY_BUILD_ID:
        raise ProtocolError("retained binary binding drift")
    fixture = plan["fixture"]
    if not _same_binding(fixture, FIXTURE_PATH, FIXTURE_SHA, FIXTURE_BYTES):
        raise ProtocolError("fixture binding drift")
    if not _same_binding(plan["source_manifest"], SOURCE_MANIFEST_PATH, SOURCE_MANIFEST_SHA) or not _same_binding(plan["source_file_manifest"], SOURCE_FILE_MANIFEST_PATH, SOURCE_FILE_MANIFEST_SHA) or plan["source_proof"].get("sha256") != SOURCE_PROOF_SHA:
        raise ProtocolError("source custody drift")
    if plan["adr_revalidation"].get("path") != "/home/zhuhe/litchi-performance-0973/receipts/adr-goal-revalidation-0973-v1.json":
        raise ProtocolError("0973 ADR receipt path drift")
    proposal = plan["proposal"]
    if proposal.get("path") != "/home/zhuhe/litchi-performance-0973/diagnostic-proposal-v1.md":
        raise ProtocolError("proposal path drift")
    if not _same_binding(plan["design"], DESIGN_PATH, DESIGN_SHA):
        raise ProtocolError("final design binding drift")
    if not _same_binding(plan["metric_proposal"], METRIC_PROPOSAL_PATH, METRIC_PROPOSAL_SHA):
        raise ProtocolError("metric proposal binding drift")
    if not synthetic_boundary and not _same_binding(plan["design_precedence"], DESIGN_PRECEDENCE_PATH, DESIGN_PRECEDENCE_SHA):
        raise ProtocolError("design precedence clarification binding drift")
    if plan["validator"].get("path") != VALIDATOR_PATH or plan["validator"].get("sha256") != VALIDATOR_SHA:
        raise ProtocolError("retained validator binding drift")
    if plan["launcher"].get("path") != LAUNCHER_PATH or plan["launcher"].get("sha256") != LAUNCHER_SHA:
        raise ProtocolError("launcher binding drift")
    if not _same_binding(plan["libdw"], LIBDW_PATH, LIBDW_SHA, LIBDW_BYTES) or plan["libdw"].get("version") != LIBDW_VERSION:
        raise ProtocolError("runtime libdw binding drift")
    if not _same_binding(plan["libelf"], LIBELF_PATH, LIBELF_SHA, LIBELF_BYTES) or plan["libelf"].get("version") != LIBELF_VERSION:
        raise ProtocolError("runtime libelf binding drift")
    if not _same_binding(plan["compiler"], COMPILER_PATH, COMPILER_SHA):
        raise ProtocolError("compiler binding drift")
    if not _same_binding(plan["header"], HEADER_PATH, HEADER_SHA):
        raise ProtocolError("libdwfl header binding drift")
    if plan.get("shim_build", {}).get("argv") != SHIM_BUILD_COMMAND:
        raise ProtocolError("shim build command drift")
    if plan["python"].get("path") != PYTHON_PATH or plan["taskset"].get("path") != TASKSET_PATH:
        raise ProtocolError("python/taskset binding drift")
    if plan.get("report_contract") != {"case": "docfloat", "edits": 1, "retention": "default_8mib", "warmups": WARMUPS, "samples_requested": SAMPLES, "sample_indices": list(range(SAMPLES)), "elapsed_ns_identity": "samples[i].elapsed_ns == elapsed_ns[i]"}:
        raise ProtocolError("report contract drift")
    if plan.get("raw_artifacts") != RAW_ARTIFACTS:
        raise ProtocolError("raw artifact contract drift")
    raw_contract = plan.get("raw_stack_contract") or {}
    expected_raw_contract = {"schema": RAW_STACK_SCHEMA, "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER", "period": PERIOD, "stack_cap": STACK_CAP, "bounds": "0 <= dyn_size <= captured_stack_size <= 32768", "missing_or_short": "missing fields reject; bounded short captured_stack_size is retained in its distribution", "raw_scope": "all PERF_RECORD_SAMPLE records without misc filtering"}
    if synthetic_boundary:
        # The author-side fixture predates the final wording correction, but
        # remains explicitly bounded to the offline perf-boundary mock.  The
        # real pre-capture plan below is exact and uses expected_raw_contract.
        synthetic_raw_contract = dict(expected_raw_contract)
        synthetic_raw_contract["missing_or_short"] = "recorded as quality findings; any violation rejects that child"
        if raw_contract not in (expected_raw_contract, synthetic_raw_contract):
            raise ProtocolError("raw STACK_USER contract drift")
    elif raw_contract != expected_raw_contract:
        raise ProtocolError("raw STACK_USER contract drift")
    if plan.get("process_policy") != {"require_sample_pid_tid_equal_exec_pid": True, "reject_any_other_thread_or_descendant": True, "sampling_monitor": False}:
        raise ProtocolError("process policy drift")
    if plan.get("native_environment") != {"LD_PRELOAD": "absent", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}:
        raise ProtocolError("native environment policy drift")
    dyn = plan.get("dynamic_layout") or {}
    if dyn.get("source") != "exact per-child perf MMAP2 + ELF64 PT_LOAD" or dyn.get("hardcoded_aslr_addresses") is not False:
        raise ProtocolError("dynamic layout contract drift")
    shim = plan.get("shim") or {}
    if shim.get("target_only") is not True or shim.get("offline_only") is not True or shim.get("env_values") != ["LITCHI_0973_TARGET_PATH", "LITCHI_0973_TARGET_NAME", "LITCHI_0973_CORRECTED_BASE", "LITCHI_0973_PERF_EXPECTED_START"] or shim.get("native_child_preload") != "absent":
        raise ProtocolError("shim scope drift")
    check_binding(shim.get("shared_object"), "shim shared object")
    if shim.get("shared_object", {}).get("path") != plan.get("helpers", {}).get("shim_so", {}).get("path") or shim.get("shared_object", {}).get("sha256") != plan.get("helpers", {}).get("shim_so", {}).get("sha256"):
        raise ProtocolError("shim shared object/helper binding mismatch")
    seal = load_json(Path(plan["helper_seal"]["path"]))
    if seal.get("schema") != HELPER_SCHEMA or seal.get("status") != "sealed":
        raise ProtocolError("helper seal not sealed")
    if plan.get("helpers_seal_sha256") != plan["helper_seal"].get("sha256"):
        raise ProtocolError("helper seal checksum binding missing")
    if set(plan.get("helpers", {})) != {"protocol", "runner", "reducer", "plan_generator", "raw_stack", "shim_source", "shim_map", "shim_so"}:
        raise ProtocolError("helper set drift")
    for name, row in seal.get("helpers", {}).items():
        if name not in plan["helpers"] or plan["helpers"][name].get("sha256") != row.get("sha256") or plan["helpers"][name].get("path") != row.get("path"):
            raise ProtocolError(f"helper binding mismatch: {name}")
        actual = file_binding(row["path"])
        if actual != row:
            raise ProtocolError(f"helper changed after seal: {name}")
    if not _same_binding(plan["review_findings"], REVIEW_FINDINGS_PATH, REVIEW_FINDINGS_SHA):
        raise ProtocolError("0973 review findings path drift")
    if not synthetic_boundary:
        # The receipt is deliberately live-bound by the plan row.  Keeping a
        # receipt digest in this helper would make the decoder/helper and
        # ASLR-control receipt mutually circular: the receipt records the
        # exact protocol helper it exercised.  check_binding above still
        # verifies the plan's path/hash/bytes row against disk.
        if plan["aslr_control"].get("path") != ASLR_CONTROL_PATH:
            raise ProtocolError("translated ASLR control binding drift")
        aslr = load_json(Path(plan["aslr_control"]["path"]))
        if aslr.get("status") != "pass" or aslr.get("source_changed") is not False:
            raise ProtocolError("translated ASLR control is not a passing offline control")
        if aslr.get("inputs", {}).get("protocol_helper") != plan.get("helpers", {}).get("protocol"):
            raise ProtocolError("translated ASLR control did not exercise the sealed protocol helper")
        if plan["synthetic_e2e"].get("path") != SYNTHETIC_E2E_PATH:
            raise ProtocolError("synthetic E2E binding drift")
        synthetic_e2e = load_json(Path(plan["synthetic_e2e"]["path"]))
        if synthetic_e2e.get("status") != "pass" or synthetic_e2e.get("workload_started") is not False or synthetic_e2e.get("build_started") is not False:
            raise ProtocolError("synthetic E2E control is not a passing offline control")
        admission = synthetic_e2e.get("admission") or {}
        if admission.get("all_twelve_packets_loaded") is not True or admission.get("all_twelve_reduced") is not True or admission.get("all_twelve_marker_gate_passed") is not True:
            raise ProtocolError("synthetic E2E control did not exercise all twelve packets")
    if plan_path is not None and plan.get("plan_path") not in {None, str(plan_path.resolve())}:
        raise ProtocolError("plan self path mismatch")


def check_go(go: dict[str, Any], plan: dict[str, Any], plan_path: Path, go_path: Path) -> None:
    if go.get("schema") != GO_SCHEMA or go.get("authorized") is not True or go.get("status") != "authorized":
        raise ProtocolError("root GO is not authorized")
    if (plan.get("synthetic_boundary") or {}).get("enabled") is True:
        raise ProtocolError("synthetic_boundary is test-only and cannot be authorized for capture")
    row = go.get("plan") or {}
    if row.get("path") != str(plan_path.resolve()) or row.get("sha256") != sha256(plan_path) or row.get("bytes") != plan_path.stat().st_size:
        raise ProtocolError("GO plan binding mismatch")
    if (go.get("quiescence") or {}).get("status") != "pass":
        raise ProtocolError("GO quiescence missing")
    for name in ("helper_seal", "independent_review"):
        if not isinstance(go.get(name), dict):
            raise ProtocolError(f"GO {name} missing")
        check_binding(go[name], f"GO {name}")
    if go["helper_seal"] != plan["helper_seal"]:
        raise ProtocolError("GO helper seal mismatch")
    seal = load_json(Path(plan["helper_seal"]["path"]))
    for name, row in seal.get("helpers", {}).items():
        if file_binding(row["path"]) != row:
            raise ProtocolError(f"GO helper changed after seal: {name}")
    if go.get("binary") != plan["binary"] or go.get("fixture") != plan["fixture"]:
        raise ProtocolError("GO input binding mismatch")
    for name in ("libdw", "libelf", "compiler", "header", "shim_source", "shim_version_script"):
        if go.get(name) != plan[name]:
            raise ProtocolError(f"GO decoder dependency binding mismatch: {name}")
    if go.get("shim_build") != plan.get("shim_build"):
        raise ProtocolError("GO shim build command mismatch")
    host = go.get("host") or {}
    audited = host.get("audited_allowed_affinity")
    if not isinstance(audited, list) or any(type(cpu) is not int for cpu in audited) or CPU not in audited:
        raise ProtocolError("GO host affinity audit missing selected CPU")
    if go.get("environment") != ENVIRONMENT or host.get("selected_cpu") != CPU:
        raise ProtocolError("GO host/environment binding mismatch")
    review = load_json(Path(go["independent_review"]["path"]))
    if review.get("status") not in {"pass", "passed"}:
        raise ProtocolError("independent review is not passing")


def check_child_layout_artifacts(packet: dict[str, Any], plan: dict[str, Any], packet_path: Path) -> None:
    checked = check_artifact_bindings_0973(packet, packet_path)
    layout = load_json(checked["layout"])
    check_layout(layout, perf_data=checked["perf_data"], binary=Path(plan["binary"]["path"]))
    if packet.get("layout") != layout:
        raise ProtocolError("packet/layout receipt mismatch")
    for key in ("corrected_base", "perf_expected_start"):
        if packet.get("shim_config", {}).get(key) != layout[key]:
            raise ProtocolError(f"shim config/layout mismatch: {key}")
