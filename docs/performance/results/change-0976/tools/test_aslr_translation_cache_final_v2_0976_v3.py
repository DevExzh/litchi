#!/usr/bin/env python3
"""Offline ASLR translation control for the sealed 0975 child-01 capture.

This helper copies the perf data file into a fresh external directory and
changes only target-DSO mapping starts and target addresses carried in the
sample records.  It then replays the fresh 0976 final corrected DWFL interface
against the translated copy.  The sealed input is never opened for writing.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import importlib.util
import json
import os
import struct
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


ROOT = Path("/home/zhuhe/litchi-performance-0976")
OUT = ROOT / "aslr-cache-final-v2-0975-v3"
SCRIPT_PATH = ROOT / "tools/test_aslr_translation_cache_final_v2_0976_v3.py"
RECEIPT_PATH = ROOT / "receipts/aslr-cache-final-v2-0975-v3.json"
RAW = Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.perf.data")
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
SHIM = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.so"
PROTOCOL_PATH = Path("/home/zhuhe/litchi-performance-0973/tools-v2/protocol_common_0973_v2.py")
SHIM_SOURCE = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.c"
SHIM_MAP = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.map"
RETAINED_LEGACY_SCRIPT = Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.corrected.perf-script-no-inline.txt")
ORIGINAL_LAYOUT = Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.target-layout.json")
ORIGINAL_MMAP_DUMP = Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.perf-mmap-dump.txt")
TARGET_PATH = str(BINARY.resolve())
TARGET_BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
TARGET_SHA256 = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
EXPECTED_PID = 3763764
EXPECTED_CPU = 31
PERIOD = 1_000_000
PAGE = 4096
DELTA = 0x1000000000

# perf_event.h values used by this retained file.
PERF_RECORD_MMAP = 1
PERF_RECORD_MMAP2 = 10
PERF_RECORD_SAMPLE = 9
PERF_SAMPLE_IP = 1 << 0
PERF_SAMPLE_TID = 1 << 1
PERF_SAMPLE_TIME = 1 << 2
PERF_SAMPLE_ADDR = 1 << 3
PERF_SAMPLE_READ = 1 << 4
PERF_SAMPLE_CALLCHAIN = 1 << 5
PERF_SAMPLE_ID = 1 << 6
PERF_SAMPLE_CPU = 1 << 7
PERF_SAMPLE_PERIOD = 1 << 8
PERF_SAMPLE_STREAM_ID = 1 << 9
PERF_SAMPLE_RAW = 1 << 10
PERF_SAMPLE_BRANCH_STACK = 1 << 11
PERF_SAMPLE_REGS_USER = 1 << 12
PERF_SAMPLE_STACK_USER = 1 << 13
PERF_SAMPLE_WEIGHT = 1 << 14
PERF_SAMPLE_DATA_SRC = 1 << 15
PERF_SAMPLE_IDENTIFIER = 1 << 16
PERF_SAMPLE_TRANSACTION = 1 << 17
PERF_SAMPLE_REGS_INTR = 1 << 18
PERF_SAMPLE_PHYS_ADDR = 1 << 19
PERF_SAMPLE_AUX = 1 << 20
PERF_SAMPLE_CGROUP = 1 << 21
PERF_SAMPLE_DATA_PAGE_SIZE = 1 << 22
PERF_SAMPLE_CODE_PAGE_SIZE = 1 << 23
PERF_SAMPLE_WEIGHT_STRUCT = 1 << 24
EXPECTED_SAMPLE_TYPE = (
    PERF_SAMPLE_IP | PERF_SAMPLE_TID | PERF_SAMPLE_TIME | PERF_SAMPLE_CALLCHAIN
    | PERF_SAMPLE_CPU | PERF_SAMPLE_PERIOD | PERF_SAMPLE_REGS_USER
    | PERF_SAMPLE_STACK_USER
)
KNOWN_SAMPLE_BITS = (
    PERF_SAMPLE_IP | PERF_SAMPLE_TID | PERF_SAMPLE_TIME | PERF_SAMPLE_ADDR
    | PERF_SAMPLE_READ | PERF_SAMPLE_CALLCHAIN | PERF_SAMPLE_ID | PERF_SAMPLE_CPU
    | PERF_SAMPLE_PERIOD | PERF_SAMPLE_STREAM_ID | PERF_SAMPLE_RAW
    | PERF_SAMPLE_BRANCH_STACK | PERF_SAMPLE_REGS_USER | PERF_SAMPLE_STACK_USER
    | PERF_SAMPLE_WEIGHT | PERF_SAMPLE_DATA_SRC | PERF_SAMPLE_IDENTIFIER
    | PERF_SAMPLE_TRANSACTION | PERF_SAMPLE_REGS_INTR | PERF_SAMPLE_PHYS_ADDR
    | PERF_SAMPLE_AUX | PERF_SAMPLE_CGROUP | PERF_SAMPLE_DATA_PAGE_SIZE
    | PERF_SAMPLE_CODE_PAGE_SIZE | PERF_SAMPLE_WEIGHT_STRUCT
)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load_protocol() -> Any:
    path = PROTOCOL_PATH
    spec = importlib.util.spec_from_file_location("litchi_protocol_common_0973_v2_aslr", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load protocol helper {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def u64(data: bytearray | bytes, offset: int) -> int:
    return struct.unpack_from("<Q", data, offset)[0]


def put_u64(data: bytearray, offset: int, value: int) -> None:
    struct.pack_into("<Q", data, offset, value)


def iter_records(data: bytearray | bytes) -> Iterable[tuple[int, int, int, int]]:
    data_offset, data_size = struct.unpack_from("<QQ", data, 0x28)
    end = data_offset + data_size
    if end > len(data):
        raise ValueError("perf data section exceeds input")
    pos = data_offset
    while pos < end:
        if pos + 8 > end:
            raise ValueError(f"truncated perf record header at 0x{pos:x}")
        typ, misc, size = struct.unpack_from("<IHH", data, pos)
        if size < 8 or pos + size > end:
            raise ValueError(f"invalid perf record size {size} at 0x{pos:x}")
        yield pos, typ, misc, size
        pos += size
    if pos != end:
        raise ValueError("perf data section did not terminate on a record boundary")


def mmap_path(data: bytes, pos: int, typ: int, size: int) -> str:
    # MMAP2 has pid/tid, addr/len/pgoff, device/inode/build-id/prot/flags,
    # followed by the NUL-terminated path.  MMAP has no device/inode/build-id.
    path_offset = 72 if typ == PERF_RECORD_MMAP2 else 40
    if path_offset >= size:
        raise ValueError(f"mapping record too short at 0x{pos:x}")
    end = data.find(b"\0", pos + path_offset, pos + size)
    if end < 0:
        raise ValueError(f"mapping path is not NUL terminated at 0x{pos:x}")
    return data[pos + path_offset:end].decode("utf-8", "surrogateescape")


def address_ranges(rows: list[dict[str, Any]]) -> list[tuple[int, int]]:
    ranges = sorted((int(row["start"]), int(row["start"]) + int(row["length"])) for row in rows)
    merged: list[list[int]] = []
    for start, end in ranges:
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [(start, end) for start, end in merged]


def in_ranges(value: int, ranges: list[tuple[int, int]]) -> bool:
    return any(start <= value < end for start, end in ranges)


def change(changes: list[dict[str, Any]], data: bytearray, offset: int, value: int, *, kind: str, record: int) -> None:
    old = u64(data, offset)
    if old == value:
        raise AssertionError(f"requested no-op address change for {kind} at 0x{offset:x}")
    put_u64(data, offset, value)
    changes.append({
        "record_offset": f"0x{record:x}",
        "byte_offset": f"0x{offset:x}",
        "byte_length": 8,
        "kind": kind,
        "old": f"0x{old:x}",
        "new": f"0x{value:x}",
    })


def parse_header(data: bytes) -> dict[str, Any]:
    if data[:8] != b"PERFILE2":
        raise ValueError("unexpected perf.data magic")
    attrs_offset, attrs_size = struct.unpack_from("<QQ", data, 0x18)
    data_offset, data_size = struct.unpack_from("<QQ", data, 0x28)
    if attrs_offset + attrs_size > len(data) or data_offset + data_size > len(data):
        raise ValueError("perf.data section outside file")
    attr = data[attrs_offset:attrs_offset + attrs_size]
    if len(attr) < 96:
        raise ValueError("perf event attribute is too short")
    sample_type = u64(attr, 24)
    regs_mask = u64(attr, 80)
    stack_size = struct.unpack_from("<I", attr, 88)[0]
    return {
        "attrs_offset": attrs_offset,
        "attrs_size": attrs_size,
        "data_offset": data_offset,
        "data_size": data_size,
        "sample_type": sample_type,
        "sample_regs_user": regs_mask,
        "sample_stack_user": stack_size,
    }


def raw_sample_headers(data: bytes, header: dict[str, Any]) -> list[dict[str, Any]]:
    """Return the raw sample record identity fields before DWFL decoding."""
    rows: list[dict[str, Any]] = []
    for pos, typ, misc, size in iter_records(data):
        if typ != PERF_RECORD_SAMPLE:
            continue
        # Parse the complete fixed ABI so a future record-layout drift cannot
        # make the period offset look valid by accident.
        sample_layout(data, pos, size, header["sample_type"], header["sample_regs_user"])
        rows.append({
            "offset": f"0x{pos:x}",
            "misc": misc,
            "size": size,
            "period": u64(data, pos + 40),
        })
    return rows


def raw_target_mmaps(data: bytes) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for pos, typ, _misc, size in iter_records(data):
        if typ not in (PERF_RECORD_MMAP, PERF_RECORD_MMAP2):
            continue
        path = mmap_path(data, pos, typ, size)
        if path != TARGET_PATH:
            continue
        rows.append({
            "offset": f"0x{pos:x}",
            "type": typ,
            "start": u64(data, pos + 16),
            "length": u64(data, pos + 24),
            "pgoff": u64(data, pos + 32),
            "path": path,
        })
    return rows


def sample_layout(data: bytearray, pos: int, size: int, sample_type: int, regs_mask: int) -> dict[str, Any]:
    if sample_type & ~KNOWN_SAMPLE_BITS:
        raise ValueError(f"unsupported sample_type bits: 0x{sample_type & ~KNOWN_SAMPLE_BITS:x}")
    q = pos + 8
    end = pos + size
    out: dict[str, Any] = {"ip": None, "ip_offset": None, "callchain": [], "regs": [], "stack": None}

    def take_u64() -> tuple[int, int]:
        nonlocal q
        if q + 8 > end:
            raise ValueError(f"sample field exceeds record at 0x{pos:x}")
        offset = q
        value = u64(data, q)
        q += 8
        return offset, value

    if sample_type & PERF_SAMPLE_IP:
        offset, value = take_u64()
        out["ip"] = value
        out["ip_offset"] = offset
    if sample_type & PERF_SAMPLE_TID:
        if q + 8 > end:
            raise ValueError("sample TID field exceeds record")
        q += 8
    if sample_type & PERF_SAMPLE_TIME:
        take_u64()
    if sample_type & PERF_SAMPLE_ADDR:
        take_u64()
    if sample_type & PERF_SAMPLE_ID:
        take_u64()
    if sample_type & PERF_SAMPLE_STREAM_ID:
        take_u64()
    if sample_type & PERF_SAMPLE_CPU:
        if q + 8 > end:
            raise ValueError("sample CPU field exceeds record")
        q += 8
    if sample_type & PERF_SAMPLE_PERIOD:
        take_u64()
    if sample_type & PERF_SAMPLE_READ:
        raise ValueError("PERF_SAMPLE_READ is not supported by this fixed-input test")
    if sample_type & PERF_SAMPLE_CALLCHAIN:
        nr_offset, nr = take_u64()
        if nr > (end - q) // 8:
            raise ValueError(f"callchain length {nr} exceeds record at 0x{pos:x}")
        ips: list[tuple[int, int]] = []
        for _ in range(nr):
            offset, value = take_u64()
            ips.append((offset, value))
        out["callchain"] = ips
        out["callchain_nr_offset"] = nr_offset
    if sample_type & PERF_SAMPLE_RAW:
        if q + 4 > end:
            raise ValueError("raw sample length exceeds record")
        raw_size = struct.unpack_from("<I", data, q)[0]
        q += 4 + raw_size
        q = (q + 7) & ~7
    if sample_type & PERF_SAMPLE_BRANCH_STACK:
        raise ValueError("PERF_SAMPLE_BRANCH_STACK is not supported by this fixed-input test")
    if sample_type & PERF_SAMPLE_REGS_USER:
        abi_offset, abi = take_u64()
        regs: list[tuple[int, int, int]] = []
        if abi != 0:
            for bit in range(64):
                if regs_mask & (1 << bit):
                    offset, value = take_u64()
                    regs.append((bit, offset, value))
        out["regs_abi"] = abi
        out["regs_abi_offset"] = abi_offset
        out["regs"] = regs
    if sample_type & PERF_SAMPLE_STACK_USER:
        stack_size_offset, stack_size = take_u64()
        if q + stack_size > end:
            raise ValueError("stack bytes exceed record")
        stack_start = q
        q += stack_size
        q = (q + 7) & ~7
        dyn_offset, dyn_size = take_u64()
        out["stack"] = {
            "size_offset": stack_size_offset,
            "size": stack_size,
            "start": stack_start,
            "dyn_offset": dyn_offset,
            "dyn_size": dyn_size,
        }
    # The remaining sample-type fields are absent in the retained capture.  A
    # future format change must fail closed rather than mutate an unknown ABI.
    unsupported_tail = sample_type & (
        PERF_SAMPLE_WEIGHT | PERF_SAMPLE_DATA_SRC | PERF_SAMPLE_IDENTIFIER
        | PERF_SAMPLE_TRANSACTION | PERF_SAMPLE_REGS_INTR | PERF_SAMPLE_PHYS_ADDR
        | PERF_SAMPLE_AUX | PERF_SAMPLE_CGROUP | PERF_SAMPLE_DATA_PAGE_SIZE
        | PERF_SAMPLE_CODE_PAGE_SIZE | PERF_SAMPLE_WEIGHT_STRUCT
    )
    if unsupported_tail:
        raise ValueError(f"unsupported trailing sample fields: 0x{unsupported_tail:x}")
    if q != end:
        raise ValueError(f"sample parser ended at 0x{q:x}, record ends at 0x{end:x}")
    return out


def mutate_perf_data(original: bytes, layout: dict[str, Any]) -> tuple[bytes, dict[str, Any]]:
    data = bytearray(original)
    header = parse_header(original)
    if header["sample_type"] != EXPECTED_SAMPLE_TYPE:
        raise ValueError(f"sample_type drift: 0x{header['sample_type']:x}")
    if header["sample_regs_user"] != 0xFF0FFF or header["sample_stack_user"] != 32768:
        raise ValueError("retained sample register/stack ABI drift")
    original_sample_headers = raw_sample_headers(original, header)
    original_target_mmaps = raw_target_mmaps(original)

    target_rows = [row for row in layout["target_mmap2"] if row["path"] == TARGET_PATH]
    if not target_rows:
        raise ValueError("layout has no exact target mappings")
    ranges = address_ranges(target_rows)
    exec_ranges = address_ranges([row for row in target_rows if "x" in row["permissions"][:3]])
    if not exec_ranges:
        raise ValueError("layout has no target executable mapping")
    changes: list[dict[str, Any]] = []
    type_counts: dict[str, int] = {}
    sample_headers: list[dict[str, Any]] = []
    sample_periods: list[int] = []
    sample_ips: list[int] = []
    target_mmap_records = 0
    sample_ip_shifted = 0
    callchain_ip_shifted = 0
    register_pointer_shifted = 0
    stack_word_shifted = 0
    stack_exec_return_word_shifted = 0

    for pos, typ, misc, size in iter_records(original):
        type_counts[str(typ)] = type_counts.get(str(typ), 0) + 1
        if typ in (PERF_RECORD_MMAP, PERF_RECORD_MMAP2):
            path = mmap_path(original, pos, typ, size)
            if path == TARGET_PATH:
                addr_offset = pos + 16
                old_addr = u64(data, addr_offset)
                change(changes, data, addr_offset, old_addr + DELTA, kind="target_mmap_start", record=pos)
                target_mmap_records += 1
        elif typ == PERF_RECORD_SAMPLE:
            fields = sample_layout(data, pos, size, header["sample_type"], header["sample_regs_user"])
            ip = int(fields["ip"])
            sample_ips.append(ip)
            period_offset = pos + 8
            # The fixed sample ABI places the period at +40.  Find it from the
            # generic parsed fields to keep the identity check explicit.
            q = pos + 8
            q += 8  # IP
            q += 8  # TID
            q += 8  # TIME
            q += 8  # CPU (CALLCHAIN is after PERIOD in this capture)
            # The actual ABI order is IP,TID,TIME,CPU,PERIOD,CALLCHAIN.
            period_offset = pos + 40
            period = u64(data, period_offset)
            sample_periods.append(period)
            sample_headers.append({"offset": f"0x{pos:x}", "misc": misc, "size": size, "period": period})
            if in_ranges(ip, ranges):
                change(changes, data, int(fields["ip_offset"]), ip + DELTA, kind="sample_ip", record=pos)
                sample_ip_shifted += 1
            for offset, value in fields["callchain"]:
                if in_ranges(value, ranges):
                    change(changes, data, offset, value + DELTA, kind="callchain_ip", record=pos)
                    callchain_ip_shifted += 1
            for bit, offset, value in fields["regs"]:
                if in_ranges(value, ranges):
                    change(changes, data, offset, value + DELTA, kind=f"register_pointer_r{bit}", record=pos)
                    register_pointer_shifted += 1
            stack = fields["stack"]
            if stack is not None:
                stack_start = int(stack["start"])
                stack_size = int(stack["size"])
                if stack_size % 8 != 0:
                    raise ValueError("stack size is not qword aligned")
                for offset in range(stack_start, stack_start + stack_size, 8):
                    value = u64(data, offset)
                    if in_ranges(value, ranges):
                        change(changes, data, value_offset := offset, value + DELTA, kind="stack_target_pointer", record=pos)
                        stack_word_shifted += 1
                        if in_ranges(value, exec_ranges):
                            stack_exec_return_word_shifted += 1

    if target_mmap_records == 0:
        raise ValueError("no target mapping records were translated")
    if sample_ip_shifted == 0 or register_pointer_shifted == 0 or stack_word_shifted == 0 or stack_exec_return_word_shifted == 0:
        raise ValueError(
            "positive-control mutation did not touch all required classes: "
            f"sample_ip={sample_ip_shifted}, register={register_pointer_shifted}, "
            f"stack={stack_word_shifted}, stack_exec={stack_exec_return_word_shifted}"
        )
    if len(sample_periods) != 6940 or sum(sample_periods) != 6_940_000_000:
        raise ValueError("retained sample count/period identity drift before translation")
    translated_sample_headers = raw_sample_headers(data, header)
    if translated_sample_headers != original_sample_headers:
        raise AssertionError("translated raw sample headers/counts/periods differ")
    translated_target_mmaps = raw_target_mmaps(data)
    if len(translated_target_mmaps) != len(original_target_mmaps):
        raise AssertionError("translated target mapping record count differs")
    for before, after in zip(original_target_mmaps, translated_target_mmaps):
        if before["start"] + DELTA != after["start"] or before["length"] != after["length"] or before["pgoff"] != after["pgoff"] or before["path"] != after["path"]:
            raise AssertionError("translated target mapping metadata changed beyond the declared delta")
    changed_offsets = {int(row["byte_offset"], 16) for row in changes}
    changed_spans = set()
    for offset in changed_offsets:
        changed_spans.update(range(offset, offset + 8))
    # The output is a byte-for-byte copy except for the explicitly listed
    # qwords.  This catches accidental edits to perf headers/feature sections.
    changed_byte_positions = {index for index, (a, b) in enumerate(zip(original, data)) if a != b}
    changed_byte_count = len(changed_byte_positions)
    if not changed_byte_positions <= changed_spans:
        raise AssertionError("one or more changed bytes is outside the recorded qword fields")
    return bytes(data), {
        "input_header": header,
        "target_mapping_count": target_mmap_records,
        "target_ranges": [{"start": f"0x{start:x}", "end": f"0x{end:x}"} for start, end in ranges],
        "delta": f"0x{DELTA:x}",
        "record_type_counts": type_counts,
        "sample_header_count": len(sample_headers),
        "sample_header_sha256": sha256_bytes(json.dumps(original_sample_headers, sort_keys=True, separators=(",", ":")).encode()),
        "translated_sample_header_sha256": sha256_bytes(json.dumps(translated_sample_headers, sort_keys=True, separators=(",", ":")).encode()),
        "raw_sample_headers_invariant": True,
        "raw_target_mmaps_invariant_except_delta": True,
        "original_target_mmaps": original_target_mmaps,
        "translated_target_mmaps": translated_target_mmaps,
        "sample_period_sum": sum(sample_periods),
        "sample_ip_shifted": sample_ip_shifted,
        "callchain_ip_shifted": callchain_ip_shifted,
        "register_pointer_shifted": register_pointer_shifted,
        "stack_word_shifted": stack_word_shifted,
        "stack_exec_return_word_shifted": stack_exec_return_word_shifted,
        "changed_byte_count": changed_byte_count,
        "changed_field_count": len(changes),
        "changes": changes,
    }


def run_to_files(argv: list[str], env: dict[str, str], stdout: Path, stderr: Path) -> dict[str, Any]:
    if stdout.exists() or stderr.exists():
        raise FileExistsError(f"refusing to overwrite decoder output: {stdout} or {stderr}")
    with stdout.open("wb") as out, stderr.open("wb") as err:
        cp = subprocess.run(argv, stdin=subprocess.DEVNULL, stdout=out, stderr=err, env=env, check=False, timeout=600)
    return {"argv": argv, "returncode": cp.returncode, "stdout": binding(stdout), "stderr": binding(stderr)}


def canonical_samples(protocol: Any, text: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    parsed = protocol.parse_perf_script_strict(
        text,
        marker_dso=TARGET_PATH,
        expected_pid=EXPECTED_PID,
        expected_cpu=EXPECTED_CPU,
        expected_period=PERIOD,
    )
    canonical: list[dict[str, Any]] = []
    for sample in parsed["samples"]:
        canonical.append({
            "pid": sample["pid"],
            "cpu": sample["cpu"],
            "time": sample["time"],
            "period": sample["period"],
            "event": sample["event"],
            "frames": [
                {
                    "dso": frame["dso"],
                    "symbol_base": frame["symbol_base"],
                    "offset": frame["offset"],
                    "resolved": frame["resolved"],
                }
                for frame in sample["frames"]
            ],
        })
    return parsed, canonical


def canonical_hash(rows: list[dict[str, Any]]) -> str:
    encoded = json.dumps(rows, sort_keys=True, separators=(",", ":")).encode()
    return sha256_bytes(encoded)


def compare_sequences(protocol: Any, original_text: str, translated_text: str) -> dict[str, Any]:
    original, original_rows = canonical_samples(protocol, original_text)
    translated, translated_rows = canonical_samples(protocol, translated_text)
    if len(original_rows) != len(translated_rows):
        raise AssertionError("translated sample count differs from original")
    if original_rows != translated_rows:
        for index, (left, right) in enumerate(zip(original_rows, translated_rows)):
            if left != right:
                raise AssertionError(f"canonical frame sequence differs at sample {index}: {left} != {right}")
        raise AssertionError("canonical frame sequence differs")
    return {
        "original": {k: original[k] for k in ("sample_count", "total_period", "marker_sample_count", "marker_period", "lost_events", "empty_callchain_count", "explicit_stack_truncated_count", "unresolved_frames")},
        "translated": {k: translated[k] for k in ("sample_count", "total_period", "marker_sample_count", "marker_period", "lost_events", "empty_callchain_count", "explicit_stack_truncated_count", "unresolved_frames")},
        "canonical_sequence_count": len(original_rows),
        "canonical_sequence_sha256": canonical_hash(original_rows),
        "canonical_sequence_invariant": True,
        "original_marker_period": original["marker_period"],
        "translated_marker_period": translated["marker_period"],
    }


def main() -> int:
    started = utc()
    input_paths = (
        RAW, BINARY, SHIM, PROTOCOL_PATH, SHIM_SOURCE, SHIM_MAP, SCRIPT_PATH,
        RETAINED_LEGACY_SCRIPT, ORIGINAL_LAYOUT, ORIGINAL_MMAP_DUMP,
    )
    for path in input_paths:
        if not path.is_file():
            raise FileNotFoundError(path)
    if DELTA % PAGE != 0:
        raise AssertionError("translation delta is not page aligned")
    input_hashes_before = {str(path): binding(path) for path in input_paths}
    OUT.mkdir(parents=True, exist_ok=False)
    (ROOT / "receipts").mkdir(parents=True, exist_ok=True)

    protocol = load_protocol()
    raw = RAW.read_bytes()
    raw_binding = binding(RAW)
    binary_binding = binding(BINARY)
    shim_binding = binding(SHIM)
    legacy_script_binding = binding(RETAINED_LEGACY_SCRIPT)
    original_layout_binding = binding(ORIGINAL_LAYOUT)
    original_mmap_binding = binding(ORIGINAL_MMAP_DUMP)
    shim_source_path = SHIM_SOURCE
    shim_map_path = SHIM_MAP
    active_shim_build = [
        "cc", "-shared", "-fPIC", "-O2", "-std=c11", "-Wall", "-Wextra", "-Werror", "-pthread",
        "-I/home/zhuhe/litchi-performance-0972/header-v2/root/usr/include",
        "-Wl,--version-script=/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.map",
        "-o", "/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.so",
        "/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.c", "-ldl", "-pthread",
    ]
    derived_layout = protocol.derive_target_layout(RAW, BINARY, expected_pid=EXPECTED_PID)
    stored_layout = json.loads(ORIGINAL_LAYOUT.read_text(encoding="utf-8"))
    if derived_layout["corrected_base"] != stored_layout["corrected_base"] or derived_layout["perf_expected_start"] != stored_layout["perf_expected_start"]:
        raise AssertionError("fresh original layout disagrees with sealed layout")
    if derived_layout["binary"] != stored_layout["binary"] or derived_layout["perf_data"] != stored_layout["perf_data"]:
        raise AssertionError("fresh original layout custody disagrees with sealed layout")

    # First replay the unmodified retained raw file through the same final-v2
    # shim.  The historical corrected script is bound below only as evidence;
    # it was produced by the old failing decoder and is not a comparison oracle.
    original_argv = [
        "/usr/bin/perf", "script", "--show-lost-events", "--no-inline",
        "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso", "-i", str(RAW),
    ]
    original_log_path = OUT / "original.final-v2.dwfl.jsonl"
    original_script = OUT / "original.final-v2.corrected.script.txt"
    original_stderr = OUT / "original.final-v2.corrected.stderr.txt"
    original_env = os.environ.copy()
    original_env.update({
        "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "LITCHI_0973_SHIM_MODE": "corrected",
        "LITCHI_0973_SHIM_LOG": str(original_log_path),
        "LITCHI_0973_TARGET_PATH": TARGET_PATH,
        "LITCHI_0973_TARGET_NAME": BINARY.name,
        "LITCHI_0973_CORRECTED_BASE": derived_layout["corrected_base"],
        "LITCHI_0973_PERF_EXPECTED_START": derived_layout["perf_expected_start"],
        "LITCHI_0976_TARGET_BUILD_ID": TARGET_BUILD_ID,
        "LITCHI_0976_TARGET_SHA256": TARGET_SHA256,
        "LITCHI_0976_EXPECTED_DELTA": "0x1000",
        "LD_PRELOAD": str(SHIM),
    })
    original_run = run_to_files(original_argv, original_env, original_script, original_stderr)
    if original_run["returncode"] != 0 or not original_log_path.is_file():
        raise RuntimeError("original final-v2 decoder replay failed")

    translated, mutation = mutate_perf_data(raw, derived_layout)
    translated_path = OUT / "child-01.translated.perf.data"
    translated_path.write_bytes(translated)
    mutation_path = OUT / "mutation-map.json"
    mutation_path.write_text(json.dumps(mutation, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    mmap_dump = OUT / "translated.perf-mmap-dump.txt"
    mmap_err = OUT / "translated.perf-mmap-dump.stderr"
    mmap_run = run_to_files(
        ["/usr/bin/perf", "script", "-D", "-i", str(translated_path)],
        {"LC_ALL": "C", "LANG": "C", "TZ": "UTC"}, mmap_dump, mmap_err,
    )
    if mmap_run["returncode"] != 0:
        raise RuntimeError("translated MMAP2 dump failed")

    layout_path = OUT / "translated.target-layout.json"
    translated_layout = protocol.derive_target_layout(translated_path, BINARY, expected_pid=EXPECTED_PID)
    layout_path.write_text(json.dumps(translated_layout, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    expected_base = int(derived_layout["corrected_base"], 16) + DELTA
    expected_perf_base = int(derived_layout["perf_expected_start"], 16) + DELTA
    if translated_layout["corrected_base"] != f"0x{expected_base:x}" or translated_layout["perf_expected_start"] != f"0x{expected_perf_base:x}":
        raise AssertionError("translated layout did not apply the declared delta")

    env = os.environ.copy()
    env.update({
        "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "LITCHI_0973_SHIM_MODE": "corrected",
        "LITCHI_0973_SHIM_LOG": str(OUT / "translated.dwfl.jsonl"),
        "LITCHI_0973_TARGET_PATH": TARGET_PATH,
        "LITCHI_0973_TARGET_NAME": BINARY.name,
        "LITCHI_0973_CORRECTED_BASE": translated_layout["corrected_base"],
        "LITCHI_0973_PERF_EXPECTED_START": translated_layout["perf_expected_start"],
        "LITCHI_0976_TARGET_BUILD_ID": TARGET_BUILD_ID,
        "LITCHI_0976_TARGET_SHA256": TARGET_SHA256,
        # Shim admission uses perf_expected_start-corrected_base, which stays
        # 0x1000 after the synthetic ASLR translation.
        "LITCHI_0976_EXPECTED_DELTA": "0x1000",
        "LD_PRELOAD": str(SHIM),
    })
    translated_script = OUT / "translated.corrected.script.txt"
    translated_stderr = OUT / "translated.corrected.stderr.txt"
    translated_argv = [
        "/usr/bin/perf", "script", "--show-lost-events", "--no-inline",
        "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso", "-i", str(translated_path),
    ]
    translated_run = run_to_files(translated_argv, env, translated_script, translated_stderr)
    log_path = OUT / "translated.dwfl.jsonl"
    if translated_run["returncode"] != 0 or not log_path.is_file():
        raise RuntimeError("translated corrected decoder replay failed")

    compare = compare_sequences(protocol, original_script.read_text(encoding="utf-8"), translated_script.read_text(encoding="utf-8"))
    if compare["original"] != compare["translated"]:
        raise AssertionError("decoded summary identity failed")
    recovery = {}
    for label, summary in (("original", compare["original"]), ("translated", compare["translated"])):
        # These are decoder-recovery conditions for this retained child, not
        # universal attribution gates.  Unresolved frames and explicit stack
        # truncation are recorded below and intentionally remain descriptive.
        conditions = {
            "sample_count_6940": summary["sample_count"] == 6940,
            "total_period_6940000000": summary["total_period"] == 6_940_000_000,
            "empty_callchain_count_zero": summary["empty_callchain_count"] == 0,
            "lost_events_zero": summary["lost_events"] == 0,
            "marker_sample_count_at_least_100": summary["marker_sample_count"] >= 100,
            "marker_period_positive": summary["marker_period"] > 0,
        }
        if not all(conditions.values()):
            raise AssertionError(f"{label} decoder recovery conditions failed: {conditions}")
        recovery[label] = {"conditions": conditions,
                           "unresolved_frames": summary["unresolved_frames"],
                           "explicit_stack_truncated_count": summary["explicit_stack_truncated_count"]}
    log_rows = [json.loads(line) for line in log_path.read_text(encoding="utf-8").splitlines() if line]
    original_log_rows = [json.loads(line) for line in original_log_path.read_text(encoding="utf-8").splitlines() if line]
    if not log_rows or not original_log_rows:
        raise AssertionError("one or both final-v2 shim replays produced no diagnostics")
    target_rows = [row for row in log_rows
                   if row.get("schema") == "litchi.performance.0976.dwfl_report.final"
                   and row.get("path_match") is True
                   and row.get("name_match") is True]
    original_target_rows = [row for row in original_log_rows
                            if row.get("schema") == "litchi.performance.0976.dwfl_report.final"
                            and row.get("path_match") is True
                            and row.get("name_match") is True]
    target_checks = {}
    for label, all_rows, rows in (("original", original_log_rows, original_target_rows), ("translated", log_rows, target_rows)):
        # Count only exact path+name target rows. Other DSO reports may be
        # legitimate and are not treated as target fallback failures.
        fallback_rows = [row for row in rows if row.get("admission") is not True]
        if len(rows) != 3:
            raise AssertionError(f"{label} final-v2 target reports: expected one initial and two target reuse reports, got {len(rows)}")
        if fallback_rows:
            raise AssertionError(f"{label} exact target fallback reports: {len(fallback_rows)}")
        if sum(row.get("initial_rewrite") is True for row in rows) != 1:
            raise AssertionError(f"{label} target reports did not contain one initial rewrite")
        if sum(row.get("reused_existing") is True for row in rows) != 2:
            raise AssertionError(f"{label} target reports did not contain two cache reuses")
        if any(row.get("identity_valid") is not True for row in rows):
            raise AssertionError(f"{label} target report lacked exact identity validation")
        target_checks[label] = {"exact_target_rows": len(rows), "fallback_rows": 0,
                                "initial_rewrite_rows": 1, "reuse_rows": 2,
                                "identity_valid_rows": len(rows),
                                "all_final_report_rows": len(all_rows)}
    input_hashes_after = {str(path): binding(path) for path in input_paths}
    if input_hashes_after != input_hashes_before:
        raise RuntimeError("frozen input changed during ASLR positive control")
    translated_script_binding = binding(translated_script)
    receipt = {
        "schema": "litchi.performance.0976.aslr_translation.final_v2.v2",
        "status": "pass",
        "created_utc": started,
        "finished_utc": utc(),
        "purpose": "offline positive ASLR translation control for retained 0975 child-01, comparing original and translated replays through the same final-v2 shim; no workload or native build",
        "inputs": {
            "raw_perf_data": raw_binding,
            "binary": binary_binding,
            "shim": shim_binding,
            "retained_legacy_corrected_script": legacy_script_binding,
            "original_layout": original_layout_binding,
            "original_mmap_dump": original_mmap_binding,
            "protocol_helper": binding(PROTOCOL_PATH),
            "test_script": binding(SCRIPT_PATH),
            "shim_source": binding(shim_source_path),
            "shim_version_script": binding(shim_map_path),
        },
        "input_integrity": {
            "before": input_hashes_before,
            "after": input_hashes_after,
            "unchanged": True,
        },
        "translation": {
            "delta": f"0x{DELTA:x}",
            "delta_page_aligned": DELTA % PAGE == 0,
            "synthetic_perf_data": binding(translated_path),
            "mutation_map": binding(mutation_path),
            "mutation": {key: value for key, value in mutation.items() if key != "changes"},
            "changed_fields_recorded": mutation["changed_field_count"],
            "original_immutable": sha256(RAW) == raw_binding["sha256"],
        },
        "layout": {
            "original": derived_layout,
            "translated": {"path": str(layout_path), "binding": binding(layout_path), "layout": translated_layout},
            "derived_from": "target PERF_RECORD_MMAP2 rows plus retained ELF64 PT_LOAD layout helper",
            "base_delta": translated_layout["corrected_base"] == f"0x{expected_base:x}",
            "perf_comparator_delta": translated_layout["perf_expected_start"] == f"0x{expected_perf_base:x}",
        },
        "commands": {
            "decoder_dependency": {
                "active_v2_shim_build_argv": active_shim_build,
                "protocol_helper": binding(PROTOCOL_PATH),
                "shim": binding(SHIM),
                "shim_source": binding(shim_source_path),
                "shim_version_script": binding(shim_map_path),
                "target_sha256_env": TARGET_SHA256,
                "expected_delta_env": "0x1000",
                "native_workload_started": False,
                "native_build_started": False,
            },
            "original_layout": ["protocol_common_0973_v2.derive_target_layout", str(RAW), str(BINARY)],
            "original_final_v2_replay": {"argv": original_argv, "environment": {key: original_env[key] for key in ("LC_ALL", "LANG", "TZ", "LITCHI_0973_SHIM_MODE", "LITCHI_0973_TARGET_PATH", "LITCHI_0973_TARGET_NAME", "LITCHI_0973_CORRECTED_BASE", "LITCHI_0973_PERF_EXPECTED_START", "LITCHI_0976_TARGET_BUILD_ID", "LITCHI_0976_TARGET_SHA256", "LITCHI_0976_EXPECTED_DELTA", "LD_PRELOAD")}, "result": original_run},
            "translated_mmap_dump": mmap_run,
            "translated_layout": ["protocol_common_0973_v2.derive_target_layout", str(translated_path), str(BINARY)],
            "translated_corrected_replay": {"argv": translated_argv, "environment": {key: env[key] for key in ("LC_ALL", "LANG", "TZ", "LITCHI_0973_SHIM_MODE", "LITCHI_0973_TARGET_PATH", "LITCHI_0973_TARGET_NAME", "LITCHI_0973_CORRECTED_BASE", "LITCHI_0973_PERF_EXPECTED_START", "LITCHI_0976_TARGET_BUILD_ID", "LITCHI_0976_TARGET_SHA256", "LITCHI_0976_EXPECTED_DELTA", "LD_PRELOAD")}, "result": translated_run},
        },
        "replay": {
            "translated_mmap_dump": binding(mmap_dump),
            "translated_mmap_stderr": binding(mmap_err),
            "original_final_v2_script": binding(original_script),
            "original_final_v2_stderr": binding(original_stderr),
            "original_final_v2_shim_log": binding(original_log_path),
            "original_final_v2_shim_log_rows": len(original_log_rows),
            "translated_script": translated_script_binding,
            "translated_stderr": binding(translated_stderr),
            "translated_shim_log": binding(log_path),
            "translated_shim_log_rows": len(log_rows),
            "comparison": compare,
            "recovery_conditions": recovery,
            "target_report_checks": target_checks,
            "canonical_resolved_sequence_invariant": True,
            "address_invariance_rule": "target DSO frame addresses shift by delta; DSO/symbol/offset/resolution sequence is unchanged",
        },
        "scope_limits": [
            "The synthetic copy changes perf sample metadata only; it does not execute a relocated binary.",
            "The retained historical corrected script is not an oracle because it contains the earlier decoder's empty-chain failure; both comparison sides are fresh final-v2 replays.",
            "The control validates this retained DWFL/MMAP2 decoder path for one sealed child, not every future perf or kernel format.",
            "The result is decoder evidence only; it provides no timing, speedup, or production attribution claim.",
        ],
        "native_workload_started": False,
        "build_started": False,
        "external_helper_build": True,
        "source_changed": False,
    }
    if RECEIPT_PATH.exists():
        raise FileExistsError(f"refusing to overwrite receipt {RECEIPT_PATH}")
    RECEIPT_PATH.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({
        "status": "pass",
        "receipt": str(RECEIPT_PATH),
        "receipt_sha256": sha256(RECEIPT_PATH),
        "synthetic_perf_data": str(translated_path),
        "canonical_sequence_sha256": compare["canonical_sequence_sha256"],
        "mutation": {key: value for key, value in mutation.items() if key.endswith("shifted") or key.endswith("count")},
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
