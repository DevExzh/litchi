#!/usr/bin/env python3
"""Bounded offline PERF_RECORD_SAMPLE/STACK_USER admission scan for 0973.

The scan consumes the already-retained ``perf script -D`` text for one child.
It does not unwind STACK_USER words and makes no ancestry claim.  It records
the raw sample layout and rejects a child when a sample violates the declared
period, PID/TID/CPU, or ``0 <= dyn_size <= captured <= stack_cap`` contract.
"""
from __future__ import annotations

import collections
import argparse
import json
import re
import struct
import tempfile
from pathlib import Path
from typing import Any

import protocol_common_0973_v2 as p

RAW_EVENT = re.compile(r"^\.\s+\.\.\. raw event: size (\d+) bytes$")
HEX_LINE = re.compile(r"^\.\s+([0-9a-fA-F]+):\s+((?:[0-9a-fA-F]{2}\s+)+)")
SAMPLE = re.compile(
    r"PERF_RECORD_SAMPLE\(IP,\s*0x([0-9a-fA-F]+)\):\s+"
    r"(\d+)/(\d+):\s+0x([0-9a-fA-F]+)\s+period:\s+(\d+)"
)
STACK_SIZE_OFFSET = 224
STACK_BYTES_OFFSET = 232
MIN_HEADER = STACK_BYTES_OFFSET + 8


def _u32(data: bytes | bytearray, offset: int) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def _u64(data: bytes | bytearray, offset: int) -> int:
    return struct.unpack_from("<Q", data, offset)[0]


def _count(counter: collections.Counter[int]) -> dict[str, int]:
    return {str(key): value for key, value in sorted(counter.items())}


def scan_dump(
    dump_path: Path,
    *,
    perf_data: Path,
    binary: Path,
    expected_pid: int,
    expected_cpu: int = p.CPU,
    expected_period: int = p.PERIOD,
    stack_cap: int = p.STACK_CAP,
    layout: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if not dump_path.is_file() or not perf_data.is_file() or not binary.is_file():
        raise p.ProtocolError("raw STACK_USER scan input is missing")
    raw: bytearray | None = None
    raw_consumed = False
    raw_event_count = 0
    sample_line_count = 0
    sample_count = 0
    unconsumed_raw_count = 0
    record_types: collections.Counter[int] = collections.Counter()
    record_sizes: collections.Counter[int] = collections.Counter()
    stack_sizes: collections.Counter[int] = collections.Counter()
    dyn_sizes: collections.Counter[int] = collections.Counter()
    pid_tid_cpu: collections.Counter[str] = collections.Counter()
    missing: collections.Counter[str] = collections.Counter()
    violations: list[dict[str, Any]] = []
    period_valid = True
    pid_tid_cpu_valid = True
    bounds_valid = True
    pf_x_sample_count = 0
    pf_x_interval_count = 0
    pf_x_intervals: list[tuple[int, int]] = []
    if layout is not None:
        try:
            base = int(layout["corrected_base"], 0)
            elf = p._parse_elf_loads(binary)
            pf_x_intervals = [
                (base + row["vaddr"], base + row["vaddr"] + row["memsz"])
                for row in elf["pt_loads"] if row["flags"] & 0x1
            ]
            pf_x_interval_count = len(pf_x_intervals)
        except (KeyError, TypeError, ValueError) as exc:
            raise p.ProtocolError(f"cannot derive PF_X intervals for raw scan: {exc}") from exc

    def violation(reason: str, **fields: Any) -> None:
        if len(violations) < 64:
            violations.append({"reason": reason, **fields})

    def process_sample(line: str, current: bytearray | None) -> None:
        nonlocal sample_count, period_valid, pid_tid_cpu_valid, bounds_valid, pf_x_sample_count
        match = SAMPLE.search(line)
        if match is None:
            return
        event_misc = int(match.group(1), 16)
        line_pid = int(match.group(2))
        line_tid = int(match.group(3))
        line_ip = int(match.group(4), 16)
        line_period = int(match.group(5))
        if line_period != expected_period:
            period_valid = False
            violation("sample_line_period", value=line_period, expected=expected_period)
        if line_pid != expected_pid or line_tid != expected_pid:
            pid_tid_cpu_valid = False
            violation("sample_line_pid_tid", pid=line_pid, tid=line_tid, expected=expected_pid)
        if current is None:
            missing["raw_record_for_sample"] += 1
            violation("missing_raw_record_for_sample", line=sample_line_count)
            return
        sample_count += 1
        if len(current) < 8:
            missing["sample_header"] += 1
            violation("sample_record_too_short", bytes=len(current))
            return
        record_type, record_misc, record_size = struct.unpack_from("<IHH", current, 0)
        record_types[record_type] += 1
        record_sizes[len(current)] += 1
        if record_size != len(current):
            violation("record_size_mismatch", record_size=record_size, bytes=len(current))
        if event_misc != record_misc:
            violation("record_misc_mismatch", event_misc=event_misc, record_misc=record_misc)
        if len(current) < MIN_HEADER:
            missing["sample_fields"] += 1
            violation("sample_fields_truncated", bytes=len(current))
            return
        raw_ip = _u64(current, 8)
        raw_pid = _u32(current, 16)
        raw_tid = _u32(current, 20)
        raw_cpu = _u32(current, 32)
        raw_period = _u64(current, 40)
        if raw_ip != line_ip or raw_period != line_period:
            violation("sample_header_payload_mismatch", line_ip=hex(line_ip), raw_ip=hex(raw_ip), line_period=line_period, raw_period=raw_period)
        if any(start <= raw_ip < end for start, end in pf_x_intervals):
            pf_x_sample_count += 1
        if raw_pid != expected_pid or raw_tid != expected_pid or raw_cpu != expected_cpu:
            pid_tid_cpu_valid = False
            violation("sample_payload_pid_tid_cpu", pid=raw_pid, tid=raw_tid, cpu=raw_cpu, expected_pid=expected_pid, expected_cpu=expected_cpu)
        pid_tid_cpu[f"{raw_pid}/{raw_tid}@{raw_cpu}"] += 1
        if raw_period != expected_period:
            period_valid = False
            violation("sample_payload_period", value=raw_period, expected=expected_period)
        captured = _u64(current, STACK_SIZE_OFFSET)
        stack_sizes[captured] += 1
        stack_end = STACK_BYTES_OFFSET + captured
        if stack_end + 8 > len(current):
            missing["dyn_size"] += 1
            violation("dyn_size_outside_record", captured=captured, record_bytes=len(current))
            return
        dyn = _u64(current, stack_end)
        dyn_sizes[dyn] += 1
        if not (0 <= dyn <= captured <= stack_cap):
            bounds_valid = False
            violation("stack_bound", dyn_size=dyn, captured_stack_size=captured, stack_cap=stack_cap)

    with dump_path.open("r", encoding="utf-8", errors="strict") as stream:
        for line in stream:
            line = line.rstrip("\n")
            match = RAW_EVENT.match(line)
            if match:
                if raw is not None and not raw_consumed:
                    unconsumed_raw_count += 1
                raw = bytearray(int(match.group(1)))
                raw_consumed = False
                raw_event_count += 1
                continue
            match = HEX_LINE.match(line)
            if match and raw is not None:
                offset = int(match.group(1), 16)
                payload = bytes.fromhex(match.group(2))
                if offset < 0 or offset + len(payload) > len(raw):
                    violation("hex_payload_outside_record", offset=offset, payload_bytes=len(payload), record_bytes=len(raw))
                else:
                    raw[offset : offset + len(payload)] = payload
                continue
            if SAMPLE.search(line):
                sample_line_count += 1
                process_sample(line, raw)
                raw = None
                raw_consumed = True
    if raw is not None and not raw_consumed:
        unconsumed_raw_count += 1
    if sample_count != sample_line_count:
        violation("sample_line_record_count", sample_lines=sample_line_count, sample_records=sample_count)
    if sample_count == 0:
        violation("no_sample_records")
    return {
        "schema": p.RAW_STACK_SCHEMA,
        "status": "complete",
        "binary": p.file_binding(binary),
        "perf_data": p.file_binding(perf_data),
        "dump": p.file_binding(dump_path),
        "command_scope": "offline perf script -D text scan; no STACK_USER unwind or ancestry claim",
        "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
        "sample_count": sample_count,
        "sample_line_count": sample_line_count,
        "raw_event_count": raw_event_count,
        "unconsumed_raw_count": unconsumed_raw_count,
        "period": expected_period,
        "stack_cap": stack_cap,
        "bounds_rule": "0 <= dyn_size <= captured_stack_size <= 32768",
        "period_valid": period_valid,
        "pid_tid_cpu_valid": pid_tid_cpu_valid,
        "bounds_valid": bounds_valid,
        "missing": dict(sorted(missing.items())),
        "violations": violations,
        "record_type_counts": _count(record_types),
        "record_size_counts": _count(record_sizes),
        "stack_size_counts": _count(stack_sizes),
        "dyn_size_counts": _count(dyn_sizes),
        "pid_tid_cpu_counts": dict(sorted(pid_tid_cpu.items())),
        "pf_x": {
            "interval_count": pf_x_interval_count,
            "intervals": [[hex(start), hex(end)] for start, end in pf_x_intervals],
            "sample_ip_in_interval_count": pf_x_sample_count,
            "scope": "raw sample IP only; no STACK_USER ancestry claim",
        },
        "scope": {
            "all_perf_record_sample_records": True,
            "misc_filter": None,
            "stack_bytes_are_potential_values_only": True,
            "no_live_ancestry_claim": True,
        },
    }


def synthetic_boundary_summary(*, binary: Path, perf_data: Path, dump_path: Path) -> dict[str, Any]:
    """Explicit test-only boundary receipt; never admissible for real capture."""
    return {
        "schema": p.RAW_STACK_SCHEMA,
        "status": "complete",
        "offline_boundary_mock": True,
        "boundary_scope": "offline perf-boundary mock; no native sample bytes",
        "binary": p.file_binding(binary),
        "perf_data": p.file_binding(perf_data),
        "dump": p.file_binding(dump_path),
        "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
        "sample_count": 100,
        "sample_line_count": 100,
        "raw_event_count": 100,
        "unconsumed_raw_count": 0,
        "period": p.PERIOD,
        "stack_cap": p.STACK_CAP,
        "bounds_rule": "0 <= dyn_size <= captured_stack_size <= 32768",
        "period_valid": True,
        "pid_tid_cpu_valid": True,
        "bounds_valid": True,
        "missing": {},
        "violations": [],
        "record_type_counts": {"9": 100},
        "record_size_counts": {"33000": 100},
        "stack_size_counts": {"32768": 100},
        "dyn_size_counts": {"32768": 100},
        "pid_tid_cpu_counts": {"synthetic": 100},
        "scope": {"all_perf_record_sample_records": True, "misc_filter": None, "stack_bytes_are_potential_values_only": True, "no_live_ancestry_claim": True},
    }


def self_test() -> int:
    root = Path("/home/zhuhe/litchi-performance-0973/replay-dwfl-v2")
    raw = Path("/home/zhuhe/litchi-performance-0971/capture/instruction-profile-v2/child-01/child-01.perf.data")
    layout = p.load_json(root / "child-01.target-layout.json")
    result = scan_dump(root / "child-01.perf-mmap-dump.txt", perf_data=raw, binary=Path(p.BINARY_PATH), expected_pid=3683783, expected_cpu=p.CPU, expected_period=5_000_000, stack_cap=p.STACK_CAP, layout=layout)
    if result["sample_count"] != 1387 or result["violations"] or not result["bounds_valid"] or not result["period_valid"] or not result["pid_tid_cpu_valid"]:
        raise p.ProtocolError("raw STACK_USER self-test failed")
    with tempfile.TemporaryDirectory(prefix="litchi-0973-raw-test-") as temp:
        bad = bytearray(240)
        struct.pack_into("<IHH", bad, 0, 9, 2, len(bad))
        struct.pack_into("<Q", bad, 8, 0x1000)
        struct.pack_into("<II", bad, 16, 367, 367)
        struct.pack_into("<I", bad, 32, p.CPU)
        struct.pack_into("<Q", bad, 40, p.PERIOD)
        struct.pack_into("<Q", bad, STACK_SIZE_OFFSET, 0)
        struct.pack_into("<Q", bad, STACK_BYTES_OFFSET, 1)
        lines = [". ... raw event: size 240 bytes"]
        for offset in range(0, len(bad), 16):
            lines.append(f". {offset:04x}: " + " ".join(f"{byte:02x}" for byte in bad[offset:offset + 16]) + " ")
        lines.append(f". PERF_RECORD_SAMPLE(IP, 0x2): 367/367: 0x1000 period: {p.PERIOD}")
        dump = Path(temp) / "bad.dump"
        data = Path(temp) / "bad.perf.data"
        dump.write_text("\n".join(lines) + "\n")
        data.write_bytes(b"synthetic perf boundary")
        malformed = scan_dump(dump, perf_data=data, binary=Path(p.BINARY_PATH), expected_pid=367, expected_cpu=p.CPU, expected_period=p.PERIOD, stack_cap=p.STACK_CAP)
        if malformed["bounds_valid"] is not False or not malformed["violations"]:
            raise p.ProtocolError("raw STACK_USER malformed-bound self-test was accepted")
    print(json.dumps({"schema": "litchi.performance.0973.raw_stack_self_test.v2", "status": "pass", "historical_period": 5_000_000, "sample_count": result["sample_count"], "bounds_valid": result["bounds_valid"], "malformed_bound_rejected": True, "workload_started": False}, sort_keys=True))
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if not args.self_test:
        raise SystemExit("use --self-test; capture-bound scan is called by the runner")
    raise SystemExit(self_test())
