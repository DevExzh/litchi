#!/usr/bin/env python3
"""Offline raw STACK_USER scan with ELF PT_LOAD executable intervals.

This reads only the retained 0971 perf.data and its retained ELF.  It never
executes the benchmark, builds code, or changes perf/debug caches.  It scans
the raw sample records emitted by ``perf script -D``; words found in the raw
STACK_USER bytes are potential address values only and are not unwound frames
or proof of live ancestry.
"""
from __future__ import annotations

import collections
import hashlib
import json
import re
import struct
import subprocess
from pathlib import Path
from typing import Any

DATA = Path(
    "/home/zhuhe/litchi-performance-0971/capture/instruction-profile-v2/child-01/child-01.perf.data"
)
BINARY = Path(
    "/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile"
)
OUT = Path(
    "/home/zhuhe/litchi-performance-0972/analysis/corrected-raw-stack-scan-0972-v2.json"
)
PERF = Path("/usr/bin/perf")
NM = Path("/usr/bin/nm")
TARGET_PATH = str(BINARY)
TARGET_BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
PAGE = 0x1000
PF_X = 0x1
PT_LOAD = 1
EXPECTED_SAMPLE_STACK_BYTES = 16384
EXPECTED_SAMPLE_REG_MASK = 0xFF0FFF
STACK_SIZE_FIELD_OFFSET = 224
STACK_BYTES_OFFSET = 232

RAW_EVENT = re.compile(r"^\.\s+\.\.\. raw event: size (\d+) bytes$")
HEX_LINE = re.compile(r"^\.\s+([0-9a-fA-F]+):\s+((?:[0-9a-fA-F]{2}\s+)+)")
SAMPLE = re.compile(
    r"PERF_RECORD_SAMPLE\(IP,\s*0x([0-9a-fA-F]+)\):\s+"
    r"\d+/\d+:\s+0x([0-9a-fA-F]+)\s+period:\s+(\d+)"
)
MMAP2 = re.compile(
    r"PERF_RECORD_MMAP2\s+\d+/\d+:\s+"
    r"\[(0x[0-9a-fA-F]+)\((0x[0-9a-fA-F]+)\)\s+@\s+"
    r"(0x[0-9a-fA-F]+|[0-9]+)\s+.*?\]:\s+\S+\s+(.*)$"
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def align_down(value: int, alignment: int = PAGE) -> int:
    return value & ~(alignment - 1)


def align_up(value: int, alignment: int = PAGE) -> int:
    return (value + alignment - 1) & ~(alignment - 1)


def parse_elf_loads(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    data = path.read_bytes()
    if data[:4] != b"\x7fELF" or data[4] != 2 or data[5] != 1:
        raise RuntimeError("expected little-endian ELF64")
    # ELF64 header fields: e_phoff at 32, e_phentsize at 54, e_phnum at 56.
    phoff = struct.unpack_from("<Q", data, 32)[0]
    phentsize = struct.unpack_from("<H", data, 54)[0]
    phnum = struct.unpack_from("<H", data, 56)[0]
    loads: list[dict[str, Any]] = []
    for index in range(phnum):
        offset = phoff + index * phentsize
        p_type, p_flags = struct.unpack_from("<II", data, offset)
        p_offset, p_vaddr, _p_paddr, p_filesz, p_memsz, p_align = struct.unpack_from(
            "<QQQQQQ", data, offset + 8
        )
        if p_type != PT_LOAD:
            continue
        item: dict[str, Any] = {
            "index": index,
            "flags": hex(p_flags),
            "p_offset": hex(p_offset),
            "p_vaddr": hex(p_vaddr),
            "p_filesz": hex(p_filesz),
            "p_memsz": hex(p_memsz),
            "p_align": hex(p_align),
        }
        if p_flags & PF_X:
            item["executable"] = True
            item["p_vaddr_end"] = hex(p_vaddr + p_memsz)
        else:
            item["executable"] = False
        loads.append(item)
    ident = {
        "class": "ELF64",
        "data": "little_endian",
        "type": hex(struct.unpack_from("<H", data, 16)[0]),
        "machine": hex(struct.unpack_from("<H", data, 18)[0]),
        "entry": hex(struct.unpack_from("<Q", data, 24)[0]),
        "program_header_offset": phoff,
        "program_header_entry_size": phentsize,
        "program_header_count": phnum,
    }
    return ident, loads


def read_symbols(path: Path) -> dict[str, dict[str, int]]:
    proc = subprocess.run(
        [str(NM), "-anS", "--defined-only", str(path)],
        check=True,
        capture_output=True,
        text=True,
    )
    wanted = {
        "_ZN36doc_retention_probe_0942_cfb_profile22profiled_measured_case17hf64a70791a9849aaE": "profiled_measured_case",
        "_ZN36doc_retention_probe_0942_cfb_profile11measure_one17h1e1e7579e7611666E": "measure_one",
        "_ZN36doc_retention_probe_0942_cfb_profile8run_case17h0894cea56c023e63E": "run_case",
        "_ZN10litchi_doc4sprm11parse_sprms17h26418f62a3c84e01E": "parse_sprms",
    }
    result: dict[str, dict[str, int]] = {}
    for line in proc.stdout.splitlines():
        fields = line.split(maxsplit=3)
        if len(fields) != 4:
            continue
        value_text, size_text, _kind, name = fields
        if name not in wanted:
            continue
        result[wanted[name]] = {"st_value": int(value_text, 16), "size": int(size_text, 16)}
    if set(result) != set(wanted.values()):
        raise RuntimeError(f"missing expected symbols: {sorted(set(wanted.values()) - set(result))}")
    return result


def in_intervals(value: int, intervals: list[tuple[int, int]]) -> bool:
    return any(start <= value < end for start, end in intervals)


def u64(data: bytes | bytearray, offset: int) -> int:
    return struct.unpack_from("<Q", data, offset)[0]


def u32(data: bytes | bytearray, offset: int) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def scan() -> dict[str, Any]:
    if not DATA.is_file() or not BINARY.is_file():
        raise FileNotFoundError("sealed perf.data or ELF input is missing")
    elf_identity, loads = parse_elf_loads(BINARY)
    symbols = read_symbols(BINARY)
    proc = subprocess.Popen(
        [str(PERF), "script", "-D", "-i", str(DATA)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
    )
    assert proc.stdout is not None and proc.stderr is not None
    raw: bytearray | None = None
    raw_event_count = 0
    sample_count = 0
    sample_line_count = 0
    sample_misc = collections.Counter()
    record_types = collections.Counter()
    record_sizes = collections.Counter()
    header_misc = collections.Counter()
    target_maps: list[dict[str, Any]] = []
    target_samples: list[dict[str, Any]] = []
    stack_size_counts = collections.Counter()
    dyn_size_counts = collections.Counter()
    callchain_counts = collections.Counter()
    abi_counts = collections.Counter()
    stack_layout_failures: list[dict[str, Any]] = []
    event_mismatches: list[dict[str, Any]] = []

    # The target's offset-zero MMAP2 record is the load-bias binding.  It is
    # emitted before the sample records in this retained perf.data stream.
    load_bias: int | None = None
    for line in proc.stdout:
        line = line.rstrip("\n")
        mmap = MMAP2.search(line)
        if mmap and TARGET_PATH in mmap.group(4):
            map_start = int(mmap.group(1), 16)
            map_length = int(mmap.group(2), 16)
            pgoff_text = mmap.group(3)
            pgoff = int(pgoff_text, 16) if pgoff_text.startswith("0x") else int(pgoff_text)
            target_maps.append(
                {
                    "start": hex(map_start),
                    "length": hex(map_length),
                    "pgoff": hex(pgoff),
                    "line": line,
                }
            )
            if pgoff == 0 and load_bias is None:
                load_bias = map_start
        match = RAW_EVENT.match(line)
        if match:
            raw = bytearray(int(match.group(1)))
            raw_event_count += 1
            record_sizes[len(raw)] += 1
            continue
        match = HEX_LINE.match(line)
        if match and raw is not None:
            offset = int(match.group(1), 16)
            payload = bytes.fromhex(match.group(2))
            raw[offset : offset + len(payload)] = payload
            continue
        match = SAMPLE.search(line)
        if not match or raw is None:
            continue
        sample_line_count += 1
        event_misc = int(match.group(1), 16)
        event_ip = int(match.group(2), 16)
        event_period = int(match.group(3))
        if len(raw) < 8:
            stack_layout_failures.append({"reason": "record_too_short_for_header", "bytes": len(raw)})
            raw = None
            continue
        record_type, record_misc, record_size = struct.unpack_from("<IHH", raw, 0)
        record_types[record_type] += 1
        header_misc[record_misc] += 1
        sample_misc[event_misc] += 1
        if event_misc != record_misc:
            event_mismatches.append({"event_misc": event_misc, "header_misc": record_misc})
        if record_size != len(raw):
            event_mismatches.append({"record_size": record_size, "raw_bytes": len(raw)})
        sample_count += 1
        if len(raw) < STACK_BYTES_OFFSET + 8:
            stack_layout_failures.append({"reason": "record_too_short", "bytes": len(raw)})
            raw = None
            continue
        ip = u64(raw, 8)
        tid = [u32(raw, 16), u32(raw, 20)]
        timestamp = u64(raw, 24)
        cpu = u32(raw, 32)
        period = u64(raw, 40)
        callchain_nr = u64(raw, 48)
        abi = u64(raw, 56)
        stack_size = u64(raw, STACK_SIZE_FIELD_OFFSET)
        stack_start = STACK_BYTES_OFFSET
        stack_end = stack_start + stack_size
        if stack_end + 8 > len(raw):
            stack_layout_failures.append(
                {"reason": "stack_out_of_record", "bytes": len(raw), "stack_size": stack_size}
            )
            raw = None
            continue
        dyn_size = u64(raw, stack_end)
        stack_size_counts[stack_size] += 1
        dyn_size_counts[dyn_size] += 1
        callchain_counts[callchain_nr] += 1
        abi_counts[abi] += 1
        if event_ip != ip or event_period != period:
            event_mismatches.append(
                {
                    "event_ip": hex(event_ip),
                    "header_ip": hex(ip),
                    "event_period": event_period,
                    "header_period": period,
                }
            )
        if (
            record_size != 16624
            or stack_size != EXPECTED_SAMPLE_STACK_BYTES
            or dyn_size != EXPECTED_SAMPLE_STACK_BYTES
            or callchain_nr != 0
        ):
            stack_layout_failures.append(
                {
                    "reason": "unexpected_layout",
                    "record_bytes": len(raw),
                    "record_size": record_size,
                    "stack_size": stack_size,
                    "dyn_size": dyn_size,
                    "callchain_nr": callchain_nr,
                }
            )
        if load_bias is not None:
            executable_loads = [
                load
                for load in loads
                if load["executable"]
            ]
            intervals = [
                (
                    load_bias + int(load["p_vaddr"], 16),
                    load_bias + int(load["p_vaddr_end"], 16),
                )
                for load in executable_loads
            ]
            if in_intervals(ip, intervals):
                stack = bytes(raw[stack_start:stack_end])
                words: list[dict[str, Any]] = []
                for relative in range(0, len(stack) - 7, 8):
                    value = u64(stack, relative)
                    if in_intervals(value, intervals):
                        words.append(
                            {
                                "stack_relative_byte": relative,
                                "record_byte": stack_start + relative,
                                "value": hex(value),
                            }
                        )
                marker_start = load_bias + symbols["profiled_measured_case"]["st_value"]
                marker_end = marker_start + symbols["profiled_measured_case"]["size"]
                marker_return = marker_start + 0x1B
                measure_start = load_bias + symbols["measure_one"]["st_value"]
                measure_return = measure_start + 0xF3
                run_start = load_bias + symbols["run_case"]["st_value"]
                run_end = run_start + symbols["run_case"]["size"]
                marker_words = [
                    word
                    for word in words
                    if marker_start <= int(word["value"], 16) < marker_end
                ]
                marker_returns = [
                    word for word in words if int(word["value"], 16) == marker_return
                ]
                measure_returns = [
                    word for word in words if int(word["value"], 16) == measure_return
                ]
                run_words = [
                    word
                    for word in words
                    if run_start <= int(word["value"], 16) < run_end
                ]
                target_samples.append(
                    {
                        "timestamp_ns": timestamp,
                        "tid": tid,
                        "cpu": cpu,
                        "misc": hex(record_misc),
                        "ip": hex(ip),
                        "period": period,
                        "record_bytes": len(raw),
                        "stack_size": stack_size,
                        "dyn_size": dyn_size,
                        "target_text_word_count": len(words),
                        "marker_range_words": marker_words,
                        "marker_return_words": marker_returns,
                        "measure_one_return_words": measure_returns,
                        "run_case_range_words": run_words,
                    }
                )
        raw = None
    stderr = proc.stderr.read()
    returncode = proc.wait()
    if load_bias is None:
        raise RuntimeError("target offset-zero MMAP2 record not found")
    executable_loads = [load for load in loads if load["executable"]]
    intervals = [
        (load_bias + int(load["p_vaddr"], 16), load_bias + int(load["p_vaddr_end"], 16))
        for load in executable_loads
    ]
    marker_start = load_bias + symbols["profiled_measured_case"]["st_value"]
    marker_end = marker_start + symbols["profiled_measured_case"]["size"]
    marker_return = marker_start + 0x1B
    measure_return = load_bias + symbols["measure_one"]["st_value"] + 0xF3
    run_start = load_bias + symbols["run_case"]["st_value"]
    run_end = run_start + symbols["run_case"]["size"]
    marker_samples = [sample for sample in target_samples if sample["marker_range_words"]]
    marker_returns = [sample for sample in target_samples if sample["marker_return_words"]]
    measure_returns = [sample for sample in target_samples if sample["measure_one_return_words"]]
    run_samples = [sample for sample in target_samples if sample["run_case_range_words"]]
    all_word_offsets = [
        word["stack_relative_byte"]
        for sample in target_samples
        for key in ("marker_range_words", "marker_return_words", "measure_one_return_words", "run_case_range_words")
        for word in sample[key]
    ]
    result: dict[str, Any] = {
        "schema": "litchi.performance.0972.corrected_raw_stack_scan.v2",
        "status": "complete"
        if returncode == 0 and sample_count == sample_line_count and not stack_layout_failures and not event_mismatches
        else "diagnostic_complete_with_validation_findings",
        "scope": "offline scan of retained perf.data; no benchmark, build, workload, or source/runtime/cache mutation",
        "inputs": {
            "perf_data": {"path": str(DATA), "bytes": DATA.stat().st_size, "sha256": sha256(DATA)},
            "binary": {"path": str(BINARY), "bytes": BINARY.stat().st_size, "sha256": sha256(BINARY), "build_id": TARGET_BUILD_ID},
            "perf": {"path": str(PERF), "sha256": sha256(PERF)},
        },
        "command": [str(PERF), "script", "-D", "-i", str(DATA)],
        "perf_script_returncode": returncode,
        "perf_script_stderr": stderr,
        "parser": {
            "raw_event_count": raw_event_count,
            "sample_line_count": sample_line_count,
            "sample_count": sample_count,
            "sample_misc_header_counts": {str(k): v for k, v in sorted(sample_misc.items())},
            "record_type_counts": {str(k): v for k, v in sorted(record_types.items())},
            "record_header_misc_counts": {str(k): v for k, v in sorted(header_misc.items())},
            "record_size_counts": {str(k): v for k, v in sorted(record_sizes.items())},
            "event_header_mismatches": event_mismatches,
            "stack_layout_failures": stack_layout_failures,
            "little_endian_u64_offsets": {
                "ip": 8,
                "timestamp": 24,
                "period": 40,
                "callchain_nr": 48,
                "stack_size": STACK_SIZE_FIELD_OFFSET,
                "stack_bytes": STACK_BYTES_OFFSET,
            },
            "expected_sample_type_layout": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
            "expected_regs_user_mask": hex(EXPECTED_SAMPLE_REG_MASK),
            "callchain_is_zero_because_dwarf_stack_dump_is_used": True,
            "all_sample_records_scanned_without_misc_filter": True,
        },
        "mmap_binding": {
            "target_path": TARGET_PATH,
            "target_maps": target_maps,
            "offset_zero_load_bias": hex(load_bias),
            "binding_rule": "target PERF_RECORD_MMAP2 with target path and pgoff=0; map start is file-offset-zero load bias",
        },
        "elf_mapping_binding": {
            "identity": elf_identity,
            "pt_loads": loads,
            "executable_intervals": [
                {"start": hex(start), "end": hex(end), "source": "PF_X PT_LOAD p_vaddr..p_vaddr+p_memsz"}
                for start, end in intervals
            ],
            "page_mapped_executable_intervals": [
                {
                    "start": hex(align_down(start)),
                    "end": hex(align_up(end)),
                    "source": "page-aligned view of PF_X PT_LOAD",
                }
                for start, end in intervals
            ],
            "parse_sprms_st_value": hex(symbols["parse_sprms"]["st_value"]),
            "parse_sprms_sample_ip": hex(0x64D1E504A6F2),
            "parse_sprms_sample_offset": hex(0x64D1E504A6F2 - load_bias),
            "parse_sprms_sample_delta": hex(0x64D1E504A6F2 - load_bias - symbols["parse_sprms"]["st_value"]),
            "old_0971_scanner_base": hex(0x64D1E4E5C000),
            "old_scanner_error": hex(0x1000),
        },
        "symbol_binding": {
            name: {
                "st_value": hex(info["st_value"]),
                "size": hex(info["size"]),
                "runtime_start": hex(load_bias + info["st_value"]),
                "runtime_end": hex(load_bias + info["st_value"] + info["size"]),
            }
            for name, info in symbols.items()
        },
        "address_ranges": {
            "target_executable_intervals": [[hex(start), hex(end)] for start, end in intervals],
            "marker": [hex(marker_start), hex(marker_end)],
            "marker_direct_return": hex(marker_return),
            "measure_one_return": hex(measure_return),
            "run_case": [hex(run_start), hex(run_end)],
        },
        "target_leaf_summary": {
            "sample_count": len(target_samples),
            "period": sum(int(sample["period"]) for sample in target_samples),
            "marker_range_sample_count": len(marker_samples),
            "marker_range_period": sum(int(sample["period"]) for sample in marker_samples),
            "marker_return_sample_count": len(marker_returns),
            "marker_return_period": sum(int(sample["period"]) for sample in marker_returns),
            "measure_one_return_sample_count": len(measure_returns),
            "measure_one_return_period": sum(int(sample["period"]) for sample in measure_returns),
            "run_case_range_sample_count": len(run_samples),
            "run_case_range_period": sum(int(sample["period"]) for sample in run_samples),
            "raw_word_relative_offset_min": min(all_word_offsets) if all_word_offsets else None,
            "raw_word_relative_offset_max": max(all_word_offsets) if all_word_offsets else None,
            "raw_word_offset_units": "bytes relative to STACK_USER payload; record_byte = 232 + stack_relative_byte",
        },
        "target_leaf_sample_details": target_samples,
        "interpretation": {
            "raw_words_are_potential_values_only": True,
            "no_live_ancestry_claim": True,
            "no_phase_ownership_claim": True,
            "all_1387_sample_records_in_denominator": True,
            "pt_load_exact_interval_used_for_leaf_classification": True,
            "1377_vs_1387": "The historical text filter selected PERF_RECORD_SAMPLE(...,0x2) user records (1377); this raw parser scans all 1387 sample records, including 10 misc=0x1 kernel-IP records. This is a filter denominator difference, not event loss.",
            "corrective_finding": "The old scanner compared against addresses shifted +0x1000 and also used a hardcoded truncated executable upper bound. This v2 derives PF_X PT_LOAD intervals from the retained ELF and the retained offset-zero MMAP2 binding. Raw word matches remain fallible values, not validated ancestry.",
        },
    }
    return result


if __name__ == "__main__":
    result = scan()
    OUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "status": result["status"],
                "sample_count": result["parser"]["sample_count"],
                "misc": result["parser"]["sample_misc_header_counts"],
                "target_samples": result["target_leaf_summary"]["sample_count"],
                "target_period": result["target_leaf_summary"]["period"],
                "marker_range_samples": result["target_leaf_summary"]["marker_range_sample_count"],
                "marker_return_samples": result["target_leaf_summary"]["marker_return_sample_count"],
                "measure_one_return_samples": result["target_leaf_summary"]["measure_one_return_sample_count"],
                "run_case_range_samples": result["target_leaf_summary"]["run_case_range_sample_count"],
                "raw_word_relative_offset_max": result["target_leaf_summary"]["raw_word_relative_offset_max"],
            },
            sort_keys=True,
        )
    )
