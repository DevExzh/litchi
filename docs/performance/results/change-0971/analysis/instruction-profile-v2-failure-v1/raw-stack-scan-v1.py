#!/usr/bin/env python3
"""Offline scan of one retained perf.data; never launches the benchmark."""
from __future__ import annotations
import collections
import json
import re
import struct
import subprocess
from pathlib import Path

DATA = Path('/home/zhuhe/litchi-performance-0971/capture/instruction-profile-v2/child-01/child-01.perf.data')
OUT = Path('/home/zhuhe/litchi-performance-0971/analysis/instruction-profile-v2-failure-v1/raw-stack-scan-v1.json')
BASE = 0x64d1e4e5c000
TEXT_LO = BASE + 0xE4000
TEXT_HI = BASE + 0x36B000
MARKER_LO = BASE + 0xEC4F0
MARKER_HI = BASE + 0xEC570
MARKER_RETURN = BASE + 0xEC50B
CALLER_RETURN = BASE + 0xE6D93
RUN_CASE_LO = BASE + 0xF15F0
RUN_CASE_HI = BASE + 0xF2700

RAW_EVENT = re.compile(r'^.*raw event: size (\d+) bytes$')
HEX_LINE = re.compile(r'^\.\s+([0-9a-fA-F]+):\s+((?:[0-9a-fA-F]{2}\s+)+)')
SAMPLE = re.compile(
    r'^(\d+)\s+(\d+) .*PERF_RECORD_SAMPLE\(IP, 0x2\):\s+'
    r'\d+/\d+:\s+0x([0-9a-fA-F]+)\s+period:\s+(\d+)'
)

proc = subprocess.Popen(
    ['/usr/bin/perf', 'script', '-D', '-i', str(DATA)],
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True,
    bufsize=1,
)
assert proc.stdout is not None and proc.stderr is not None
raw: bytearray | None = None
sample_records = 0
target_samples: list[dict[str, object]] = []
record_sizes: collections.Counter[int] = collections.Counter()
for line in proc.stdout:
    line = line.rstrip('\n')
    match = RAW_EVENT.match(line)
    if match:
        raw = bytearray(int(match.group(1)))
        record_sizes[len(raw)] += 1
        continue
    match = HEX_LINE.match(line)
    if match and raw is not None:
        offset = int(match.group(1), 16)
        payload = bytes.fromhex(match.group(2))
        raw[offset:offset + len(payload)] = payload
        continue
    match = SAMPLE.match(line)
    if not match or raw is None:
        continue
    sample_records += 1
    timestamp_ns = int(match.group(2))
    ip = int(match.group(3), 16)
    period = int(match.group(4))
    if TEXT_LO <= ip < TEXT_HI:
        words: list[tuple[int, int]] = []
        for offset in range(0, len(raw) - 7, 8):
            value = struct.unpack_from('<Q', raw, offset)[0]
            if TEXT_LO <= value < TEXT_HI:
                words.append((offset, value))
        target_samples.append({
            'timestamp_ns': timestamp_ns,
            'ip': hex(ip),
            'period': period,
            'record_bytes': len(raw),
            'target_range_word_count': len(words),
            'marker_range_offsets': [o for o, v in words if MARKER_LO <= v < MARKER_HI],
            'marker_return_offsets': [o for o, v in words if v == MARKER_RETURN],
            'caller_return_offsets': [o for o, v in words if v == CALLER_RETURN],
            'run_case_range_offsets': [o for o, v in words if RUN_CASE_LO <= v < RUN_CASE_HI],
        })
    raw = None
proc.wait()
stderr = proc.stderr.read()
result = {
    'schema': 'litchi.performance.0971.raw_stack_scan.v1',
    'status': 'complete',
    'input': str(DATA),
    'perf_script_dump_argv': ['/usr/bin/perf', 'script', '-D', '-i', str(DATA)],
    'perf_script_dump_returncode': proc.returncode,
    'perf_script_dump_stderr': stderr,
    'binary_runtime_base_from_mmap_and_symbol_offset': hex(BASE),
    'target_text_runtime_range': [hex(TEXT_LO), hex(TEXT_HI)],
    'marker_runtime_range_from_nm': [hex(MARKER_LO), hex(MARKER_HI)],
    'marker_direct_call_return_runtime_address': hex(MARKER_RETURN),
    'measure_one_call_return_runtime_address': hex(CALLER_RETURN),
    'run_case_runtime_range_from_nm': [hex(RUN_CASE_LO), hex(RUN_CASE_HI)],
    'sample_records_seen_by_perf_dump_filter': sample_records,
    'record_size_counts': {str(k): v for k, v in sorted(record_sizes.items())},
    'target_leaf_ip_sample_count': len(target_samples),
    'target_leaf_ip_period': sum(int(s['period']) for s in target_samples),
    'target_leaf_ip_time_ns': ([min(int(s['timestamp_ns']) for s in target_samples), max(int(s['timestamp_ns']) for s in target_samples)] if target_samples else None),
    'target_leaf_samples_with_marker_range_word': sum(bool(s['marker_range_offsets']) for s in target_samples),
    'target_leaf_samples_with_marker_return_word': sum(bool(s['marker_return_offsets']) for s in target_samples),
    'target_leaf_samples_with_measure_one_return_word': sum(bool(s['caller_return_offsets']) for s in target_samples),
    'target_leaf_samples_with_run_case_range_word': sum(bool(s['run_case_range_offsets']) for s in target_samples),
    'target_leaf_sample_details': target_samples,
    'interpretation': 'Descriptive raw-record diagnostic only. It does not assign samples to a workload phase or establish a causal mechanism for missing marker ancestry.',
}
OUT.write_text(json.dumps(result, indent=2, sort_keys=True) + '\\n', encoding='utf-8')
print(json.dumps({k: result[k] for k in ('status', 'sample_records_seen_by_perf_dump_filter', 'target_leaf_ip_sample_count', 'target_leaf_ip_period', 'target_leaf_samples_with_marker_range_word', 'target_leaf_samples_with_marker_return_word', 'target_leaf_samples_with_measure_one_return_word', 'target_leaf_samples_with_run_case_range_word')}, sort_keys=True))
