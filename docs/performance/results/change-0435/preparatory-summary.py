#!/usr/bin/env python3
"""Derive the pre-implementation buffered ODT evidence without rerunning workloads."""
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def load(path):
    return json.loads(path.read_text())


def main():
    rows = []
    inputs = {}
    for shape in ('tiny', 'medium', 'large'):
        reports = {}
        for mode in ('normal', 'allocator'):
            path = ROOT / 'pilots/before-buffered/v3' / f'{mode}-{shape}.json'
            receipt_path = path.with_name(f'{mode}-{shape}-receipt.json')
            receipt = load(receipt_path)
            assert receipt['status'] == 'pass' and receipt['source_unchanged'] is True
            assert sha(path) == receipt['artifacts'][str(path.relative_to(ROOT))]['sha256']
            inputs[str(path.relative_to(ROOT))] = sha(path)
            inputs[str(receipt_path.relative_to(ROOT))] = sha(receipt_path)
            reports[mode] = load(path)['results'][0]
        normal, allocator = reports['normal'], reports['allocator']
        allocation = allocator['operation_metrics']['allocation']
        vectors = {key: value['values'] for key, value in allocation.items() if isinstance(value, dict)}
        rows.append({
            'shape': shape, 'paragraphs': normal['corpus']['entry_count'],
            'pilot_normal_p50_ns': normal['elapsed_ns']['p50'],
            'allocation_calls': vectors['allocation_calls'],
            'allocated_bytes': vectors['allocated_bytes'],
            'region_peak_above_entry_bytes': [peak - entry for peak, entry in
                zip(vectors['region_peak_live_bytes'], vectors['live_bytes_before'], strict=True)],
            'operation_live_delta_bytes': [end - start for end, start in
                zip(vectors['live_bytes_after'], vectors['live_bytes_before'], strict=True)],
        })
    profiles = {}
    for kind in ('stat', 'record'):
        directory = ROOT / 'profiles/preparatory-before-buffered/v3' / kind
        receipt = load(directory / 'receipt.json')
        assert receipt['status'] == 'pass' and receipt['source_unchanged'] is True
        inputs[str((directory / 'receipt.json').relative_to(ROOT))] = sha(directory / 'receipt.json')
        profiles[kind] = {'receipt': str((directory / 'receipt.json').relative_to(ROOT)),
                          'scope': receipt['scope']}
    report_path = ROOT / 'profiles/preparatory-before-buffered/v3/record/perf-report.txt'
    raw = report_path.read_text()
    profiles['record']['addr2line_warning_count'] = raw.count('could not read first record')
    profiles['record']['top_self_rows'] = [line.strip() for line in raw.splitlines()
        if re.match(r'^\s+\d+\.\d+%\s+normal\s+', line)][:15]
    output = {
        'change': 435, 'phase': 'before-production-implementation',
        'status': 'baseline-evidence-retained', 'inputs': inputs, 'rows': rows,
        'profiles': profiles,
        'hypothesis': 'Once-consumed bounded paragraph publication can remove Builder document-proportional model/content/archive retention while preserving the supported fresh paragraph semantics.',
        'claim_limits': 'Three-sample pilots establish preparatory counter observations, not formal latency evidence. Allocator vectors cover the timed operation; region peak above entry is not RSS or total process memory. Whole-process profiles include corpus setup, binary/output hashing, warmups, timed calls, and untimed oracles. L1 zero and absent LLC counters do not prove zero misses. No causal speedup or 10x claim is made.',
    }
    path = ROOT / 'buffered-hypothesis.json'
    with path.open('x') as stream:
        json.dump(output, stream, indent=2)
        stream.write('\n')
    print(json.dumps(rows))


if __name__ == '__main__':
    main()
