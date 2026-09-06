#!/usr/bin/env python3
"""Re-derive the pre-implementation ODT text batching hypothesis evidence."""
import argparse
import gzip
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent

def raw(path):
    return path.read_bytes() if path.exists() else gzip.decompress(Path(str(path) + '.gz').read_bytes())

def sha(value):
    return hashlib.sha256(value).hexdigest()

def load(path):
    return json.loads(raw(path))

def derive():
    inputs = {}
    rows = []
    for shape, count in [('tiny', 64), ('medium', 8192), ('large', 32768)]:
        reports = {}
        for mode in ('normal', 'allocator'):
            path = ROOT / 'pilots/before/initial' / f'{mode}-{shape}.json'
            receipt_path = path.with_name(f'{mode}-{shape}-receipt.json')
            receipt = load(receipt_path)
            assert receipt['status'] == 'pass' and receipt['source_unchanged'] is True
            assert sha(raw(path)) == receipt['artifacts'][str(path.relative_to(ROOT))]['sha256']
            inputs[str(path.relative_to(ROOT))] = sha(raw(path))
            inputs[str(receipt_path.relative_to(ROOT))] = sha(raw(receipt_path))
            reports[mode] = load(path)['results'][0]
        allocation = reports['allocator']['operation_metrics']['allocation']
        vectors = {key: value['values'] for key, value in allocation.items() if isinstance(value, dict)}
        rows.append({'shape': shape, 'paragraphs': count,
            'pilot_normal_p50_ns': reports['normal']['elapsed_ns']['p50'],
            'allocation_calls': vectors['allocation_calls'], 'allocated_bytes': vectors['allocated_bytes'],
            'region_peak_above_entry_bytes': [peak - entry for peak, entry in zip(vectors['region_peak_live_bytes'], vectors['live_bytes_before'], strict=True)],
            'operation_live_delta_bytes': [end - start for end, start in zip(vectors['live_bytes_after'], vectors['live_bytes_before'], strict=True)]})
    path = ROOT / 'preparatory/baseline-profile/provenance.json'
    provenance = load(path)
    inputs[str(path.relative_to(ROOT))] = sha(raw(path))
    assert provenance['baseline_binary_copies_sha256'] == sha(raw(ROOT / 'before/binary-copies.json'))
    copies = load(ROOT / 'before/binary-copies.json')
    profiles = {}
    for kind, item in provenance['profiles'].items():
        artifact = item['artifact']; payload = raw(ROOT / artifact['path'])
        assert len(payload) == artifact['bytes'] and sha(payload) == artifact['sha256']
        assert item['prior_executable_sha256'] == copies['normal']['sha256']
        inputs[artifact['path']] = sha(payload)
        profiles[kind] = dict(item)
        if kind == 'record':
            text = payload.decode()
            profiles[kind]['top_self_rows'] = [line.strip() for line in text.splitlines() if re.match(r'^\s+\d+\.\d+%\s+normal\s+', line)][:15]
            profiles[kind]['addr2line_warning_count'] = text.count('could not read first record')
    return {'change': 436, 'phase': 'before-production-implementation', 'status': 'baseline-evidence-retained', 'inputs': inputs, 'rows': rows, 'profiles': profiles,
        'hypothesis': 'Batch already-safe ordinary UTF-8 into borrowed spans of at most 256 bytes to reduce repeated hierarchical Work charges and scratch writes; preserve scalar cancellation polling and scalar fallback at XML and Work limits.',
        'claim_limits': 'Three-sample fresh pilots are preparatory observations. Prior whole-process profile is reused only because both current baseline executable hashes exactly match 0435 after binaries; it includes setup, hashing, warmups, samples and oracle. A 45.74% consume self share motivates measurement but does not establish an operation-only causal fraction. No latency gain or RSS improvement is yet claimed.'}

def main():
    parser = argparse.ArgumentParser(); parser.add_argument('--verify', action='store_true'); args = parser.parse_args()
    result = derive(); path = ROOT / 'batching-hypothesis.json'
    if args.verify:
        assert load(path) == result
    else:
        with path.open('x') as output:
            output.write(json.dumps(result, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'rows': result['rows']}))

if __name__ == '__main__':
    main()
