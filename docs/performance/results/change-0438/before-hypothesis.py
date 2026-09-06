#!/usr/bin/env python3
"""Bind fresh ODP baseline pilots to the unchanged-code historical hotspot."""
import argparse
import gzip
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent


def raw(name):
    p = ROOT / name
    return p.read_bytes() if p.is_file() else gzip.decompress(Path(str(p) + '.gz').read_bytes())


def load(name):
    return json.loads(raw(name))


def bound(name):
    data = raw(name)
    return {'path': name, 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}


def derive():
    build = load('checks/before-build.json')
    binaries = load('before/binary-copies.json')
    reference = load('prior-profile/reference.json')
    assert reference['source_manifest_sha256'] == build['source_before']['sha256']
    profiles = {}
    for kind, entry in reference['profiles'].items():
        receipt = entry['receipt']
        assert receipt['source_manifest'] == build['source_before']
        assert receipt['binary']['sha256'] == binaries['normal']['sha256']
        name = entry['retained_hotspot_artifact']
        value = bound(name)
        record = receipt['artifacts']['perf_stat' if kind == 'stat' else 'perf_report']
        assert value['sha256'] == record['sha256'] and value['bytes'] == record['bytes']
        profiles[kind] = value
    report = raw(profiles['record']['path']).decode()
    symbols = [{'self_percent': float(m[1]), 'symbol': m[2]}
        for line in report.splitlines()
        if (m := re.match(r'^\s*(\d+\.\d+)%\s+normal\s+\S+\s+\[.\]\s+(.+)$', line))
        and float(m[1]) >= 1]
    rows = []
    for shape, slides in [('tiny', 64), ('medium', 4096), ('large', 8192)]:
        paths = [f'pilots/before-streaming/initial/{mode}-{shape}.json' for mode in ('normal', 'allocator')]
        normal, allocator = [load(name)['results'][0] for name in paths]
        a = allocator['operation_metrics']['allocation']
        metric = lambda key: a[key]['values']
        assert metric('live_bytes_before') == metric('live_bytes_after')
        peaks = [peak - entry for peak, entry in zip(metric('region_peak_live_bytes'), metric('live_bytes_before'), strict=True)]
        assert len(set(peaks)) == 1
        rows.append({'shape': shape, 'slides': slides,
            'normal_p50_ns': normal['elapsed_ns']['p50'],
            'allocation_calls': metric('allocation_calls'),
            'allocated_bytes': metric('allocated_bytes'),
            'peak_above_entry_bytes': peaks,
            'input_bindings': [bound(name) for name in paths]})
    return {'change': 438, 'before_revision': build['revision'],
        'source_manifest': build['source_before'],
        'baseline_pilots': rows,
        'historical_profile_bindings': profiles,
        'historical_profile_symbols_at_least_one_percent': symbols,
        'profile_scope': 'Historical 0437 whole executable; its binary and source exactly match this fresh baseline. Setup, corpus, hashing, warmups and samples are included; later Python oracle and symbolization are excluded. No operation-only causal fraction.',
        'hypothesis': 'Batch adjacent controlled fixed-markup pieces to remove repeated Work atomics and bounds/cancellation dispatch. Use bounded 256-byte stack scratch with no heap allocation; retain the original piece sequence on Work or XML bound refusal.',
        'preservation': 'Require exact same-API archive/content/semantic/sink identities and allocator vectors. Preserve original first failing piece, typed local/ancestor budget errors and rollback; cancellation remains bounded and failed current fragments remain unpublished.',
        'retention_rule': 'Keep only if both medium and large normal p50 improve at least five percent in both repeats, with nonoverlapping mean 95% intervals in the favorable direction, no allocation/peak increase, and every approximately five-percent latency/throughput/RSS or repeat flag reviewed explicitly. Otherwise revert speculative production complexity and retain the negative evidence.',
        'scope': 'Same ODP plain fresh creation workload, no new grammar, public API, limit profile, parallelism, existing-document append or iWork work.',
        'adr_tree': 'c950b6c8be822561b498d7bbe87c460873dcbf49'}


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--check', action='store_true')
    a = p.parse_args()
    value = derive()
    if a.check:
        assert load('before-hypothesis.json') == value
    else:
        target = ROOT / 'before-hypothesis.json'
        assert not target.exists()
        target.write_text(json.dumps(value, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'baseline_pilots': 6, 'profile_source_and_binary_match': True}))


if __name__ == '__main__':
    main()
