#!/usr/bin/env python3
"""Derive the pre-implementation ODP decision from retained exploratory inputs."""
import argparse
import gzip
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent


def raw(path):
    member = ROOT / path
    compressed = Path(str(member) + '.gz')
    if member.is_file():
        return member.read_bytes()
    assert compressed.is_file(), f'missing raw/gzip member: {path}'
    return gzip.decompress(compressed.read_bytes())


def load(path):
    return json.loads(raw(path))


def bound(path):
    value = raw(path)
    return {'path': path, 'bytes': len(value), 'sha256': hashlib.sha256(value).hexdigest()}


def derive():
    rows = []
    for shape, count in [('tiny', 64), ('medium', 4096), ('large', 8192)]:
        stem = 'pilots/before-buffered/initial/'
        normal_path = stem + f'normal-{shape}.json'
        allocator_path = stem + f'allocator-{shape}.json'
        normal = load(normal_path)['results'][0]
        allocation = load(allocator_path)['results'][0]['operation_metrics']['allocation']
        metric = lambda key: allocation[key]['values']
        peaks = [peak - entry for peak, entry in zip(
            metric('region_peak_live_bytes'), metric('live_bytes_before'), strict=True)]
        assert len(set(peaks)) == 1
        assert metric('live_bytes_after') == metric('live_bytes_before')
        assert metric('allocated_bytes') == metric('deallocated_bytes')
        rows.append({
            'shape': shape, 'slides': count, 'normal_p50_ns': normal['elapsed_ns']['p50'],
            'allocation_calls': metric('allocation_calls'),
            'allocated_bytes': metric('allocated_bytes'),
            'region_peak_above_entry_bytes': peaks, 'live_delta_bytes': [0] * 3,
            'archive_bytes': normal['corpus']['archive_bytes'],
            'content_xml_bytes': normal['corpus']['target_payload_bytes'],
            'inputs': [bound(normal_path), bound(allocator_path)],
        })
    profile_base = 'profiles/preparatory-before-buffered/initial/'
    stat_path = profile_base + 'stat/perf-stat.txt'
    record_path = profile_base + 'record/perf-report.txt'
    counters = {}
    for line in raw(stat_path).decode().splitlines():
        columns = line.split(',')
        if len(columns) > 2 and columns[0].isdigit():
            counters[columns[2]] = int(columns[0])
    raw_report = raw(record_path).decode()
    symbols = []
    for line in raw_report.splitlines():
        match = re.match(r'^\s*(\d+\.\d+)%\s+normal\s+\S+\s+\[.\]\s+(.+)$', line)
        if match and float(match[1]) >= 1:
            symbols.append({'self_percent': float(match[1]), 'symbol': match[2]})
    assert '# Total Lost Samples: 0' in raw_report
    result = {
        'change': 437, 'before_revision': load('checks/before-build.json')['revision'],
        'classification': 'Pre-implementation exploratory decision; pilots use three samples and one warmup, not the formal comparison.',
        'pilot_rows': rows,
        'profile': {
            'scope': 'Whole benchmark executable including corpus construction, built-in verification, binary hashing, warmups and samples; excludes the independent Python report oracle and symbolization.',
            'samples': 30, 'warmups': 3, 'slides': 8192, 'counters': counters,
            'ipc': counters['instructions:u'] / counters['cycles:u'],
            'branch_miss_percent': 100 * counters['branch-misses:u'] / counters['branches:u'],
            'symbols_at_least_one_percent': symbols, 'lost_samples': 0,
            'addr2line_warnings': raw_report.count('could not read first record'),
            'cache_limitations': 'L1 miss counter returned zero without a load denominator; LLC was not collected. No cache-rate claim.',
            'inputs': [bound(stat_path), bound(record_path)],
        },
        'decision': 'Implement and measure an opt-in bounded plain-slide source and sequential sink API, plus the narrow common balanced-prelude XML envelope enabler.',
        'mechanism': 'Avoid retaining the whole slide model, all generated page names, body/final content Strings and final archive Vec. Keep one source item, a fixed slide fragment window and bounded ZIP/manifest staging.',
        'acceptance': 'Require identical title/body, page/frame geometry, fixed styles/meta and package semantics. Measure constant operation memory as slide count grows; inspect all latency/RSS regressions above approximately five percent. Preserve the buffered control and existing limits. No speedup is assumed.',
        'limitations': 'Fresh plain titled slides only. The measured memmove/allocator symbols include setup and verification, so their whole-process percentages are not an operation-only Amdahl fraction. Native rendering/runtime is unavailable. Rich slides and append-to-existing are separate work.',
    }
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = derive()
    if args.check:
        assert load('before-hypothesis.json') == result, 'retained hypothesis differs'
    else:
        (ROOT / 'before-hypothesis.json').write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'decision': result['decision']}))


if __name__ == '__main__':
    main()
