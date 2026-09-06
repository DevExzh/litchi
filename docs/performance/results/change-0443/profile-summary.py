#!/usr/bin/env python3
"""Derive whole-process counter observations from retained profile artifacts."""
import argparse
import gzip
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def artifact(row):
    path = ROOT / row['path']
    raw = path.read_bytes() if path.exists() else gzip.decompress(Path(str(path) + '.gz').read_bytes())
    assert len(raw) == row['bytes'] and hashlib.sha256(raw).hexdigest() == row['sha256']
    return raw.decode()


def derive():
    roles = {}
    for role in ('before', 'after'):
        stat = json.loads((ROOT / f'profiles/{role}/stat/formal/receipt.json').read_text())
        record = json.loads((ROOT / f'profiles/{role}/record/formal/receipt.json').read_text())
        assert stat['status'] == record['status'] == 'pass'
        counters = {}
        for line in artifact(stat['artifacts']['perf_stat']).splitlines():
            columns = line.split(',')
            if len(columns) >= 3 and columns[2] in stat['stat_events']:
                counters[columns[2]] = int(columns[0])
        assert set(counters) == set(stat['stat_events'])
        roles[role] = {
            'counters': counters,
            'ipc': counters['instructions:u'] / counters['cycles:u'],
            'perf_script_addr2line_warnings': artifact(record['artifacts']['perf_script']).count('could not read first record'),
            'perf_report_addr2line_warnings': artifact(record['artifacts']['perf_report']).count('could not read first record'),
        }
    changes = {}
    for name, before in roles['before']['counters'].items():
        changes[name] = (roles['after']['counters'][name] - before) * 100 / before if before else None
    return {
        'change': 443, 'roles': roles, 'counter_change_percent': changes,
        'scope': 'Whole fresh executable including fixture setup, warmups and Rust oracle work. Counters are not operation-only causal attribution. Zero L1 values support no cache-miss claim.',
    }


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = derive()
    path = ROOT / 'profile-summary.json'
    if args.check:
        assert json.loads(path.read_text()) == result
    else:
        path.write_text(json.dumps(result, indent=2) + '\n')
    print('VALID: profile summary')
