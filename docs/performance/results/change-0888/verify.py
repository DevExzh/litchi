#!/usr/bin/env python3
"""Independently replay diagnostic custody, semantics, and raw mapping totals."""
import argparse
import datetime
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys

P = Path(__file__).resolve().parent
PREVIOUS = P.parent / 'change-0887'
sys.path.insert(0, str(PREVIOUS))
import experiment


def read(path):
    return json.loads(path.read_text())


def sha(path):
    h = hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def check(record):
    path = Path(record['path'])
    assert path.stat().st_size == record['bytes'], str(path)
    assert sha(path) == record['sha256'], str(path)


def sums(snapshot):
    # Independent VMA parser: classify only exact executable path; keep COW
    # Anonymous separately from its file-backed VMA category.
    totals = {'rss_kib': 0, 'exe_rss_kib': 0, 'exe_anonymous_kib': 0,
              'heap_rss_kib': 0, 'stack_rss_kib': 0, 'anonymous_kib': 0,
              'file_pmd_mapped_kib': 0, 'anon_huge_pages_kib': 0}
    name = None
    for line in snapshot['smaps'].splitlines():
        if re.match(r'^[0-9a-f]+-[0-9a-f]+ ', line):
            parts = line.split(maxsplit=5)
            name = parts[5] if len(parts) == 6 else ''
        elif name is not None and ':' in line:
            key, value = line.split(':', 1)
            if key not in ('Rss', 'Anonymous', 'FilePmdMapped', 'AnonHugePages'):
                continue
            amount, unit = value.split()
            assert unit == 'kB'
            amount = int(amount)
            if key == 'Rss':
                totals['rss_kib'] += amount
                if name == snapshot['exe']:
                    totals['exe_rss_kib'] += amount
                elif name == '[heap]':
                    totals['heap_rss_kib'] += amount
                elif name.startswith('[stack'):
                    totals['stack_rss_kib'] += amount
            elif key == 'Anonymous':
                totals['anonymous_kib'] += amount
                if name == snapshot['exe']:
                    totals['exe_anonymous_kib'] += amount
            elif key == 'FilePmdMapped':
                totals['file_pmd_mapped_kib'] += amount
            else:
                totals['anon_huge_pages_kib'] += amount
    return totals


def verify(packet):
    plan = read(packet / 'plan.json')
    capture = read(packet / 'capture.json')
    check(capture['plan'])
    for item in [*plan['helpers'], *plan['tools'], plan['previous_plan']]:
        check(item)
    assert capture['status'] == 'pass' and capture['latency_claim'] is None
    assert plan['observer_affinity'] == [13] and plan['cpu'] == '12'
    assert plan['warmup'] == 3 and plan['samples'] == 200
    assert plan['environment'] == experiment.CHILD_ENV
    assert plan['child_cwd'] == str(PREVIOUS)
    assert plan['no_retries_or_replacements'] is True
    copy = 'copy_proof' in plan
    variants = ['before-original', 'after-original', 'before-copy', 'after-copy']
    expected_rows = ([{'ordinal': b * 4 + i, 'block': b, 'role': role} for b in range(4) for i, role in enumerate(variants[b:] + variants[:b])] if copy else [{'ordinal': b * 2 + i, 'block': b, 'role': role} for b in range(4) for i, role in enumerate(('before', 'after') if b % 2 == 0 else ('after', 'before'))])
    assert plan['rows'] == expected_rows
    assert capture['children'] == len(expected_rows)
    assert capture['measured_iterations'] == len(expected_rows) * 200
    assert capture['warmups'] == len(expected_rows) * 3
    old_plan = read(PREVIOUS / 'plan.json')
    identities = {role: experiment._binary(old_plan, role, 'native') for role in ('before', 'after')}
    expected = experiment._load_expected_corpora()[experiment.PRIMARY_CASE]
    if copy:
        for role, proof in plan['copy_proof'].items():
            check(proof['source'])
            assert proof['source']['sha256'] == proof['copy']['sha256'] == identities[role]['identity']['sha256']
            assert proof['source']['bytes'] == proof['copy']['bytes']
            assert proof['source_device_inode'] != proof['copy_device_inode']
            assert proof['mode'] == identities[role]['identity']['mode']
    children = []
    previous_end = datetime.datetime.fromisoformat(plan['frozen_utc'])
    for row in expected_rows:
        role = row['role'].split('-')[0]
        binary = plan['binaries'][row['role']]
        assert binary['sha256'] == identities[role]['identity']['sha256']
        assert binary['bytes'] == identities[role]['identity']['bytes']
        # Original bytes remain available after temporary executable copies
        # are removed; report identity is checked against its captured path.
        identity = {'path': binary['path'], 'identity': {**identities[role]['identity'], 'path': binary['path']}}
        folder = packet / 'runs' / f"{row['ordinal']:02d}-{row['role']}"
        receipt = read(folder / 'receipt.json')
        command = read(folder / 'command.json')
        assert receipt['status'] == 'pass' and receipt['row'] == row
        assert set(receipt['files']) == {f.name for f in folder.iterdir() if f.name != 'receipt.json'}
        for item in receipt['files'].values():
            check(item)
        argv = experiment._command(binary['path'], plan['case'], 3, 200, folder / 'report.json', folder / 'producer.json', folder / 'peak-rss-kib.txt')
        assert command['argv'] == argv and command['cwd'] == plan['child_cwd']
        assert command['environment'] == plan['environment'] and command['row'] == row
        assert command['exit_code'] == 0 and command['plan_sha256'] == sha(packet / 'plan.json')
        start, end = [datetime.datetime.fromisoformat(command[k]) for k in ('started_utc', 'finished_utc')]
        assert previous_end <= start <= end
        previous_end = end
        experiment._validate_child(read(folder / 'report.json'), read(folder / 'producer.json'), plan['case'], 'native', 3, 200, expected, identity)
        peak = experiment._read_peak(folder / 'peak-rss-kib.txt')
        assert peak == receipt['peak_rss_kib']
        observed = read(folder / 'observations.json')
        with gzip.open(folder / 'snapshots.jsonl.gz', 'rt') as f:
            snapshots = [json.loads(line) for line in f]
        assert len(snapshots) == len(observed['snapshots']) == receipt['snapshots'] > 0
        assert len({s['pid'] for s in snapshots}) == 1
        previous_clock = -1
        totals = []
        for index, (snapshot, metadata) in enumerate(zip(snapshots, observed['snapshots'])):
            assert {k: v for k, v in snapshot.items() if k not in ('status', 'smaps')} == metadata
            assert snapshot['index'] == index and snapshot['exe'] == binary['path']
            assert snapshot['monotonic_ns'] > previous_clock
            previous_clock = snapshot['monotonic_ns']
            pid = re.search(r'^Pid:\s+(\d+)$', snapshot['status'], flags=re.M)
            assert pid and int(pid[1]) == snapshot['pid']
            totals.append(sums(snapshot))
        maxima = {k: max(t[k] for t in totals) for k in totals[0]}
        at_peak = totals[max(range(len(totals)), key=lambda i: totals[i]['rss_kib'])]
        children.append({'row': row, 'snapshots': len(totals), 'gnu_time_peak_rss_kib': peak, 'sampled_maxima_separate_not_simultaneous': maxima, 'at_sampled_total_rss_peak': at_peak})
    return {'status': 'pass', 'packet': str(packet), 'plan_sha256': sha(packet / 'plan.json'), 'children': children, 'semantics_pass': True, 'copy_binaries_required_for_replay': False, 'scope': 'Sampled mapping maxima remain separate from GNU time high-water; no latency or adoption claim.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        raise ValueError('verification output exists')
    result = {'status': 'pass', 'initial': verify(P), 'copy_probe': verify(P / 'copy-probe')}
    with args.output.open('x') as f:
        json.dump(result, f, indent=2, sort_keys=True)
        f.write('\n')
    print('24 children: custody, semantics, raw VMA totals replayed')


if __name__ == '__main__':
    main()
