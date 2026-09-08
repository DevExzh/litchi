#!/usr/bin/env python3
"""Paired whole-process counters, excluded from formal sample statistics."""
import csv
import subprocess
import sys
from common import ROOT, REPO, ENV, read, write, meta

EVENTS = ('cycles', 'instructions', 'branches', 'branch-misses', 'cache-misses', 'page-faults')


def main():
    directory = ROOT / 'profiles'
    directory.mkdir()
    records = {}
    for arm in ('control', 'candidate'):
        binary = read(ROOT / f'{arm}-binaries.json')['normal']
        assert meta(binary['path']) == {key: binary[key] for key in ('bytes', 'sha256')}
        report = directory / f'{arm}.report.json'
        counters = directory / f'{arm}.csv'
        label = f'profile-{arm}'
        argv = ['/usr/bin/taskset', '-c', '2', 'perf', 'stat', '-x,', '-o', str(counters),
                '-e', ','.join(EVENTS), '--', binary['path'], '--mode', 'total',
                '--counts', '131072', '--samples', '30', '--warmups', '3', '--json', str(report)]
        result = subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'), label, *argv],
                                cwd=REPO, env=ENV)
        records[arm] = dict(binary=binary, exit_code=result.returncode,
                            receipt=dict(path=f'validation/{label}.json', **meta(ROOT / 'validation' / f'{label}.json')))
        if result.returncode == 0:
            values = {}
            with counters.open() as stream:
                for row in csv.reader(stream):
                    if not row or row[0].startswith('#'):
                        continue
                    assert row[2] in EVENTS and row[2] not in values
                    values[row[2]] = dict(value=int(row[0]), running_percent=float(row[4]))
            assert set(values) == set(EVENTS)
            records[arm].update(counters=dict(path=f'profiles/{arm}.csv', **meta(counters)),
                                report=dict(path=f'profiles/{arm}.report.json', **meta(report)),
                                values=values)
    write(ROOT / 'profiles.json', dict(
        schema='docx-borrowed-names-paired-pmu-v1', records=records,
        formal_samples=False, excluded_samples=60,
        scope='separate whole-process normal runs including corpus/oracles, warmups, 30 lifecycles and report teardown; no isolated operation PMU delta',
        status='pass' if all(x['exit_code'] == 0 for x in records.values()) else 'partial'))


if __name__ == '__main__':
    main()
