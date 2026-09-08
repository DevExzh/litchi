#!/usr/bin/env python3
"""Retain wider process CPU/syscall profiles, excluded from formal samples."""
import gzip
import shutil
import subprocess
import sys
from common import ROOT, REPO, TEMP, ENV, read, write, meta, sha


def gate(label, argv):
    result = subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'),
                             label, *argv], cwd=REPO, env=ENV)
    return dict(exit_code=result.returncode, path=f'validation/{label}.json',
                sha256=sha(ROOT / 'validation' / f'{label}.json'))


def main():
    binary = read(ROOT / 'binaries.json')['normal']
    assert meta(binary['path']) == {key: binary[key] for key in ('bytes', 'sha256')}
    directory = ROOT / 'profiles'
    directory.mkdir()
    records = {}
    for kind in ('perf-stat', 'perf-record', 'strace'):
        operation = [binary['path'], '--mode', 'total', '--counts', '131072',
                     '--samples', '30', '--warmups', '3', '--json',
                     str(directory / f'{kind}.report.json')]
        if kind == 'perf-stat':
            prefix = ['perf', 'stat', '-x,', '-o', str(directory / 'perf-stat.csv'),
                      '-e', 'cycles,instructions,branches,branch-misses,cache-misses,page-faults', '--']
        elif kind == 'perf-record':
            prefix = ['perf', 'record', '--call-graph', 'fp', '-F', '99',
                      '-o', str(TEMP / 'perf.data'), '--']
        else:
            prefix = ['strace', '-c', '-o', str(directory / 'strace.txt'), '--']
        records[kind] = gate(f'profile-{kind}', ['/usr/bin/taskset', '-c', '2', *prefix, *operation])
    if records['perf-record']['exit_code'] == 0:
        records['perf-report'] = gate('profile-perf-report', [
            'perf', 'report', '--stdio', '--no-children', '--percent-limit', '0',
            '-i', str(TEMP / 'perf.data')])
        records['perf-script'] = gate('profile-perf-script', [
            'perf', 'script', '-i', str(TEMP / 'perf.data')])
        raw = TEMP / 'perf.data'
        raw_identity = meta(raw)
        with raw.open('rb') as source, (directory / 'perf.data.gz').open('xb') as target:
            with gzip.GzipFile(fileobj=target, mode='wb', filename='', mtime=0) as packed:
                shutil.copyfileobj(source, packed)
        import hashlib
        with gzip.open(directory / 'perf.data.gz', 'rb') as source:
            assert hashlib.file_digest(source, 'sha256').hexdigest() == raw_identity['sha256']
        records['raw_profile'] = dict(uncompressed=raw_identity,
                                     compressed=meta(directory / 'perf.data.gz'))
        raw.unlink()
    write(ROOT / 'profiles.json', dict(
        schema='docx-plain-paragraph-tail-append-profiles-v1',
        scope='whole process including corpus generation, oracles, warmup, measured lifecycles, report serialization and teardown',
        formal_samples=False, binary=binary, records=records,
        status='pass' if all(r.get('exit_code', 0) == 0 for r in records.values()) else 'partial'))


if __name__ == '__main__':
    main()
