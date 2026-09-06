#!/usr/bin/env python3
"""Capture preparatory ODP pilots; these are separate from the formal matrix."""
import argparse
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
ROLES = {'before-buffered': ('before', 'odp_buffered_create'),
         'after-buffered': ('after', 'odp_buffered_create'),
         'after-streaming': ('after', 'odp_streaming_create')}


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def write(path, value):
    path.write_text(json.dumps(value, indent=2) + '\n')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--role', choices=tuple(ROLES), required=True)
    parser.add_argument('--attempt', default='initial')
    args = parser.parse_args()
    assert args.attempt.isalnum()
    build_role, selector = ROLES[args.role]
    spec = importlib.util.spec_from_file_location('custody0437', ROOT / 'check.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    build_path = ROOT / 'checks' / f'{build_role}-build.json'
    build = json.loads(build_path.read_text())
    assert build['status'] == 'pass' and build['source_unchanged'] is True
    expected_source = build['source_before']
    assert custody.sources() == expected_source
    copies_path = ROOT / build_role / 'binary-copies.json'
    copies = json.loads(copies_path.read_text())
    oracle = ROOT / 'verify-report-candidate.py'
    protocol = ROOT / 'oracle-protocol.json'
    directory = ROOT / 'pilots' / args.role / args.attempt
    directory.mkdir(parents=True, exist_ok=False)
    for mode in ('normal', 'allocator'):
        binary = copies[mode]
        assert sha(Path(binary['path'])) == binary['sha256']
        for shape in ('tiny', 'medium', 'large'):
            stem = f'{mode}-{shape}'
            report = directory / f'{stem}.json'
            catalog = directory / f'{stem}-catalog.json'
            resource = directory / f'{stem}-resource.log'
            log = directory / f'{stem}.log'
            oracle_log = directory / f'{stem}-oracle.log'
            receipt = directory / f'{stem}-receipt.json'
            argv = ['taskset', '-c', '2', '/usr/bin/time', '-v', '-o', str(resource),
                    binary['path'], '--case', selector, '--semantic-shape', shape,
                    '--workers', '1', '--samples', '3', '--warmup', '1',
                    '--json', str(report), '--corpus-manifest', str(catalog)]
            row = {'change': 437, 'phase': 'preparatory-pilot', 'role': args.role,
                   'mode': mode, 'shape': shape, 'samples': 3, 'warmups': 1,
                   'argv': argv, 'cwd': str(REPO), 'binary': binary,
                   'driver_sha256': sha(Path(__file__)), 'oracle_sha256': sha(oracle),
                   'protocol_sha256': sha(protocol), 'oracle_path': oracle.name,
                   'protocol_path': protocol.name, 'build_receipt': str(build_path.relative_to(ROOT)),
                   'build_receipt_sha256': sha(build_path), 'source_before': custody.sources(),
                   'started_utc': now(), 'status': 'running'}
            write(receipt, row)
            try:
                with log.open('xb') as stream:
                    result = subprocess.run(argv, cwd=REPO, stdout=stream,
                                            stderr=subprocess.STDOUT)
                row['exit_code'] = result.returncode
                assert result.returncode == 0, 'pilot workload failed'
                oracle_argv = [sys.executable, '-B', str(oracle), '--report', str(report),
                               '--mode', mode, '--shape', shape, '--role', args.role,
                               '--samples', '3', '--warmups', '1']
                row['oracle_argv'] = oracle_argv
                with oracle_log.open('xb') as stream:
                    result = subprocess.run(oracle_argv, cwd=REPO, stdout=stream,
                                            stderr=subprocess.STDOUT)
                row['oracle_exit_code'] = result.returncode
                assert result.returncode == 0, 'pilot oracle failed'
                row['status'] = 'pass'
            except Exception as error:
                row['status'] = 'failed'
                row['error'] = repr(error)
            finally:
                row['source_after'] = custody.sources()
                row['source_unchanged'] = row['source_before'] == row['source_after'] == expected_source
                if not row['source_unchanged']:
                    row['status'] = 'failed'
                    row['error'] = 'source changed during pilot'
                row['finished_utc'] = now()
                row['artifacts'] = {str(path.relative_to(ROOT)): {
                    'bytes': path.stat().st_size, 'sha256': sha(path)}
                    for path in (report, catalog, resource, log, oracle_log) if path.exists()}
                write(receipt, row)
            print(row['status'], args.role, mode, shape, flush=True)
            if row['status'] != 'pass':
                return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
