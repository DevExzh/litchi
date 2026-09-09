#!/usr/bin/env python3
"""Attribute replay synchronization with retained binaries; never build or edit source."""
from __future__ import annotations

import argparse
import decimal
import fcntl
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import signal
import subprocess
import sys
from datetime import datetime, timezone

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
OLD = ROOT.parent / "change-0484"
sys.path.insert(0, str(OLD))
import measure_routes as routes

LOCK = Path('/home/zhuhe/.cache/litchi-goal-0484/cpu.lock')
ENV = dict(os.environ, LC_ALL='C', DEBUGINFOD_URLS='', PYTHONDONTWRITEBYTECODE='1')
CASE = routes.ROUTE_CASE_BY_LABEL['s64-a64-short-c64']
SPEC = routes.ROUTE_BY_NAME['file_store']
TRACE = 'read,write,pread64,pwrite64,readv,writev,lseek,openat,close,fsync,fdatasync,unlink,unlinkat,statx,fstat,newfstatat'
CALL = re.compile(r'^(?:\d+\s+)?(?P<time>\d+\.\d+)\s+(?P<call>fdatasync|fsync)\((?P<fd>\d+)<(?P<path>[^>]+)>\)\s+=\s+(?P<result>-?\d+)(?:\s+.*?)?\s+<(?P<duration>\d+\.\d+)>$')


def now():
    return datetime.now(timezone.utc).isoformat()


def read(path):
    return json.loads(Path(path).read_text())


def meta(path):
    path = Path(path)
    if path.is_symlink() or not path.is_file():
        raise ValueError(f'expected regular file: {path}')
    with path.open('rb') as stream:
        return {'bytes': path.stat().st_size, 'sha256': hashlib.file_digest(stream, 'sha256').hexdigest()}


def write(path, value):
    with Path(path).open('x') as stream:
        json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
        stream.write('\n')


def bindings():
    builds = {}
    for phase, number in [('before', '0487'), ('after', '0489')]:
        path = ROOT.parent / f'change-{number}' / 'build-normal.json'
        value = read(path)
        if meta(value['binary']['path']) != {k: value['binary'][k] for k in ('bytes', 'sha256')}:
            raise ValueError('retained executable changed')
        if not os.access(value['binary']['path'], os.X_OK):
            raise ValueError('retained executable is not executable')
        if meta(value['gate']['path'])['sha256'] != value['gate']['sha256']:
            raise ValueError('build gate changed')
        gate = read(value['gate']['path'])
        if gate['exit_code'] != 0 or not gate['source_unchanged'] or gate['source_after'] != value['source_after']:
            raise ValueError('build gate/source mismatch')
        builds[phase] = {'path': str(path), **meta(path), 'binary': value['binary'], 'source': value['source_after']}
    return builds


def source_check():
    build = read(ROOT.parent / 'change-0489' / 'build-normal.json')
    source = ROOT.parent / 'change-0489' / build['source_after']['path']
    expected = read(source)
    names = subprocess.check_output(['git', 'ls-files', '-co', '--exclude-standard', '-z'], cwd=REPO).decode().split('\0')
    selected = {name for name in names if name and (Path(name).suffix in {'.rs', '.toml', '.lock'} or (name.startswith('crates/') and '/src/' in name and Path(name).suffix == '.xml')) and (REPO / name).is_file()}
    if selected != set(expected) or any(meta(REPO / name)['sha256'] != digest for name, digest in expected.items()):
        raise ValueError('current source differs from retained 0489 executable')
    return {'path': str(source), **meta(source), 'files': len(expected)}


def runs():
    return [{'phase': phase, 'repeat': repeat, 'kind': kind,
             'label': f'{kind}-r{repeat}-{phase}',
             'samples': 30 if kind == 'sync' else 1,
             'warmups': 3 if kind == 'sync' else 1}
            for kind in ['sync', 'summary'] for repeat in [1, 2]
            for phase in (['before', 'after'] if repeat == 1 else ['after', 'before'])]


def protocol():
    return {'schema': 'docx-replay-sync-profile-protocol-v1', 'driver': meta(__file__),
            'validators': routes._script_hashes(), 'builds': bindings(), 'source': source_check(),
            'tools': {name: {'path': path, **meta(path)} if (path := shutil.which(name)) else {'status': 'unavailable'} for name in ['strace', 'time', 'taskset']},
            'machine': meta(ROOT / 'machine.json'), 'runs': runs(), 'trace_filter': TRACE,
            'sync_filter': 'fdatasync,fsync', 'expected_sync_syscall': 'fdatasync',
            'unexpected_sync_guard': 'On this Linux host fsync is traced so an unexpected replacement or extra sync fails closed.',
            'environment': {k: ENV[k] for k in ['LC_ALL', 'DEBUGINFOD_URLS']},
            'scope': 'Separate strace children including setup and oracles. Sync-only durations align after exactly one untimed file-store preflight and all warmups; no untraced latency claim.'}


def check_protocol():
    retained = read(ROOT / 'sync-protocol.json')
    if retained != protocol():
        raise ValueError('sync profile protocol binding changed')
    return retained


def parse_sync(text, replay_dir, samples, warmups):
    events = []
    for line in text.splitlines():
        if not line.strip():
            continue
        match = CALL.fullmatch(line.strip())
        if not match or match['call'] != 'fdatasync' or match['result'] != '0':
            raise ValueError(f'unexpected or failed sync event: {line}')
        path = Path(match['path'])
        if path.parent != replay_dir or path.suffix != '.replay':
            raise ValueError('sync event is not the exact caller replay file')
        if events and (decimal.Decimal(match['time']) < decimal.Decimal(events[-1]['timestamp']) or str(path) != events[0]['path']):
            raise ValueError('sync event order or replay pathname changed')
        events.append({'timestamp': match['time'], 'path': str(path),
                       'duration_ns': int(decimal.Decimal(match['duration']) * 1_000_000_000)})
    if len(events) != 1 + warmups + samples:
        raise ValueError('sync count differs from preflight + warmups + measured samples')
    return {'events': events, 'preflight_count': 1, 'warmups': warmups,
            'measured': events[1 + warmups:]}


def parse_summary(text, samples, warmups):
    rows = {}
    total = None
    header = False
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith('---'):
            continue
        if line.startswith('% time'):
            if header or rows:
                raise ValueError('duplicate or misplaced syscall summary header')
            header = True
            continue
        fields = line.split()
        if not header or len(fields) not in (5, 6):
            raise ValueError('invalid syscall summary row')
        percent, seconds = (decimal.Decimal(v) for v in fields[:2])
        usecs, calls = (int(v) for v in fields[2:4])
        errors = int(fields[4]) if len(fields) == 6 else 0
        name = fields[-1]
        if not percent.is_finite() or not seconds.is_finite() or not (0 <= percent <= 100) or seconds < 0 or usecs < 0 or calls < 0 or not (0 <= errors <= calls):
            raise ValueError('invalid syscall summary value')
        record = {'calls': calls, 'errors': errors, 'seconds': str(seconds), 'reported_percent': str(percent)}
        if name == 'total':
            if total is not None:
                raise ValueError('duplicate syscall summary total')
            total = record
        else:
            if name not in TRACE.split(',') or name in rows or total is not None:
                raise ValueError('unexpected syscall summary name/order')
            rows[name] = record
    if not rows or total is None or total['calls'] != sum(v['calls'] for v in rows.values()) or total['errors'] != sum(v['errors'] for v in rows.values()):
        raise ValueError('syscall summary totals or inventory missing')
    sync = rows.get('fdatasync', {})
    if sync.get('calls') != 1 + samples + warmups or sync.get('errors') != 0 or rows.get('fsync', {}).get('calls', 0):
        raise ValueError('summary sync count differs from the frozen lifecycle')
    for names in [('openat',), ('close',), ('read', 'pread64'), ('write', 'pwrite64'), ('unlink', 'unlinkat'), ('fstat', 'statx', 'newfstatat')]:
        if sum(rows.get(name, {}).get('calls', 0) for name in names) == 0:
            raise ValueError('summary lacks expected file lifecycle calls')
    return {'syscalls': rows, 'total': total,
            'scope': 'Whole traced child including preflight, warmups, measured operation, cleanup, setup and report output.'}


def argv_for(run, binary):
    directory = ROOT / 'sync-profiles' / run['label']
    replay = directory / 'replay'
    argv = routes._route_argv(binary, CASE, SPEC, samples=run['samples'], warmups=run['warmups'],
                             report=directory / 'report.json', resource=directory / 'resource.txt', replay_dir=replay)
    command = ['strace', '-f', '-qq', '-e', 'signal=none']
    if run['kind'] == 'sync':
        command += ['-ttt', '-T', '-yy', '-e', 'trace=fdatasync,fsync']
    else:
        command += ['-c', '-e', 'trace=' + TRACE]
    command += ['-o', str(directory / 'profile.txt'), '--', *argv]
    return directory, replay, argv, command


def validate_run(run, frozen):
    binary = frozen['builds'][run['phase']]['binary']
    directory, replay, argv, command = argv_for(run, binary)
    receipt = read(directory / 'receipt.json')
    if receipt['run'] != run or receipt['argv'] != argv or receipt['command'] != command or receipt['returncode'] != 0 or receipt.get('timed_out') is not False or receipt['protocol'] != meta(ROOT / 'sync-protocol.json'):
        raise ValueError('profile receipt identity or command failed')
    expected_names = {'started.json', 'report.json', 'resource.txt', 'profile.txt', 'stdout.txt', 'stderr.txt'}
    if set(receipt['artifacts']) != expected_names or set(p.name for p in directory.iterdir()) != expected_names | {'receipt.json'}:
        raise ValueError('profile artifact inventory mismatch')
    for name, digest in receipt['artifacts'].items():
        if meta(directory / name) != digest:
            raise ValueError('profile artifact changed')
    if any((directory / name).stat().st_size == 0 for name in ['report.json', 'resource.txt', 'profile.txt']):
        raise ValueError('empty profile evidence')
    report = routes.check_route_report(directory / 'report.json', 'normal', CASE, SPEC,
                                       samples=run['samples'], warmups=run['warmups'], binary=binary, argv=argv, replay_dir=replay)
    if receipt['replay_cleanup'] != {'path': str(replay), 'empty_before_removal': True, 'removed': True} or replay.exists():
        raise ValueError('replay cleanup missing')
    case = report['cases'][0]
    result = {'run': run, 'receipt': meta(directory / 'receipt.json'), 'source': case['source'], 'authored': case['authored'], 'oracle': case['oracle']}
    if run['kind'] == 'sync':
        parsed = parse_sync((directory / 'profile.txt').read_text(), replay, run['samples'], run['warmups'])
        result['sync'] = parsed
        result['paired_samples'] = []
        for sample, event in zip(case['samples'], parsed['measured'], strict=True):
            elapsed = sample['elapsed_ns']
            if event['duration_ns'] > elapsed:
                raise ValueError('sync duration exceeds containing measured operation')
            result['paired_samples'].append({'elapsed_ns': elapsed, 'fdatasync_ns': event['duration_ns'],
                                             'sync_fraction': event['duration_ns'] / elapsed})
    else:
        result['summary'] = parse_summary((directory / 'profile.txt').read_text(), run['samples'], run['warmups'])
    return result


def capture():
    frozen = check_protocol()
    if shutil.which('strace') is None:
        write(ROOT / 'sync-unavailable.json', {'status': 'unavailable', 'reason': 'strace not installed'})
        return
    with LOCK.open('a') as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        for run in frozen['runs']:
            binary = frozen['builds'][run['phase']]['binary']
            directory, replay, argv, command = argv_for(run, binary)
            directory.mkdir(parents=True, exist_ok=False)
            replay.mkdir()
            started = {'run': run, 'argv': argv, 'command': command, 'protocol': meta(ROOT / 'sync-protocol.json'), 'started_utc': now()}
            write(directory / 'started.json', started)
            timed_out = False
            with (directory / 'stdout.txt').open('xb') as out, (directory / 'stderr.txt').open('xb') as err:
                process = subprocess.Popen(command, cwd=REPO, env=ENV, stdout=out, stderr=err, start_new_session=True)
                try:
                    process.wait(timeout=180)
                except subprocess.TimeoutExpired:
                    timed_out = True
                    os.killpg(process.pid, signal.SIGKILL)
                    process.wait()
            cleanup = {'path': str(replay), 'empty_before_removal': not any(replay.iterdir()), 'removed': False}
            if process.returncode == 0 and cleanup['empty_before_removal']:
                replay.rmdir()
                cleanup['removed'] = True
            write(directory / 'receipt.json', dict(started, returncode=process.returncode, timed_out=timed_out, finished_utc=now(), replay_cleanup=cleanup,
                                                   artifacts={p.name: meta(p) for p in directory.iterdir() if p.is_file()}))
            validate_run(run, frozen)
            print(run['label'], 'pass', flush=True)
        source_check()


def analyze():
    frozen = check_protocol()
    root = ROOT / 'sync-profiles'
    if {p.name for p in root.iterdir()} != {r['label'] for r in runs()}:
        raise ValueError('sync profile child inventory mismatch')
    rows = [validate_run(run, frozen) for run in runs()]
    identity = [{k: row[k] for k in ['source', 'authored', 'oracle']} for row in rows]
    if any(value != identity[0] for value in identity):
        raise ValueError('profile source/authored/candidate identity changed')
    return {'schema': 'docx-replay-sync-profile-summary-v1', 'status': 'pass',
            'protocol': meta(ROOT / 'sync-protocol.json'), 'rows': rows,
            'scope': frozen['scope'], 'diagnostic_children': len(rows)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=['freeze', 'capture', 'analyze', 'verify'])
    args = parser.parse_args()
    if args.command == 'freeze':
        write(ROOT / 'sync-protocol.json', protocol())
    elif args.command == 'capture':
        capture()
    elif args.command == 'analyze':
        write(ROOT / 'sync-summary.json', analyze())
    elif read(ROOT / 'sync-summary.json') != analyze():
        raise ValueError('sync summary changed')
    print(args.command, 'pass')


if __name__ == '__main__':
    main()
