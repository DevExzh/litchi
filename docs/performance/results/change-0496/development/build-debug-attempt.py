#!/usr/bin/env python3
"""Build isolated before/after phase-diagnostic harnesses with source custody."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time

ROOT = Path(__file__).resolve().parent
TEMP = Path('/home/zhuhe/.cache/litchi-goal-0496')
REPO = ROOT.parents[3]


def meta(path):
    path = Path(path)
    with path.open('rb') as stream:
        digest = hashlib.file_digest(stream, 'sha256').hexdigest()
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': digest}


def write(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x') as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write('\n')


def source_manifest(source):
    names = subprocess.check_output(['git', 'ls-files', '-co', '--exclude-standard', '-z'], cwd=source).decode().split('\0')
    names.append('Cargo.lock')
    return {name: meta(source / name) for name in sorted(set(names))
            if name and (source / name).is_file() and (Path(name).suffix in ('.rs', '.toml', '.lock', '.xml') or name.startswith('.cargo/'))}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('phase', choices=('before', 'after'))
    parser.add_argument('--attempt', required=True)
    args = parser.parse_args()
    if not args.attempt.replace('-', '').isalnum():
        raise ValueError('invalid attempt')
    source = TEMP / args.phase
    target = TEMP / 'target'
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4', CARGO_INCREMENTAL='0',
               CARGO_PROFILE_RELEASE_DEBUG='1', RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
               CARGO_TARGET_DIR=str(target), TMPDIR=str(TEMP / 'tmp'), DEBUGINFOD_URLS='', LC_ALL='C',
               RUSTDOCFLAGS='-Dwarnings')
    env_keys = ('RUSTUP_TOOLCHAIN', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_RELEASE_DEBUG',
                'RUSTFLAGS', 'CARGO_TARGET_DIR', 'TMPDIR', 'DEBUGINFOD_URLS', 'LC_ALL', 'RUSTDOCFLAGS')
    manifest = source_manifest(source)
    manifest_path = ROOT / 'builds' / f'{args.phase}-{args.attempt}-source.json'
    write(manifest_path, manifest)
    common = ['--locked', '--offline', '--manifest-path', 'tools/perf-baseline/Cargo.toml']
    jobs = [
        ('normal', ['cargo', 'build', '--release', *common, '--bin', 'litchi-perf-baseline']),
        ('allocator', ['cargo', 'build', '--release', *common, '--features', 'allocator-metrics', '--bin', 'litchi-perf-baseline-alloc']),
    ]
    if args.phase == 'after':
        jobs += [
            ('test', ['cargo', 'test', *common, '--features', 'allocator-metrics', '--lib', '--', '--test-threads=1']),
            ('clippy', ['cargo', 'clippy', *common, '--features', 'allocator-metrics', '--all-targets', '--', '-D', 'warnings']),
            ('rustdoc', ['cargo', 'doc', *common, '--features', 'allocator-metrics', '--no-deps', '--lib']),
        ]
    builds = {}
    for name, argv in jobs:
        gate = ROOT / 'builds' / f'{args.phase}-{args.attempt}-{name}'
        receipt = {'argv': argv, 'cwd': str(source), 'environment': {k: env[k] for k in env_keys},
                   'source_manifest': meta(manifest_path), 'started_ns': time.time_ns(), 'driver': meta(__file__)}
        write(gate.with_suffix('.started.json'), receipt)
        with gate.with_suffix('.stdout').open('xb') as stdout, gate.with_suffix('.stderr').open('xb') as stderr:
            child = subprocess.Popen(argv, cwd=source, env=env, stdout=stdout, stderr=stderr, start_new_session=True)
            print(name, 'pid', child.pid, flush=True)
            receipt['pid'] = child.pid
            code = child.wait()
        unchanged = manifest == source_manifest(source)
        receipt.update(exit_code=code, finished_ns=time.time_ns(), source_unchanged=unchanged,
                       stdout=meta(gate.with_suffix('.stdout')), stderr=meta(gate.with_suffix('.stderr')))
        write(gate.with_suffix('.json'), receipt)
        if code or not unchanged:
            raise RuntimeError(f'{name} failed: exit={code}, unchanged={unchanged}')
        if name in ('normal', 'allocator'):
            binary_name = 'litchi-perf-baseline' + ('-alloc' if name == 'allocator' else '')
            retained = TEMP / 'retained' / args.phase / name / binary_name
            retained.parent.mkdir(parents=True, exist_ok=True)
            if retained.exists():
                raise RuntimeError('retained binary already exists')
            shutil.copy2(target / 'release' / binary_name, retained)
            builds[args.phase + '/' + name] = {'binary': meta(retained), 'git_revision': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=source, text=True).strip(), 'source_manifest': meta(manifest_path), 'gate': meta(gate.with_suffix('.json'))}
        print(name, 'passed', flush=True)
    write(ROOT / f'build-{args.phase}-{args.attempt}.json', builds)


if __name__ == '__main__':
    main()
