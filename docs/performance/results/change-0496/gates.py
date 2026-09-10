#!/usr/bin/env python3
"""Run the established release-profile full harness gates for 0496."""
import json
import os
from pathlib import Path
import subprocess
import time
from build import ROOT, TEMP, meta, source_manifest, write

source = TEMP / 'after'
env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4', CARGO_INCREMENTAL='0',
           CARGO_PROFILE_RELEASE_DEBUG='1', RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
           CARGO_TARGET_DIR=str(TEMP / 'target'), TMPDIR=str(TEMP / 'tmp'), DEBUGINFOD_URLS='', LC_ALL='C',
           RUSTDOCFLAGS='-Dwarnings')
keys = ('RUSTUP_TOOLCHAIN', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_RELEASE_DEBUG',
        'RUSTFLAGS', 'CARGO_TARGET_DIR', 'TMPDIR', 'DEBUGINFOD_URLS', 'LC_ALL', 'RUSTDOCFLAGS')
manifest_path = ROOT / 'builds/after-final1-source.json'
manifest = json.loads(manifest_path.read_text())
assert source_manifest(source) == manifest
common = ['--release', '--locked', '--offline', '--manifest-path', 'tools/perf-baseline/Cargo.toml']
jobs = [
    ('test', ['cargo', 'test', *common, '--lib', '--features', 'allocator-metrics', '--', '--test-threads=1']),
    ('clippy', ['cargo', 'clippy', *common, '--all-targets', '--features', 'allocator-metrics', '--', '-D', 'warnings']),
    ('rustdoc', ['cargo', 'doc', *common, '--lib', '--features', 'allocator-metrics', '--no-deps']),
]
for name, argv in jobs:
    gate = ROOT / 'builds' / ('after-final2-' + name)
    receipt = {'argv': argv, 'cwd': str(source), 'environment': {k: env[k] for k in keys},
               'source_manifest': meta(manifest_path), 'started_ns': time.time_ns(),
               'driver': meta(__file__), 'support': meta(ROOT / 'build.py')}
    write(gate.with_suffix('.started.json'), receipt)
    with gate.with_suffix('.stdout').open('xb') as stdout, gate.with_suffix('.stderr').open('xb') as stderr:
        child = subprocess.Popen(argv, cwd=source, env=env, stdout=stdout, stderr=stderr, start_new_session=True)
        receipt['pid'] = child.pid
        print(name, 'pid', child.pid, flush=True)
        code = child.wait()
    unchanged = source_manifest(source) == manifest
    receipt.update(exit_code=code, finished_ns=time.time_ns(), source_unchanged=unchanged,
                   stdout=meta(gate.with_suffix('.stdout')), stderr=meta(gate.with_suffix('.stderr')))
    write(gate.with_suffix('.json'), receipt)
    if code or not unchanged:
        raise RuntimeError(f'{name} failed: exit={code}, unchanged={unchanged}')
    print(name, 'passed', flush=True)
