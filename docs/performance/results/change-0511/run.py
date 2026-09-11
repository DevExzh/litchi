#!/usr/bin/env python3
"""Run one exclusive 0511 build, test, or descriptive measurement lane."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SCRATCH = Path('/tmp/litchi-goal-0511')
BINARY = SCRATCH / 'target/release/litchi-perf-baseline'

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def sources():
    paths = subprocess.check_output(['git', 'ls-files', '--cached', '--others', '--exclude-standard', '-z', 'crates', 'tools/perf-baseline', 'Cargo.toml', 'Cargo.lock', 'rust-toolchain.toml', '.cargo'], cwd=REPO).decode().split('\0')
    return {p: sha(REPO / p) for p in sorted(paths) if p and Path(p).suffix in {'.rs', '.toml', '.lock'}}

def write(path, obj):
    with path.open('x') as stream:
        json.dump(obj, stream, indent=2)
        stream.write('\n')

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('lane', choices=['build', 'preflight', 'r1', 'r2', 'tests', 'clippy', 'rustdoc', 'doctests', 'check-features'])
    args = parser.parse_args()
    env = dict(os.environ, CARGO_TARGET_DIR=str(SCRATCH / 'target'), CARGO_BUILD_JOBS='2', CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_DEBUG='0', CARGO_PROFILE_DEV_DEBUG='0', CARGO_PROFILE_TEST_DEBUG='0', TMPDIR=str(SCRATCH))
    common = ['--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml']
    if args.lane == 'build':
        command = ['cargo', 'build', '--release', '--bin', 'litchi-perf-baseline', *common]
    elif args.lane == 'tests':
        command = ['cargo', 'test', '--release', '--lib', *common]
    elif args.lane == 'clippy':
        command = ['cargo', 'clippy', '--release', '--lib', '--bin', 'litchi-perf-baseline', *common, '--', '-D', 'warnings']
    elif args.lane == 'rustdoc':
        env['RUSTDOCFLAGS'] = '-D warnings'
        command = ['cargo', 'doc', '--release', '--lib', '--no-deps', *common]
    elif args.lane == 'doctests':
        command = ['cargo', 'test', '--release', '--doc', *common]
    elif args.lane == 'check-features':
        command = ['cargo', 'check', '--release', '--all-targets', '--features', 'allocator-metrics', *common]
    else:
        warmup, samples = (0, 1) if args.lane == 'preflight' else (3, 15)
        command = [str(BINARY), '--warmup', str(warmup), '--samples', str(samples), '--json', str(HERE / f'{args.lane}-report.json'), '--corpus-manifest', str(HERE / f'{args.lane}-catalog.json')]
    receipt = HERE / f'{args.lane}-receipt.json'
    assert not receipt.exists()
    before = sources()
    if args.lane == 'build':
        write(HERE / 'source-manifest.json', before)
    else:
        assert before == json.loads((HERE / 'source-manifest.json').read_text()), 'source changed after build'
    measured = args.lane in {'preflight', 'r1', 'r2'}
    if measured:
        assert sha(BINARY) == json.loads((HERE / 'build-receipt.json').read_text())['binary_sha256']
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    tick = time.monotonic()
    with (HERE / f'{args.lane}.log').open('x') as stream:
        child = subprocess.run(['/usr/bin/time', '-v', *command], cwd=REPO, env=env, stdout=stream, stderr=subprocess.STDOUT)
    unchanged = sources() == before
    result = {'lane': args.lane, 'command': command, 'environment': {k: env[k] for k in ('CARGO_TARGET_DIR', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_RELEASE_DEBUG', 'CARGO_PROFILE_DEV_DEBUG', 'CARGO_PROFILE_TEST_DEBUG', 'TMPDIR')}, 'started_utc': started, 'elapsed_seconds': time.monotonic()-tick, 'exit_code': child.returncode, 'source_unchanged': unchanged, 'source_manifest_sha256': sha(HERE / 'source-manifest.json'), 'log_sha256': sha(HERE / f'{args.lane}.log')}
    if args.lane == 'build' and child.returncode == 0 or measured:
        result['binary_sha256'] = sha(BINARY)
    if args.lane == 'rustdoc':
        result['environment']['RUSTDOCFLAGS'] = env['RUSTDOCFLAGS']
    write(receipt, result)
    print(json.dumps(result))
    assert child.returncode == 0 and unchanged

if __name__ == '__main__':
    main()
