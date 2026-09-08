#!/usr/bin/env python3
"""Matched fixed-path captures using binaries built in the shared checkout."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parent
TEMP = Path('/tmp/litchi-goal-0467')
SHARED_TREE = TEMP / 'shared-tree'
FLAGS = '-C force-frame-pointers=yes -C force-unwind-tables=yes'
LANES = ('A1-fixed', 'B1-fixed', 'B2-fixed', 'A2-fixed', 'A-full-fixed', 'B-full-fixed')
ROLE_FOR = {
    'A1': 'control', 'A2': 'control', 'B1': 'candidate', 'B2': 'candidate',
    'A-full': 'control', 'B-full': 'candidate',
}


def sha(path):
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def git(*args):
    return subprocess.check_output(['git', *args], cwd=SHARED_TREE, text=True).strip()


def require_clean(expected_revision=None):
    if expected_revision is not None and git('rev-parse', 'HEAD') != expected_revision:
        raise RuntimeError('shared tree is at the wrong revision')
    if git('status', '--porcelain'):
        raise RuntimeError('shared tree is not clean')


def checkout_revision(revision):
    require_clean()
    verify_argv = ['git', 'cat-file', '-e', revision + '^{commit}']
    subprocess.run(verify_argv, cwd=SHARED_TREE, check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
    checkout_argv = ['git', 'checkout', '--detach', revision]
    subprocess.run(checkout_argv, cwd=SHARED_TREE, check=True)
    require_clean(revision)
    return verify_argv, checkout_argv


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in LANES:
        raise SystemExit('usage: capture-fixed.py A1-fixed|B1-fixed|B2-fixed|A2-fixed|A-full-fixed|B-full-fixed')
    lane = sys.argv[1]
    phase = lane.removesuffix('-fixed')
    role = ROLE_FOR[phase]
    binding_path = ROOT / (role + '-fixed-binding.json')
    binding = json.loads(binding_path.read_text())
    receipt_path = ROOT / binding.get('build_receipt', '')
    if (binding.get('role') != role or not binding.get('clean_build')
            or not receipt_path.is_file()
            or sha(receipt_path) != binding.get('build_receipt_sha256')):
        raise RuntimeError('fixed binding is not a clean role binding')
    binary = TEMP / (role + '-fixed')
    if (not binary.is_file() or binary.stat().st_size != binding.get('bytes')
            or sha(binary) != binding['binary_sha256']):
        raise RuntimeError('fixed binary does not match its binding')
    revision = binding['revision']
    verify_argv, checkout_argv = checkout_revision(revision)
    output = ROOT / lane
    output.mkdir(exist_ok=False)
    full = phase.endswith('-full')
    samples, warmups = (15, 3) if full else (500, 5)
    workload = [
        str(binary), '--workers', '1', '--warmup', str(warmups), '--samples', str(samples),
        '--json', str(output / 'report.json'),
        '--corpus-manifest', str(output / 'corpus-catalog.json'),
    ]
    if not full:
        workload += [
            '--case', 'doc_fresh_write_to,xlsx_one_percent_commit_save',
            '--xlsx-shape', 'dense-wide',
        ]
    argv = ['taskset', '-c', '2', '/usr/bin/time', '-v', '-o', str(output / 'resource.log')]
    argv += workload
    env = dict(
        os.environ,
        RUSTUP_TOOLCHAIN='1.98.1',
        RUSTFLAGS=FLAGS,
        CARGO_PROFILE_RELEASE_DEBUG='1',
        CARGO_INCREMENTAL='0',
        CARGO_BUILD_JOBS='4',
        DEBUGINFOD_URLS='',
        LC_ALL='C',
        PYTHONDONTWRITEBYTECODE='1',
    )
    receipt = dict(
        schema='litchi-0467-capture-fixed-v1',
        lane=lane,
        role=role,
        revision=revision,
        binary_sha256=sha(binary),
        binding_sha256=sha(binding_path),
        driver_sha256=sha(Path(__file__)),
        checkout_verify_argv=verify_argv,
        checkout_argv=checkout_argv,
        argv=argv,
        cwd=str(SHARED_TREE),
        samples=samples,
        warmups=warmups,
        cases=None if full else ['doc_fresh_write_to', 'xlsx_one_percent_commit_save'],
        xlsx_shape=None if full else 'dense-wide',
        environment={key: env[key] for key in (
            'RUSTUP_TOOLCHAIN', 'RUSTFLAGS', 'CARGO_PROFILE_RELEASE_DEBUG',
            'CARGO_INCREMENTAL', 'CARGO_BUILD_JOBS', 'DEBUGINFOD_URLS', 'LC_ALL',
        )},
        started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        clean_before=True,
    )
    (output / 'started.json').write_text(json.dumps(receipt, indent=2) + '\n')
    with (output / 'stdout.log').open('x') as stdout, (output / 'stderr.log').open('x') as stderr:
        result = subprocess.run(argv, cwd=SHARED_TREE, env=env, stdout=stdout, stderr=stderr)
    require_clean(revision)
    binary_unchanged = sha(binary) == binding['binary_sha256']
    report_ok = False
    report_path = output / 'report.json'
    if result.returncode == 0 and report_path.is_file():
        report = json.loads(report_path.read_text())
        report_ok = (
            report['environment']['git_revision'] == revision
            and report['environment']['git_worktree_dirty'] is False
        )
    receipt.update(
        exit_code=result.returncode,
        clean_after=True,
        binary_unchanged=binary_unchanged,
        report_metadata_matches_clean_role=report_ok,
        finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        artifacts={
            path.name: {'sha256': sha(path), 'bytes': path.stat().st_size}
            for path in output.iterdir() if path.is_file()
        },
    )
    (output / 'receipt.json').write_text(json.dumps(receipt, indent=2) + '\n')
    print(json.dumps(dict(lane=lane, exit_code=result.returncode)), flush=True)
    if result.returncode:
        raise SystemExit(result.returncode)
    if not binary_unchanged or not report_ok:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
