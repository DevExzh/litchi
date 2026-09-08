#!/usr/bin/env python3
"""Build both comparison roles from one authenticated shared checkout path."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys


ROOT = Path(__file__).resolve().parent
TEMP = Path('/tmp/litchi-goal-0467')
SHARED_TREE = TEMP / 'shared-tree'
FLAGS = '-C force-frame-pointers=yes -C force-unwind-tables=yes'
TARGET = ROOT.parents[3] / 'tools/perf-baseline/target'
SOURCE_BINDINGS = ROOT / 'source-bindings.json'
ROLES = ('control', 'candidate')


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


def authenticated_binding(role):
    path = ROOT / (role + '-binding.json')
    binding = json.loads(path.read_text())
    revision = binding['revision']
    if not re.fullmatch(r'[0-9a-f]{40}', revision):
        raise RuntimeError('original binding has a non-commit revision')
    receipt_path = ROOT / binding['build_receipt']
    if sha(receipt_path) != binding['build_receipt_sha256']:
        raise RuntimeError('original build receipt does not match its binding')
    receipt = json.loads(receipt_path.read_text())
    if (receipt.get('role') != role or receipt.get('revision') != revision
            or receipt.get('exit_code') != 0 or not receipt.get('clean_before')
            or not receipt.get('clean_after') or not binding.get('clean_build')):
        raise RuntimeError('original build receipt is not an authenticated clean build')
    return binding, receipt_path


def checkout_revision(revision):
    if not SHARED_TREE.is_dir():
        raise RuntimeError(f'missing shared checkout: {SHARED_TREE}')
    require_clean()
    verify_argv = ['git', 'cat-file', '-e', revision + '^{commit}']
    subprocess.run(verify_argv, cwd=SHARED_TREE, check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
    checkout_argv = ['git', 'checkout', '--detach', revision]
    subprocess.run(checkout_argv, cwd=SHARED_TREE, check=True)
    require_clean(revision)
    return verify_argv, checkout_argv


def verify_source_manifest(role, revision):
    source_bindings = json.loads(SOURCE_BINDINGS.read_text())
    role_info = source_bindings['roles'][role]
    if role_info['revision'] != revision:
        raise RuntimeError('source binding revision does not match role binding')
    tree = git('rev-parse', 'HEAD^{tree}')
    if tree != role_info['git_tree']:
        raise RuntimeError('shared tree does not match the authenticated source tree')
    manifest_path = ROOT / role_info['path']
    if sha(manifest_path) != role_info['sha256']:
        raise RuntimeError('source manifest does not match source-bindings.json')
    manifest = json.loads(manifest_path.read_text())
    if len(manifest) != role_info['files']:
        raise RuntimeError('source manifest file count changed')
    for relative, expected in manifest.items():
        path = SHARED_TREE / relative
        if not path.is_file() or sha(path) != expected:
            raise RuntimeError(f'source manifest mismatch: {relative}')
    fixtures = role_info['included_fixtures']
    for relative, expected in fixtures.items():
        path = SHARED_TREE / relative
        if not path.is_file() or sha(path) != expected:
            raise RuntimeError(f'included fixture mismatch: {relative}')
    return dict(
        source_manifest=manifest_path.name,
        source_manifest_sha256=role_info['sha256'],
        source_manifest_files=len(manifest),
        included_fixtures=len(fixtures),
    )


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in ROLES:
        raise SystemExit('usage: build-fixed.py control|candidate')
    role = sys.argv[1]
    binding, original_receipt_path = authenticated_binding(role)
    revision = binding['revision']
    verify_argv, checkout_argv = checkout_revision(revision)
    source_proof = verify_source_manifest(role, revision)
    env = dict(
        os.environ,
        RUSTUP_TOOLCHAIN='1.98.1',
        RUSTFLAGS=FLAGS,
        CARGO_PROFILE_RELEASE_DEBUG='1',
        CARGO_INCREMENTAL='0',
        CARGO_BUILD_JOBS='4',
        DEBUGINFOD_URLS='',
        LC_ALL='C',
    )
    argv = [
        'cargo', 'build', '--release', '--locked',
        '--manifest-path', 'tools/perf-baseline/Cargo.toml',
        '--target-dir', str(TARGET), '--bin', 'litchi-perf-baseline',
    ]
    receipt = dict(
        schema='litchi-0467-fixed-build-v1',
        role=role,
        revision=revision,
        original_binding=role + '-binding.json',
        original_binding_sha256=sha(ROOT / (role + '-binding.json')),
        original_build_receipt=str(original_receipt_path.name),
        original_build_receipt_sha256=sha(original_receipt_path),
        checkout_verify_argv=verify_argv,
        checkout_argv=checkout_argv,
        **source_proof,
        source_manifest_verified_before_build=True,
        argv=argv,
        cwd=str(SHARED_TREE),
        clean_before=True,
        environment={key: env[key] for key in (
            'RUSTUP_TOOLCHAIN', 'RUSTFLAGS', 'CARGO_PROFILE_RELEASE_DEBUG',
            'CARGO_INCREMENTAL', 'CARGO_BUILD_JOBS', 'DEBUGINFOD_URLS',
        )},
        started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
    )
    stdout_path = ROOT / (role + '-fixed-build.stdout')
    stderr_path = ROOT / (role + '-fixed-build.stderr')
    receipt_path = ROOT / (role + '-fixed-build.json')
    binding_path = ROOT / (role + '-fixed-binding.json')
    for path in (stdout_path, stderr_path, receipt_path, binding_path, TEMP / (role + '-fixed')):
        if path.exists():
            raise RuntimeError(f'refusing to overwrite existing artifact: {path}')
    with stdout_path.open('x') as stdout, stderr_path.open('x') as stderr:
        result = subprocess.run(argv, cwd=SHARED_TREE, env=env, stdout=stdout, stderr=stderr)
    require_clean(revision)
    source_proof_after = verify_source_manifest(role, revision)
    if source_proof_after != source_proof:
        raise RuntimeError('source manifest proof changed during build')
    receipt.update(
        exit_code=result.returncode,
        clean_after=True,
        source_manifest_verified_after_build=True,
        finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
    )
    receipt_path.write_text(json.dumps(receipt, indent=2) + '\n')
    if result.returncode:
        raise SystemExit(result.returncode)
    built = TARGET / 'release/litchi-perf-baseline'
    binary = TEMP / (role + '-fixed')
    shutil.copy2(built, binary)
    fixed_binding = dict(
        role=role,
        revision=revision,
        binary_sha256=sha(binary),
        bytes=binary.stat().st_size,
        build_receipt=role + '-fixed-build.json',
        build_receipt_sha256=sha(receipt_path),
        clean_build=True,
    )
    binding_path.write_text(
        json.dumps(fixed_binding, indent=2) + '\n'
    )
    print(json.dumps(fixed_binding), flush=True)


if __name__ == '__main__':
    main()
