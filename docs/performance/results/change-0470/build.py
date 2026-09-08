#!/usr/bin/env python3
"""Build the candidate from the exact absolute checkout path used by the control."""
import datetime
import gzip
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path('/tmp/litchi-goal-0470')
TREE = Path('/tmp/litchi-goal-0468/profile-tree')
BINARY = TEMP / 'candidate'
ENV = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4',
           CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_DEBUG='1',
           RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
           DEBUGINFOD_URLS='', LC_ALL='C', PYTHONDONTWRITEBYTECODE='1')


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def write(path, value):
    with path.open('x') as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write('\n')


def check_source():
    binding = json.loads((ROOT / 'candidate-source-binding.json').read_text())
    revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=TREE, text=True).strip()
    assert revision == binding['revision']
    assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=TREE)
    assert sha(ROOT / 'sources/candidate.json') == binding['source_manifest_sha256']
    for name, digest in json.loads((ROOT / 'sources/candidate.json').read_text()).items():
        assert sha(TREE / name) == digest, name
    for name, digest in binding['included_fixtures'].items():
        assert sha(TREE / name) == digest, name
    return binding


def run(argv, prefix):
    receipt = dict(schema='litchi-0470-command-v1', argv=argv, cwd=str(TREE),
                   driver_sha256=sha(Path(__file__)),
                   environment={k: ENV[k] for k in ('RUSTUP_TOOLCHAIN', 'RUSTFLAGS',
                       'CARGO_PROFILE_RELEASE_DEBUG', 'CARGO_BUILD_JOBS',
                       'CARGO_INCREMENTAL', 'DEBUGINFOD_URLS', 'LC_ALL')},
                   started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    write(prefix.with_suffix('.started.json'), receipt)
    with prefix.with_suffix('.stdout').open('xb') as out, prefix.with_suffix('.stderr').open('xb') as err:
        p = subprocess.run(argv, cwd=TREE, env=ENV, stdout=out, stderr=err)
    receipt.update(exit_code=p.returncode,
                   finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    receipt['outputs'] = {suffix: dict(sha256=sha(prefix.with_suffix(suffix)),
                         bytes=prefix.with_suffix(suffix).stat().st_size)
                         for suffix in ('.stdout', '.stderr')}
    write(prefix.with_suffix('.json'), receipt)
    print(prefix.name, p.returncode, flush=True)
    if p.returncode:
        raise SystemExit(p.returncode)


def build():
    source = check_source()
    run(['cargo', 'build', '--release', '--locked', '--manifest-path',
         'tools/perf-baseline/Cargo.toml', '--target-dir',
         str(REPO / 'tools/perf-baseline/target'), '--bin', 'litchi-perf-baseline'], ROOT / 'build')
    check_source()
    assert not BINARY.exists()
    shutil.copy2(REPO / 'tools/perf-baseline/target/release/litchi-perf-baseline', BINARY)
    write(ROOT / 'candidate-binding.json', dict(schema='litchi-0470-role-binding-v1',
          role='candidate', revision=source['revision'], clean_build=True,
          binary_sha256=sha(BINARY), bytes=BINARY.stat().st_size, binary_path=str(BINARY),
          build_path=str(TREE), build_receipt_path='build.json',
          build_receipt_sha256=sha(ROOT / 'build.json'),
          source_manifest='sources/candidate.json',
          source_manifest_sha256=source['source_manifest_sha256'],
          included_fixtures=source['included_fixtures'],
          source_binding_sha256=sha(ROOT / 'candidate-source-binding.json')))


if __name__ == '__main__':
    build()
