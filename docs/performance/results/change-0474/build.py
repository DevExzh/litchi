#!/usr/bin/env python3
"""Build both streaming measurement targets from the authenticated clean source."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path('/tmp/litchi-goal-0474')
TREE = TEMP / 'tree'
ENV = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4', CARGO_INCREMENTAL='0',
           CARGO_PROFILE_RELEASE_DEBUG='1', RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
           DEBUGINFOD_URLS='', LC_ALL='C', PYTHONDONTWRITEBYTECODE='1')

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def write(path, value):
    with path.open('x') as stream:
        json.dump(value, stream, indent=2, sort_keys=True); stream.write('\n')

def check():
    source = json.loads((ROOT / 'source-binding.json').read_text())
    assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=TREE, text=True).strip() == source['revision']
    assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=TREE).strip()
    assert sha(ROOT / source['source_manifest']['path']) == source['source_manifest']['sha256']
    for path, digest in json.loads((ROOT / source['source_manifest']['path']).read_text()).items():
        assert sha(TREE / path) == digest, path
    for path, digest in source['fixtures'].items():
        assert sha(TREE / path) == digest, path
    return source

def main():
    source = check()
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert sha(Path(__file__)) == protocol['build_driver_sha256']
    argv = ['cargo', 'build', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
            '--target-dir', str(REPO / 'tools/perf-baseline/target'), '--features', 'allocator-metrics',
            '--bin', 'litchi-perf-baseline', '--bin', 'litchi-perf-baseline-alloc']
    record = dict(argv=argv, cwd=str(TREE), driver_sha256=sha(Path(__file__)), protocol_sha256=sha(ROOT / 'protocol.json'),
                  source_binding_sha256=sha(ROOT / 'source-binding.json'),
                  environment={k: ENV[k] for k in ['RUSTUP_TOOLCHAIN', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_RELEASE_DEBUG', 'RUSTFLAGS', 'DEBUGINFOD_URLS', 'LC_ALL']},
                  started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    write(ROOT / 'build-command.started.json', record)
    with (ROOT / 'build-command.stdout').open('xb') as out, (ROOT / 'build-command.stderr').open('xb') as err:
        result = subprocess.run(argv, cwd=TREE, env=ENV, stdout=out, stderr=err)
    record.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                  artifacts={name: dict(sha256=sha(ROOT / name), bytes=(ROOT / name).stat().st_size)
                             for name in ['build-command.stdout', 'build-command.stderr']})
    write(ROOT / 'build-command.json', record)
    if result.returncode:
        raise SystemExit(result.returncode)
    check()
    binaries = {}
    for mode, name in [('normal', 'litchi-perf-baseline'), ('allocator', 'litchi-perf-baseline-alloc')]:
        target = TEMP / mode
        assert not target.exists()
        shutil.copy2(REPO / 'tools/perf-baseline/target/release' / name, target)
        binaries[mode] = dict(path=str(target), sha256=sha(target), bytes=target.stat().st_size)
    write(ROOT / 'build.json', dict(schema='litchi-0474-build-v1', **source, build_path=str(TREE), binaries=binaries,
          receipt=dict(path='build-command.json', sha256=sha(ROOT / 'build-command.json')),
          protocol_sha256=sha(ROOT / 'protocol.json')))
    print('build pass', flush=True)

if __name__ == '__main__':
    main()
