"""Exclusive, source-bound serial Rust check attempts before capture freeze."""

import argparse
import json
import os
from pathlib import Path
import subprocess
import time

import run as R


def source_manifest():
    names = subprocess.check_output(
        ['git', 'ls-files', '-z', 'crates', 'tools/perf-baseline', 'Cargo.toml',
         '.cargo', 'rust-toolchain.toml'], cwd=R.REPO).split(b'\0')
    untracked = subprocess.check_output(
        ['git', 'ls-files', '--others', '--exclude-standard', '-z', '--',
         'crates', 'tools/perf-baseline'], cwd=R.REPO).split(b'\0')
    paths = {p.decode() for p in names if p} | {
        p.decode() for p in untracked if p.endswith(b'.rs')}
    paths.add('Cargo.lock')
    return {name: R.sha(R.REPO / name) for name in sorted(paths)
            if (R.REPO / name).is_file()}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('name')
    parser.add_argument('command', nargs=argparse.REMAINDER)
    args = parser.parse_args()
    assert args.name and '/' not in args.name and args.name not in ('.', '..')
    command = args.command
    if command and command[0] == '--':
        command = command[1:]
    assert command
    folder = R.HERE / 'check-attempts' / args.name
    folder.mkdir(parents=True, exist_ok=False)
    manifest = source_manifest()
    R.write(folder / 'source-manifest.json', manifest)
    (folder / 'tracked-source.patch').write_bytes(subprocess.check_output(
        ['git', 'diff', '--', 'crates', 'tools/perf-baseline'], cwd=R.REPO))
    binding = json.loads((R.HERE / 'workspace-lock.json').read_text())
    assert manifest['Cargo.lock'] == binding['sha256']
    environment = os.environ.copy()
    environment.update(TMPDIR=str(R.TARGET / 'tmp'), CARGO_TARGET_DIR=str(R.TARGET),
                       CARGO_BUILD_JOBS='2', CARGO_INCREMENTAL='0',
                       RUSTDOCFLAGS='-D warnings')
    start, tick = R.now(), time.monotonic()
    script_hash = R.sha(Path(__file__))
    with (folder / 'stdout').open('x') as out, (folder / 'stderr').open('x') as err:
        child = subprocess.run(command, cwd=R.REPO, env=environment, stdout=out, stderr=err)
    stable = source_manifest() == manifest and R.sha(Path(__file__)) == script_hash
    R.write(folder / 'receipt.json', {
        'command': command, 'cwd': str(R.REPO), 'start_utc': start, 'end_utc': R.now(),
        'seconds': time.monotonic() - tick, 'exit_code': child.returncode,
        'source_stable': stable, 'source_manifest_sha256': R.sha(folder / 'source-manifest.json'),
        'script_sha256': script_hash, 'run_sha256': R.sha(R.HERE / 'run.py'),
        'plan_sha256': R.sha(R.HERE / 'plan.json'),
        'environment': {key: environment.get(key) for key in
                        ('TMPDIR', 'CARGO_TARGET_DIR', 'CARGO_BUILD_JOBS',
                         'CARGO_INCREMENTAL', 'RUSTDOCFLAGS', 'RUSTFLAGS',
                         'CARGO_ENCODED_RUSTFLAGS', 'LD_PRELOAD')},
        'artifacts': {path.name: R.sha(path) for path in sorted(folder.iterdir())
                      if path.is_file()},
    })
    assert stable, 'source changed during check'
    print(args.name, 'exit', child.returncode, flush=True)
    raise SystemExit(child.returncode)


if __name__ == '__main__':
    main()
