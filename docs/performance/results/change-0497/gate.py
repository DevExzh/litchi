#!/usr/bin/env python3
"""Retain one isolated validation attempt with source and terminal custody."""
import argparse
import os
from pathlib import Path
import signal
import subprocess
import time
from build import ROOT, TEMP, meta, source_manifest, write


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('label')
    parser.add_argument('argv', nargs=argparse.REMAINDER)
    args = parser.parse_args()
    if not args.label.replace('-', '').replace('_', '').isalnum() or not args.argv:
        raise ValueError('unique label and command required')
    source = TEMP / 'after'
    prefix = ROOT / 'validation' / args.label
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4', CARGO_INCREMENTAL='0',
               CARGO_PROFILE_RELEASE_DEBUG='1', RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
               CARGO_TARGET_DIR=str(TEMP / 'target'), TMPDIR=str(TEMP / 'tmp'), DEBUGINFOD_URLS='',
               LC_ALL='C', RUSTDOCFLAGS='-Dwarnings', PYTHONDONTWRITEBYTECODE='1')
    keys = ('RUSTUP_TOOLCHAIN', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_RELEASE_DEBUG',
            'RUSTFLAGS', 'CARGO_TARGET_DIR', 'TMPDIR', 'DEBUGINFOD_URLS', 'LC_ALL', 'RUSTDOCFLAGS', 'PYTHONDONTWRITEBYTECODE')
    before = source_manifest(source)
    manifest = prefix.with_suffix('.source.json')
    write(manifest, before)
    record = dict(argv=args.argv, cwd=str(source), environment={k: env[k] for k in keys},
                  driver=meta(__file__), source_manifest=meta(manifest), started_ns=time.time_ns(),
                  timeout_seconds=3600)
    write(prefix.with_suffix('.started.json'), record)
    with prefix.with_suffix('.stdout').open('xb') as out, prefix.with_suffix('.stderr').open('xb') as err:
        child = subprocess.Popen(args.argv, cwd=source, env=env, stdout=out, stderr=err, start_new_session=True)
        print(args.label, 'pid', child.pid, flush=True)
        record.update(pid=child.pid, timed_out=False, termination=None)
        try:
            child.wait(timeout=3600)
        except subprocess.TimeoutExpired:
            record.update(timed_out=True, termination='SIGTERM')
            os.killpg(child.pid, signal.SIGTERM)
            try:
                child.wait(timeout=5)
            except subprocess.TimeoutExpired:
                record['termination'] = 'SIGKILL'
                os.killpg(child.pid, signal.SIGKILL)
                child.wait()
    record.update(exit_code=child.returncode, finished_ns=time.time_ns(), source_unchanged=before == source_manifest(source),
                  stdout=meta(prefix.with_suffix('.stdout')), stderr=meta(prefix.with_suffix('.stderr')))
    write(prefix.with_suffix('.json'), record)
    print(args.label, 'exit', child.returncode, 'source_unchanged', record['source_unchanged'], flush=True)
    raise SystemExit(0 if child.returncode == 0 and record['source_unchanged'] and not record['timed_out'] else 1)


if __name__ == '__main__':
    main()
