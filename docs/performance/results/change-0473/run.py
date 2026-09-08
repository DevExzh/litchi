#!/usr/bin/env python3
"""Two schema pilots followed by the frozen formal matrix, under cpu.lock."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent

def main():
    build = json.loads((ROOT / 'build.json').read_text())
    tree = Path(build['build_path'])
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
               CARGO_PROFILE_RELEASE_DEBUG='1', DEBUGINFOD_URLS='', LC_ALL='C')
    for mode in ['normal', 'allocator']:
        directory = ROOT / 'pilots' / mode
        directory.mkdir(parents=True, exist_ok=False)
        argv = ['taskset', '-c', '2', build['binaries'][mode]['path'], '--case', 'docx_streaming_create',
                '--semantic-shape', 'tiny', '--workers', '1', '--warmup', '0', '--samples', '1',
                '--json', str(directory / 'report.json'), '--corpus-manifest', str(directory / 'corpus-catalog.json')]
        record = dict(argv=argv, cwd=str(tree), started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                      driver_sha256=hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
                      environment={k: env[k] for k in ['RUSTUP_TOOLCHAIN', 'RUSTFLAGS', 'CARGO_PROFILE_RELEASE_DEBUG', 'DEBUGINFOD_URLS', 'LC_ALL']})
        with (directory / 'stdout.log').open('xb') as out, (directory / 'stderr.log').open('xb') as err:
            result = subprocess.run(argv, cwd=tree, env=env, stdout=out, stderr=err)
        record.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                      artifacts={p.name: dict(sha256=hashlib.sha256(p.read_bytes()).hexdigest(), bytes=p.stat().st_size)
                                 for p in directory.iterdir() if p.is_file()})
        (directory / 'receipt.json').write_text(json.dumps(record, indent=2)+'\n')
        print('pilot', mode, result.returncode, flush=True)
        if result.returncode:
            raise SystemExit(result.returncode)
    for lane in json.loads((ROOT / 'protocol.json').read_text())['order']:
        subprocess.run(['python3', '-B', str(ROOT / 'capture.py'), lane['lane']], check=True)

if __name__ == '__main__':
    main()
