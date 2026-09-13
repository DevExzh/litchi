#!/usr/bin/env python3
"""Build and capture the already-prepared isolated source snapshot, serially."""
import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import runpy
import shutil
import subprocess

PERF = Path(__file__).resolve().parent
ROOT = PERF.parents[4]
SCRATCH = Path('/home/zhuhe/code/litchi-bitwise-worktree')
TARGET = Path('/home/zhuhe/code/litchi-bitwise-target')
TMP = Path('/home/zhuhe/code/litchi-bitwise-tmp')

def now(): return dt.datetime.now(dt.timezone.utc).isoformat()
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p, v): p.write_text(json.dumps(v, indent=2) + '\n')

def main():
    parser = argparse.ArgumentParser(); parser.add_argument('variant', choices=['baseline', 'candidate']); args = parser.parse_args()
    variant = args.variant; out = PERF / variant; out.mkdir(exist_ok=True)
    harness = PERF / 'bitwise-harness'; relative = harness.relative_to(ROOT)
    shutil.copytree(harness, SCRATCH / relative, dirs_exist_ok=True)
    collector = runpy.run_path(str(SCRATCH / PERF.parent.relative_to(ROOT) / 'gates/run.py'))
    before_time = now(); before = collector['hashes']()
    if variant == 'candidate': assert before == json.loads((PERF.parent / 'gates/results.json').read_text())['source_after']
    command = ['cargo', 'build', '--locked', '--offline', '--release', '--manifest-path', str(relative / 'Cargo.toml')]
    env = {'CARGO_TARGET_DIR': str(TARGET), 'TMPDIR': str(TMP), 'CARGO_INCREMENTAL': '0'}
    started = now()
    with (out / 'build.log').open('w') as log:
        result = subprocess.run(command, cwd=SCRATCH, env=dict(os.environ, **env), stdout=log, stderr=subprocess.STDOUT)
    finished = now(); after = collector['hashes']()
    build = dict(command=command, cwd=str(SCRATCH), environment=env, started_at=started, finished_at=finished, status=result.returncode)
    write(out / 'build-command.json', build)
    write(out / 'source-sha256.json', dict(build, source_before_captured_at=before_time, source_after_captured_at=now(), source_before=before, source_after=after, sources_unchanged=before == after))
    result.check_returncode(); assert before == after
    name = 'ods-formula-bitwise-evaluation-profile'; binary = out / name
    shutil.copyfile(TARGET / 'release' / name, binary); binary.chmod(0o755)
    write(out / 'binary-provenance.json', dict(sha256=sha(binary), size_bytes=binary.stat().st_size, binary=str(binary), build_status=0, base_commit=json.loads((PERF.parent / 'specification.json').read_text())['base_commit']))
    files = ['Cargo.toml', 'Cargo.lock', 'run.py', 'src/main.rs']
    (out / 'harness-sha256.txt').write_text(''.join(f'{sha(harness / f)}  {f}\n' for f in files))
    print(variant, 'built', sha(binary), flush=True)
    for group in ['comparable'] + (['bitwise'] if variant == 'candidate' else []):
        folder = out / group
        cmd = ['python3', str(harness / 'run.py'), '--binary', str(binary), '--output', str(folder), '--revision', variant, '--group', group, '--phase', 'all', '--warmups', '3', '--iterations', '15']
        start = now(); r = subprocess.run(cmd, cwd=ROOT)
        write(folder / 'capture.json', dict(command=cmd, started_at=start, finished_at=now(), status=r.returncode, binary_sha256=sha(binary), raw_sha256=sha(folder / 'raw.csv') if (folder / 'raw.csv').exists() else None))
        r.check_returncode(); print(variant, group, 'captured', flush=True)

if __name__ == '__main__': main()
