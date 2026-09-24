#!/usr/bin/env python3
"""Capture the frozen logical evaluator against its saved baseline binary."""
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
SCRATCH = Path('/var/tmp/ods-logical-worktree')
TARGET = Path('/var/tmp/ods-logical-target')
TMP = Path('/var/tmp/ods-logical-tmp')
BINARY = 'ods-formula-logical-evaluation-profile'

def now():
    return dt.datetime.now(dt.timezone.utc).isoformat()

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def write(path, value):
    path.write_text(json.dumps(value, indent=2) + '\n')

def run_group(variant, group):
    folder = PERF / variant / group
    command = ['python3', str(PERF / 'logical-harness/run.py'), '--binary',
               str(PERF / variant / BINARY), '--output', str(folder),
               '--revision', variant, '--group', group, '--phase', 'all',
               '--warmups', '3', '--iterations', '15']
    start = now()
    result = subprocess.run(command, cwd=ROOT)
    write(folder / 'capture.json', {'command': command, 'started_at': start,
          'finished_at': now(), 'status': result.returncode,
          'binary_sha256': digest(PERF / variant / BINARY),
          'raw_sha256': digest(folder / 'raw.csv') if (folder / 'raw.csv').exists() else None})
    result.check_returncode()
    print(variant, group, 'captured', flush=True)

def main():
    baseline = PERF / 'baseline'
    for line in (baseline / 'harness-sha256.txt').read_text().splitlines():
        expected, name = line.split(maxsplit=1)
        assert digest(PERF / 'logical-harness' / name) == expected
    assert digest(baseline / BINARY) == json.loads((baseline / 'binary-provenance.json').read_text())['sha256']
    run_group('baseline', 'comparable')
    out = PERF / 'candidate'
    out.mkdir(exist_ok=True)
    relative_harness = (PERF / 'logical-harness').relative_to(ROOT)
    shutil.copytree(PERF / 'logical-harness', SCRATCH / relative_harness, dirs_exist_ok=True)
    collector = runpy.run_path(str(SCRATCH / PERF.parent.relative_to(ROOT) / 'gates/run.py'))
    before_time = now()
    before = collector['hashes']()
    gates = json.loads((PERF.parent / 'gates/results.json').read_text())
    assert before == gates['source_after']
    command = ['cargo', 'build', '--locked', '--offline', '--release', '--manifest-path', str(relative_harness / 'Cargo.toml')]
    environment = {'CARGO_TARGET_DIR': str(TARGET), 'TMPDIR': str(TMP), 'CARGO_INCREMENTAL': '0'}
    start = now()
    with (out / 'build.log').open('w') as log:
        result = subprocess.run(command, cwd=SCRATCH, env=dict(os.environ, **environment), stdout=log, stderr=subprocess.STDOUT)
    finish = now()
    after = collector['hashes']()
    build = {'command': command, 'cwd': str(SCRATCH), 'environment': environment,
             'started_at': start, 'finished_at': finish, 'status': result.returncode}
    write(out / 'build-command.json', build)
    source = dict(build, base_commit=gates['head'], files_count=len(before),
                  source_before_captured_at=before_time, source_before=before,
                  source_after_captured_at=now(), source_after=after, sources_unchanged=before == after)
    write(out / 'source-sha256.json', source)
    result.check_returncode()
    assert before == after
    shutil.copy2(TARGET / 'release' / BINARY, out / BINARY)
    shutil.copy2(baseline / 'harness-sha256.txt', out / 'harness-sha256.txt')
    write(out / 'binary-provenance.json', {'binary': str(out / BINARY),
          'target_binary': str(TARGET / 'release' / BINARY), 'sha256': digest(out / BINARY),
          'size_bytes': (out / BINARY).stat().st_size, 'base_commit': gates['head'],
          'source_manifest': 'source-sha256.json', 'harness_manifest': 'harness-sha256.txt',
          'build_command': 'build-command.json', 'build_status': 0, 'captured_at': now()})
    print('candidate built', digest(out / BINARY), flush=True)
    run_group('candidate', 'comparable')
    run_group('candidate', 'logical')

if __name__ == '__main__':
    main()
