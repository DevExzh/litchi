#!/usr/bin/env python3
"""Supplemental ABBA to review meaningful full-matrix regression triggers."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
TEMP = Path('/tmp/litchi-goal-0467')
FLAGS = '-C force-frame-pointers=yes -C force-unwind-tables=yes'


def sha(path):
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def main():
    lane = sys.argv[1]
    assert lane in ('A1-guard-clean', 'B1-guard-clean', 'B2-guard-clean', 'A2-guard-clean')
    phase = lane.removesuffix('-guard-clean')
    role = {'A1': 'control', 'A2': 'control', 'B1': 'candidate', 'B2': 'candidate',
            'A-full': 'control', 'B-full': 'candidate', 'A-heap': 'control', 'B-heap': 'candidate'}[phase]
    binding = json.loads((ROOT / (role + '-binding.json')).read_text())
    binary, worktree = TEMP / role, TEMP / (role + '-tree')
    assert sha(binary) == binding['binary_sha256']
    def check():
        assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=worktree, text=True).strip() == binding['revision']
        assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=worktree).strip()
    check()
    output = ROOT / lane
    output.mkdir(exist_ok=False)
    full, heap = phase.endswith('-full'), phase.endswith('-heap')
    samples, warmups = (15, 3) if full else (5, 1) if heap else (100, 5)
    workload = [str(binary), '--workers', '1', '--warmup', str(warmups), '--samples', str(samples),
                '--json', str(output / 'report.json'), '--corpus-manifest', str(output / 'corpus-catalog.json')]
    if not full:
        workload += ['--case', 'cfb_list_streams,cfb_read_one,cfb_shared_concurrent_reads,doc_fresh_write_to,xlsx_source_first_cell,xlsx_source_narrow_column_range_scan',
                     '--xlsx-shape', 'medium']
    argv = ['taskset', '-c', '2', '/usr/bin/time', '-v', '-o', str(output / 'resource.log')]
    if heap:
        argv += ['heaptrack', '-o', str(output / 'heaptrack'), '--']
    argv += workload
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', RUSTFLAGS=FLAGS, CARGO_PROFILE_RELEASE_DEBUG='1',
               CARGO_INCREMENTAL='0', CARGO_BUILD_JOBS='4', DEBUGINFOD_URLS='', LC_ALL='C', PYTHONDONTWRITEBYTECODE='1')
    receipt = dict(schema='litchi-0467-capture-v1', lane=lane, role=role, revision=binding['revision'],
                   binary_sha256=sha(binary), binding_sha256=sha(ROOT / (role + '-binding.json')),
                   driver_sha256=sha(Path(__file__)), argv=argv, cwd=str(worktree), samples=samples, warmups=warmups,
                   environment={k: env[k] for k in ['RUSTUP_TOOLCHAIN','RUSTFLAGS','CARGO_PROFILE_RELEASE_DEBUG','DEBUGINFOD_URLS','LC_ALL']},
                   started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(), clean_before=True)
    (output / 'started.json').write_text(json.dumps(receipt, indent=2) + '\n')
    with (output / 'stdout.log').open('w') as stdout, (output / 'stderr.log').open('w') as stderr:
        result = subprocess.run(argv, cwd=worktree, env=env, stdout=stdout, stderr=stderr)
    check()
    assert sha(binary) == binding['binary_sha256']
    report_ok = False
    if result.returncode == 0:
        report = json.loads((output / 'report.json').read_text())
        report_ok = (report['environment']['git_revision'] == binding['revision']
                     and report['environment']['git_worktree_dirty'] is False)
    receipt.update(exit_code=result.returncode, clean_after=True, binary_unchanged=True,
                   report_metadata_matches_clean_role=report_ok,
                   finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                   artifacts={p.name: dict(sha256=sha(p), bytes=p.stat().st_size) for p in output.iterdir() if p.is_file()})
    (output / 'receipt.json').write_text(json.dumps(receipt, indent=2) + '\n')
    print(json.dumps(dict(lane=lane, exit_code=result.returncode)), flush=True)
    raise SystemExit(result.returncode if result.returncode else int(not report_ok))


if __name__ == '__main__':
    main()
