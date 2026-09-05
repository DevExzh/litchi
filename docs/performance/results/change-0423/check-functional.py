#!/usr/bin/env python3
"""Four one-sample normal-binary checks excluded from the formal matrix."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
OUTPUT = ROOT / 'checks/functional'
if OUTPUT.exists():
    raise SystemExit('refusing existing functional checks')
build = json.loads((ROOT / 'build-candidate.json').read_text())
protocol = json.loads((ROOT / 'protocol.json').read_text())
source = build['source_before']
binary = build['binaries']['normal']
assert build['status'] == 'pass' and source == build['source_after']
def digest(path):
    with Path(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()
assert digest(binary['path']) == binary['sha256']
assert subprocess.check_output(['git', 'status', '--porcelain'], cwd=source['worktree']) == b''
OUTPUT.mkdir()
records = []
for selector in protocol['selectors']:
    folder = OUTPUT / selector
    folder.mkdir()
    command = ['taskset', '-c', str(protocol['cpu']), binary['path'], '--case', selector, *protocol['common_flags'], '--samples', '1', '--warmup', '0', '--json', str(folder / 'report.json'), '--corpus-manifest', str(folder / 'catalog.json')]
    with (folder / 'stdout.txt').open('wb') as stdout, (folder / 'stderr.txt').open('wb') as stderr:
        result = subprocess.run(command, cwd=source['worktree'], env=os.environ | {'RUSTUP_TOOLCHAIN': '1.98.1'}, stdout=stdout, stderr=stderr)
    assert result.returncode == 0, selector
    report = json.loads((folder / 'report.json').read_text())
    assert report['tool']['instrumentation'] == 'none'
    assert 'allocator_counter_revision' not in report['tool']
    assert report['results'][0]['operation_metrics']['allocation']['status'] == 'unavailable'
    assert report['environment']['git_revision'] == source['revision']
    assert report['environment']['git_worktree_dirty'] is False
    assert report['binary_identity']['binary_sha256'] == binary['sha256']
    verify_command = [sys.executable, str(ROOT / 'verify.py'), '--repo-root', str(REPO), '--report', str(folder / 'report.json'), '--catalog', str(folder / 'catalog.json'), '--lane', 'normal', '--contract', 'functional', '--samples', '1', '--warmups', '0', '--selector', selector]
    verification = subprocess.run(verify_command, capture_output=True, text=True, env=os.environ | {'PYTHONDONTWRITEBYTECODE': '1'})
    (folder / 'verification.txt').write_text(verification.stdout + verification.stderr)
    assert verification.returncode == 0, selector
    assert subprocess.check_output(['git', 'status', '--porcelain'], cwd=source['worktree']) == b''
    assert digest(binary['path']) == binary['sha256']
    record = {'selector': selector, 'status': 'pass', 'argv': command, 'verify_argv': verify_command, 'source_revision': source['revision'], 'binary_sha256': binary['sha256'], 'counter_marker_omitted': True, 'allocation_unavailable': True, 'artifacts': {p.name: {'sha256': digest(p), 'bytes': p.stat().st_size} for p in folder.iterdir()}}
    (folder / 'check.json').write_text(json.dumps(record, indent=2) + '\n')
    records.append(record)
    print(selector, 'pass', flush=True)
(OUTPUT / 'check.json').write_text(json.dumps({'change': 423, 'classification': 'functional one-sample zero-warmup checks; excluded from formal baseline', 'status': 'pass', 'runs': records}, indent=2) + '\n')
