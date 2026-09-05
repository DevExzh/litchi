#!/usr/bin/env python3
"""One normal-binary identity check, excluded from the allocator baseline."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
OUTPUT = ROOT / 'checks/normal-identity'
if OUTPUT.exists():
    raise SystemExit('refusing existing normal identity check')
build = json.loads((ROOT / 'build-candidate.json').read_text())
protocol = json.loads((ROOT / 'protocol.json').read_text())
source = build['source_before']
binary = build['binaries']['normal']
assert build['status'] == 'pass' and source == build['source_after']
with Path(binary['path']).open('rb') as stream:
    assert hashlib.file_digest(stream, 'sha256').hexdigest() == binary['sha256']
assert subprocess.check_output(['git', 'status', '--porcelain'], cwd=source['worktree']) == b''
OUTPUT.mkdir()
command = [binary['path'], '--case', 'pptx_cross_copy_plain_lifecycle', *protocol['common_flags'], '--samples', '1', '--warmup', '0', '--json', str(OUTPUT / 'report.json'), '--corpus-manifest', str(OUTPUT / 'catalog.json')]
with (OUTPUT / 'stdout.txt').open('wb') as stdout, (OUTPUT / 'stderr.txt').open('wb') as stderr:
    result = subprocess.run(command, cwd=source['worktree'], env=os.environ | {'RUSTUP_TOOLCHAIN': '1.98.1'}, stdout=stdout, stderr=stderr)
assert result.returncode == 0
report = json.loads((OUTPUT / 'report.json').read_text())
assert report['tool']['instrumentation'] == 'none'
assert 'allocator_counter_revision' not in report['tool']
assert report['results'][0]['operation_metrics']['allocation']['status'] == 'unavailable'
assert report['environment']['git_revision'] == source['revision']
assert report['environment']['git_worktree_dirty'] is False
assert report['binary_identity']['binary_sha256'] == binary['sha256']
verify_command = [sys.executable, str(ROOT / 'verify.py'), '--repo-root', str(REPO), '--report', str(OUTPUT / 'report.json'), '--catalog', str(OUTPUT / 'catalog.json'), '--lane', 'normal', '--samples', '1', '--warmups', '0', '--selector', 'pptx_cross_copy_plain_lifecycle']
verification = subprocess.run(verify_command, capture_output=True, text=True)
(OUTPUT / 'verification.txt').write_text(verification.stdout + verification.stderr)
assert verification.returncode == 0
assert subprocess.check_output(['git', 'status', '--porcelain'], cwd=source['worktree']) == b''
record = {'change': 422, 'classification': 'functional normal-binary identity check; excluded from baseline timing', 'status': 'pass', 'argv': command, 'verify_argv': verify_command, 'source_revision': source['revision'], 'binary_sha256': binary['sha256'], 'counter_marker_omitted': True, 'allocation_unavailable': True}
(OUTPUT / 'check.json').write_text(json.dumps(record, indent=2) + '\n')
print(json.dumps({'status': 'pass', 'normal_counter_marker_omitted': True}))
