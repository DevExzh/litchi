#!/usr/bin/env python3
"""Verify a separate evidence copy after the original raw artifacts are removed."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent
for name in ['/tmp/litchi-goal-0454-native-copy', '/tmp/litchi-goal-0454-opc-fuzz']:
    if Path(name).exists():
        raise RuntimeError('owned raw artifacts must be removed before portable replay')
started = datetime.datetime.now(datetime.timezone.utc).isoformat()
with tempfile.TemporaryDirectory(prefix='litchi-goal-0454-portable-') as directory:
    parent = Path(directory)
    target = parent / 'docs/performance/results/change-0454'
    shutil.copytree(ROOT, target)
    argv = [sys.executable, '-B', str(target / 'verify.py'), '--cleanup']
    result = subprocess.run(argv, cwd=parent, env=os.environ | {
        'PYTHONPATH': '', 'PYTHONDONTWRITEBYTECODE': '1',
    }, capture_output=True, text=True)
    raw = result.stdout + result.stderr
    status = 'pass' if result.returncode == 0 else 'failed'
    if status == 'pass' and json.loads(result.stdout)['status'] != 'pass':
        raise RuntimeError('portable verifier did not certify the retained bundle')
    record = {
        'status': status, 'argv': argv, 'cwd': str(parent),
        'exit_code': result.returncode, 'started_utc': started,
        'finished_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
        'verifier_sha256': hashlib.sha256((target / 'verify.py').read_bytes()).hexdigest(),
        'mode': 'separate evidence copy; original binaries, fixture, raw outputs and fuzz directory absent',
        'stdout': result.stdout, 'stderr': result.stderr,
    }
record['temporary_directory_absent'] = not parent.exists()
with (ROOT / 'checks/portable-verification.json').open('x') as stream:
    stream.write(json.dumps(record, indent=2) + '\n')
print(raw, end='')
raise SystemExit(result.returncode)
