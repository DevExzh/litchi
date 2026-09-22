#!/usr/bin/env python3
"""Replay final evidence after cleanup and retain terminal receipts."""
import hashlib
import json
from pathlib import Path
import os
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def read(p):
    return json.loads(p.read_text())

cleanup = read(P / 'cleanup.json')
assert cleanup['removed'] and all(not Path(r).exists() for r in cleanup['roots'])
constraints = read(P / 'constraints.json')
assert all(sha(ROOT / name) == value for name, value in constraints.items())
stable = {name: sha(P / name) for name in ['analysis.json', 'negative-checks.json', 'report-stats.json']}
out = P / 'terminal'
out.mkdir(exist_ok=True)
runs = []
for name in ['source-guard.py', 'analyze.py', 'audit.py', 'negative-checks.py', 'report-stats.py']:
    command = ['python3', str(P / name)]
    result = subprocess.run(command, cwd=ROOT, env=dict(os.environ, PYTHONDONTWRITEBYTECODE='1'),
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    path = out / (name + '.log')
    path.write_bytes(result.stdout)
    assert result.returncode == 0, result.stdout.decode()
    runs.append(dict(command=command, exit_code=result.returncode,
                     output=str(path.relative_to(P)), sha256=sha(path)))
assert stable == {name: sha(P / name) for name in stable}
reports = ['0732-ppt-native-phase-attribution.md', 'BASELINE.md', 'HOTSPOTS.md', 'REPORT.md']
receipt = dict(status='passed', post_cleanup=True, constraint_count=len(constraints),
               stable_replays=stable, runs=runs,
               reports={str((P.parents[1] / name).relative_to(ROOT)): sha(P.parents[1] / name)
                        for name in reports})
(P / 'terminal.json').write_text(json.dumps(receipt, indent=2) + '\n')
print('PASS post-cleanup source, constraints, independent audit, and deterministic negative/statistics replay')
