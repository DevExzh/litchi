#!/usr/bin/env python3
"""Freeze the qualified pilot before its first native measurement."""
import datetime
import json
from custody import P, ROOT, census, sha

assert not (P / 'capture-freeze.json').exists()
assert not (P / 'capture.json').exists()
assert census() == json.loads((P / 'source-candidate.json').read_text())
for lane, count in [('docx', 5), ('harness', 1)]:
    rows = json.loads((P / f'quality-{lane}.json').read_text())
    assert len(rows) == count and all(row['exit_code'] == 0 for row in rows)
for lane in ['baseline', 'candidate']:
    rows = json.loads((P / f'build-{lane}.json').read_text())
    assert len(rows) == 2 and all(row['exit_code'] == 0 for row in rows)
    assert json.loads((P / 'oracle' / lane / 'result.json').read_text())['exit_code'] == 0
assert (P / 'oracle/baseline/report.json').read_bytes() == (P / 'oracle/candidate/report.json').read_bytes()
assert json.loads((P / 'trace-analysis.json').read_text())['status'] == 'pass'
files = {
    path: expected for path, expected in json.loads(
        (P.parent / 'change-0721/capture-freeze.json').read_text()
    )['files'].items() if '/change-0721/' not in path
}
for path, expected in files.items():
    assert sha(ROOT / path) == expected
names = [
    'freeze.py', 'plan.json', 'pilot.py', 'read-controls.py', 'read-controls-plan.json',
    'analyze.py', 'read-controls-analyze.py', 'analyze-run.py', 'capture.py', 'custody.py',
    'source-guard.py', 'source-guard.json', 'source-baseline.json', 'source-candidate.json',
    'constraints.json', 'host.json', 'hypothesis.md', 'build-baseline.json', 'build-candidate.json',
    'quality-docx.json', 'quality-harness.json', 'oracle/baseline/result.json',
    'oracle/candidate/result.json', 'trace-analysis.json',
    'trace.py', 'trace-run.py', 'trace.fragment', 'trace-analyze.py',
]
for name in names:
    path = P / name
    files[str(path.relative_to(ROOT))] = sha(path)
value = {
    'utc_created': datetime.datetime.now(datetime.timezone.utc).isoformat(),
    'order': 'eight stages: primary native then read controls; after all native stages, first four allocator stages',
    'files': dict(sorted(files.items())),
}
(P / 'capture-freeze.json').write_text(json.dumps(value, indent=2) + '\n')
print(f'Frozen {len(files)} qualified inputs before native capture')
