#!/usr/bin/env python3
"""Record the final post-cleanup evidence replays before artifact sealing."""
import json
from pathlib import Path
import subprocess
import sys
from custody import P, ROOT, census, sha

output = P / 'terminal-checks.json'
assert not output.exists()
source = json.loads((P / 'source-final.json').read_text())
assert census() == source
jobs = [('audit.py', []), ('source-guard.py', ['--check']),
        ('memory-diagnostics.py', ['--check']), ('trace-details.py', ['--check']),
        ('quality-summary.py', ['--check'])]
rows = []
for name, args in jobs:
    command = [sys.executable, '-B', str(P / name), *args]
    result = subprocess.run(command, cwd=ROOT, capture_output=True, text=True)
    rows.append({'command': command, 'script_sha256': sha(P / name),
                 'exit_code': result.returncode, 'stdout': result.stdout,
                 'stderr': result.stderr})
    output.write_text(json.dumps({'source_final_sha256': sha(P / 'source-final.json'),
                                 'script_sha256': sha(Path(__file__)),
                                 'checks': rows}, indent=2) + '\n')
    assert result.returncode == 0, rows[-1]
assert census() == source
print('PASS: five post-cleanup checks and exact source custody')
