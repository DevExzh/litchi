#!/usr/bin/env python3
"""Recompute retained conclusions and verify source, constraints and gates."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def main():
    helper = json.loads((P / 'tooling-checks.json').read_text())['profile_helper']
    assert sha(ROOT / helper['path']) == helper['sha256']
    for manifest in ['constraints.json', 'workspace-inputs.json', 'source-manifest.json']:
        for name, digest in json.loads((P / manifest).read_text()).items():
            assert sha(ROOT / name) == digest, name
    base = json.loads((P / 'plan.json').read_text())['revision']
    subprocess.run(['git', 'merge-base', '--is-ancestor', base, 'HEAD'], cwd=ROOT, check=True)
    assert json.loads((P / 'initial-build.json').read_text())['binary_sha256'] == json.loads((P / 'build.json').read_text())['binary_sha256']
    for receipt in json.loads((P / 'evidence/results.json').read_text()):
        assert receipt['exit_code'] == 0, receipt['name']
        assert sha(P / 'evidence' / (receipt['name']+'.log')) == receipt['log_sha256']
    assert len(json.loads((P / 'evidence/results.json').read_text())) == 6
    with tempfile.TemporaryDirectory(prefix='litchi-0705-audit-') as temporary:
        for script, output, positional in [
            ('analyze.py','analysis.json',True),
            ('analyze_allocations.py','allocation-analysis.json',False),
            ('analyze_profiles.py','profile-analysis.json',False),
        ]:
            result = Path(temporary) / output
            command = [sys.executable, '-B', str(P / script)]
            command += [str(result)] if positional else ['--output', str(result)]
            subprocess.run(command, cwd=ROOT, check=True)
            assert json.loads(result.read_text()) == json.loads((P / output).read_text()), output
    expected = sha(P / 'control-analysis.json')
    subprocess.run([sys.executable, '-B', str(P/'analyze_controls.py')], cwd=ROOT, check=True)
    assert sha(P / 'control-analysis.json') == expected
    print('PASS source, constraints, six evidence gates and four independent recomputations')

if __name__ == '__main__': main()
