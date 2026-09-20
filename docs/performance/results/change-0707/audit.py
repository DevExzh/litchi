#!/usr/bin/env python3
"""Independently replay 0707 diagnostics and verify exact custody and gates."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def main():
    for manifest in ['source-baseline.json','constraints.json']:
        for name,digest in json.loads((P/manifest).read_text()).items():
            assert sha(ROOT/name)==digest,name
    revision=json.loads((P/'plan.json').read_text())['revision']
    subprocess.run(['git','merge-base','--is-ancestor',revision,'HEAD'],cwd=ROOT,check=True)
    for name,digest in json.loads((P/'helper-identities.json').read_text()).items():
        assert sha(ROOT/name)==digest,name
    rows=json.loads((P/'evidence/results.json').read_text())
    assert {r['name'] for r in rows} == {'crate-boundaries','claims','claims-structural','report','coverage','non-iwork'}
    assert len(rows)==6
    for row in rows:
        assert row['exit_code']==0,row['name']
        assert sha(P/'evidence'/(row['name']+'.log'))==row['log_sha256']
        assert row['source_manifest_sha256']==sha(P/'source-baseline.json')
    final=json.loads((P/'final-report-gate.json').read_text())
    assert final['exit_code']==0
    assert sha(P/'final-report-gate.log')==final['log_sha256']
    for name,digest in final['docs'].items():
        assert sha(ROOT/name)==digest,name
    negative=json.loads((P/'analyzer-negative-checks.json').read_text())
    assert negative['analyzer_sha256']==sha(P/'analyze.py')
    assert len(negative['checks'])==3 and all(x['rejected'] for x in negative['checks'])
    with tempfile.TemporaryDirectory(prefix='litchi-0707-audit-') as temporary:
        for script,output in [('analyze.py','analysis.json'),('analyze_profiles.py','profile-analysis.json')]:
            target=Path(temporary)/output
            subprocess.run([sys.executable,'-B',str(P/script),'--output',str(target)],cwd=ROOT,check=True)
            assert target.read_bytes()==(P/output).read_bytes(),output
    print('PASS exact source, constraints, helper identities, six evidence gates and byte-identical diagnostic replays')
if __name__=='__main__': main()
