#!/usr/bin/env python3
"""Replay the finalized 0708 experiment and verify quality/evidence custody."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def read(path):return json.loads(path.read_text())
def main():
    for manifest in ['source-final.json','constraints.json','helper-identities.json']:
        for name,digest in read(P/manifest).items():assert sha(ROOT/name)==digest,name
    spec=importlib.util.spec_from_file_location('custody0708',P/'build.py')
    B=importlib.util.module_from_spec(spec);spec.loader.exec_module(B)
    assert B.census()==read(P/'source-final.json')
    base=read(P/'plan.json')['revision']
    subprocess.run(['git','merge-base','--is-ancestor',base,'HEAD'],cwd=ROOT,check=True)
    for name,digest in read(P/'capture-freeze.json')['sha256'].items():assert sha(P/name)==digest,name
    rows=read(P/'evidence/results.json')
    assert len(rows)==6 and {r['name'] for r in rows}=={'crate-boundaries','claims','claims-structural','report','coverage','non-iwork'}
    for row in rows:
        assert row['exit_code']==0,row['name']
        assert sha(P/'evidence'/(row['name']+'.log'))==row['log_sha256']
        assert row['source_manifest_sha256']==sha(P/'source-final.json')
    disposition=read(P/'disposition.json')
    analysis=read(P/'analysis.json')
    assert analysis['status']=='pass' and analysis['decision']==disposition['decision']=='reject'
    assert disposition['retained'] is False and disposition['production_restored'] is True
    native=analysis['admission']['native']
    assert len(native['rows'])==disposition['native_primary_rows_total']==8
    assert sum(not row['passed'] for row in native['rows'])==disposition['native_primary_rows_failed']==8
    assert analysis['admission']['passed'] is False
    assert analysis['admission']['allocator']['passed']==disposition['allocation_gate_passed']==True
    assert sha(P/'pilot-analysis.json')==disposition['pilot_analysis_sha256']
    corrections=read(P/'retained-test-corrections.json')
    assert sha(P/corrections['patch'])==corrections['patch_sha256']
    candidate=read(P/'source-candidate.json')
    for name,entry in corrections['changes'].items():
        assert entry['before_sha256']==candidate[name] and entry['after_sha256']==sha(ROOT/name)
    negative=read(P/'negative-checks.json')
    assert negative['status']=='pass' and negative['analyzer_sha256']==sha(P/'analyze.py')
    assert len(negative['checks'])==7 and all(row['rejected'] for row in negative['checks'])
    phases=['baseline-focused','candidate-focused']
    phases += ['candidate-full'] if disposition['retained'] else ['restored-focused','restored-quality']
    for phase in phases:
        rows=read(P/('quality-'+phase+'.json'));assert len(rows)==(5 if phase=='candidate-full' else 3)
        for row in rows:
            assert row['exit_code']==0,(phase,row['name'])
            assert sha(P/row['log'])==row['log_sha256']
            assert sha(P/row['source_manifest'])==row['source_manifest_sha256']
    final=read(P/'final-report-gate.json');assert final['exit_code']==0
    assert sha(P/'final-report-gate.log')==final['log_sha256']
    for name,digest in final['docs'].items():assert sha(ROOT/name)==digest,name
    scripts=[('analyze.py','analysis.json')]
    if (P/'mechanism-analysis.json').is_file():scripts.append(('analyze_mechanism.py','mechanism-analysis.json'))
    if disposition['retained']:assert len(scripts)==2,'retention requires mechanism gate'
    with tempfile.TemporaryDirectory(prefix='litchi-0708-audit-') as temporary:
        for script,output in scripts:
            target=Path(temporary)/output
            subprocess.run([sys.executable,'-B',str(P/script),'--output',str(target)],cwd=ROOT,check=True)
            assert target.read_bytes()==(P/output).read_bytes(),output
    print('PASS exact source, constraints, capture freeze, quality/evidence gates and byte-identical analysis replay')
if __name__=='__main__':main()
