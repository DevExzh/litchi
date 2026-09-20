#!/usr/bin/env python3
"""Audit unchanged source, fixed diagnostic custody, replay and cleanup."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def read(path): return json.loads(path.read_text())
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    spec=importlib.util.spec_from_file_location('custody0716audit',P/'custody.py')
    C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
    assert C.census()==read(P/'source.json')==read(P.parent/'change-0715/source-final.json')
    motivation=read(P/'motivation.json')
    for n,h in motivation['files'].items(): assert sha(ROOT/motivation['packet']/n)==h
    assert not read(ROOT/motivation['packet']/'pilot-analysis.json')['decision']['accepted']
    for filename in ['constraints.json','helper-freeze.json']:
        for n,h in read(P/filename).items(): assert sha(ROOT/n)==h
    for filename in ['capture-freeze.json','analysis-freeze.json']:
        for n,h in read(P/filename).items(): assert sha(P/n)==h
    build=read(P/'build.json')
    assert build['exit_code']==0 and build['source_sha256']==sha(P/'source.json') and build['log_sha256']==sha(P/'build.log')
    reused=read(P/'reused-verification.json');prior=ROOT/reused['packet']
    for n,h in reused['files'].items(): assert sha(prior/n)==h
    assert read(prior/'quality-source.json')==read(P/'source.json')
    for row in read(prior/'quality.json'): assert row['exit_code']==0 and row['log_sha256']==sha(prior/row['log'])
    for row in read(prior/'evidence/results.json'):
        assert row['exit_code']==0 and row['log_sha256']==sha(prior/'evidence'/(row['name']+'.log'))
    negative=read(P/'negative-checks.json')
    assert negative['status']=='pass' and negative['retained_inputs_unchanged'] and negative['positive_exact_replay']
    assert negative['analysis_sha256']==sha(P/'analysis.json') and negative['analyzer_sha256']==sha(P/'analyze.py')
    assert len(negative['checks'])>=4 and all(c['rejected'] for c in negative['checks'])
    cleanup=read(P/'cleanup.json')
    expected=['/home/zhuhe/code/litchi-target-0716','/home/zhuhe/code/litchi-0716-bin','/home/zhuhe/code/litchi-0716-fs']
    assert cleanup['owned_paths']==expected and cleanup['owned_paths_absent'] and all(not Path(n).exists() for n in expected)
    assert cleanup['binaries']==[build['binary']]
    final=read(P/'final-report-gate.json')
    assert final['exit_code']==0 and final['log_sha256']==sha(P/'final-report-gate.log')
    for n,h in final['docs'].items(): assert sha(ROOT/n)==h
    subprocess.run(['python3','-B',str(P/'analyze.py'),'--check'],cwd=ROOT,check=True)
    print('PASS unchanged source, exact diagnostic replay, negative checks, reused quality and cleanup')

if __name__=='__main__': main()
