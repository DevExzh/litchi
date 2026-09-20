#!/usr/bin/env python3
"""Audit restored source, rejected-candidate evidence, replays and cleanup."""
import hashlib,importlib.util,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
    s=importlib.util.spec_from_file_location('custody0715audit',P/'custody.py');C=importlib.util.module_from_spec(s);s.loader.exec_module(C)
    baseline=read(P/'source-baseline.json');candidate=read(P/'source-candidate.json')
    assert C.census()==baseline==read(P/'source-final.json')
    assert candidate==read(P/'quality-source.json')==read(P/'candidate-evidence-source.json')
    assert baseline==read(P/'source.json')==read(P.parent/'change-0714/source.json')
    paths={'crates/litchi-docx/src/package/codec.rs','crates/litchi-docx/src/writer/doc/model.rs','crates/litchi-docx/src/writer/doc/package.rs'}
    assert {n for n in set(baseline)|set(candidate) if baseline.get(n)!=candidate.get(n)}==paths
    for label,source in [('baseline',baseline),('candidate',candidate)]:
        for n in paths:assert sha(P/'source-snapshots'/label/n)==source[n]
        for b in read(P/('build-'+label+'.json')):
            assert b['exit_code']==0 and b['source_manifest_sha256']==sha(P/('source-'+label+'.json'))
            assert b['log_sha256']==sha(P/f"build-{label}-{b['lane']}.log")
    for n,h in read(P/'constraints.json').items():assert sha(ROOT/n)==h
    for n,h in read(P/'helper-freeze.json').items():assert sha(ROOT/n)==h
    for f in ['capture-freeze.json','pilot-freeze.json','mechanism-freeze.json','analysis-freeze.json']:
        for n,h in read(P/f).items():assert sha(P/n)==h
    q=read(P/'quality.json');assert [r['name'] for r in q]==['fmt','tests','clippy','doctests','rustdoc']
    for r in q:
        assert r['exit_code']==0 and r['source_manifest_sha256']==sha(P/'quality-source.json') and r['log_sha256']==sha(P/r['log'])
    gates=read(P/'evidence/results.json');assert len(gates)==6
    for r in gates:
        assert r['exit_code']==0 and r['source_manifest_sha256']==sha(P/'candidate-evidence-source.json') and r['log_sha256']==sha(P/'evidence'/(r['name']+'.log'))
    decision=read(P/'pilot-analysis.json')['decision'];assert not decision['accepted'] and decision['deterministic_output_parity_pass']
    failed=[r for r in decision['hard_gates'] if not r['pass']]
    assert len(decision['hard_gates'])==48 and len(failed)==2
    disposition=read(P/'disposition.json')
    assert disposition['decision']=='rejected' and disposition['production_restored']
    assert disposition['failed_gates']==failed and disposition['pilot_analysis_sha256']==sha(P/'pilot-analysis.json')
    assert not list(P.glob('mechanism-*.receipt.json')) and not (P/'mechanism-analysis.json').exists()
    reused=read(P/'reused-verification.json');prior=ROOT/reused['packet']
    for n,h in reused['files'].items():assert sha(prior/n)==h
    assert read(prior/'quality-source.json')==baseline
    for r in read(prior/'quality.json'):assert r['exit_code']==0 and r['log_sha256']==sha(prior/r['log'])
    for filename,analysis,parser in [('profile-negative-checks.json','profile-analysis.json','analyze_profiles.py'),('pilot-negative-checks.json','pilot-analysis.json','analyze_pilot.py')]:
        r=read(P/filename);assert r['status']=='pass' and r['retained_inputs_unchanged'] and r['positive_exact_replay']
        assert r['analysis_sha256']==sha(P/analysis) and r.get('parser_sha256',r.get('analyzer_sha256'))==sha(P/parser)
        assert all(c['rejected'] for c in r['checks']) and len(r['checks'])>2
    cleanup=read(P/'cleanup.json');expected=['/home/zhuhe/code/litchi-target-0715','/home/zhuhe/code/litchi-0715-bin','/home/zhuhe/code/litchi-0715-fs']
    assert cleanup['owned_paths']==expected and cleanup['owned_paths_absent'] and all(not Path(n).exists() for n in expected)
    binaries=[read(P/'build.json')['binary']]
    for label in ['baseline','candidate']:
        binaries += [dict(path=b['binary'],sha256=b['binary_sha256'],bytes=b['binary_bytes']) for b in read(P/('build-'+label+'.json'))]
    assert cleanup['binaries']==binaries
    final=read(P/'final-report-gate.json');assert final['exit_code']==0 and final['log_sha256']==sha(P/'final-report-gate.log')
    for n,h in final['docs'].items():assert sha(ROOT/n)==h
    for script in ['analyze.py','analyze_pilot.py','analyze-child-shift.py']:
        subprocess.run(['python3','-B',str(P/script),'--check'],cwd=ROOT,check=True)
    print('PASS restored baseline, rejected-candidate evidence, negative checks, exact replay and cleanup')
if __name__=='__main__':main()
