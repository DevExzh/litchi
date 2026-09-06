#!/usr/bin/env python3
"""Replay OPC correctness/I/O/fuzz evidence without Cargo or the original checkout."""
import argparse,hashlib,importlib.util,json,re
from pathlib import Path
ROOT=Path(__file__).resolve().parent
s=importlib.util.spec_from_file_location('derive',ROOT/'derive.py');d=importlib.util.module_from_spec(s);s.loader.exec_module(d)
def sha(raw):return hashlib.sha256(raw).hexdigest()
def member(name):
    p=Path(name);d.require(not p.is_absolute() and '..' not in p.parts,'safe path')
    p=(ROOT/p).resolve();d.require(p.is_relative_to(ROOT.resolve()),'artifact escape');return p
def artifact(r):
    member(r['path']);raw=d.raw(r['path']);d.require(len(raw)==r['bytes'] and sha(raw)==r['sha256'],'artifact '+r['path']);return raw
def manifest(r):
    member(r['path']);raw=d.raw(r['path']);v=json.loads(raw);d.require(sha(raw)==r['sha256'] and len(v)==r['files'],'source manifest');return v
def check(sealed=False,cleanup=False):
    files=d.load('source-files.json')
    d.require(files==['crates/litchi-opc/src/source_backed.rs','crates/litchi-opc/tests/source_part_transfer.rs','crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs'],'source set')
    final=manifest(d.load('checks/final-opc-tests.json')['source_after'])
    for n in files:d.require(sha(d.raw('candidate/after-'+Path(n).name+'.txt'))==final[n],'exact tested candidate')
    common={k:v for k,v in final.items() if k not in files}
    failed={'candidate-opc-check','candidate-opc-clippy'}
    required=['final-opc-clippy','final-opc-tests','final-pptx-cross-copy-tests','final-harness-tests','final-harness-feature-bins','final-opc-doc','final-workspace-check','final-format','final-boundaries','fuzz-lock','fuzz-build','fuzz-smoke','final-fuzz-strict','final-fuzz-format']
    for tag in required:d.require((ROOT/'checks'/f'{tag}.json').is_file(),'required '+tag)
    for p in (ROOT/'checks').glob('*.json'):
        r=d.load(str(p.relative_to(ROOT)))
        if 'source_before' not in r:continue
        d.require(r['change']==451 and r['driver_sha256']==sha(d.raw('check.py')),'check driver')
        d.require(r['source_unchanged'] and r['source_before']==r['source_after'],'source changed during check')
        m=manifest(r['source_after']);d.require({k:v for k,v in m.items() if k not in files}==common,'unrelated source difference')
        if p.stem in required:d.require(m==final,'final source scope '+p.stem)
        passed=p.stem not in failed
        d.require(r['status']==('pass' if passed else 'failed') and (r['exit_code']==0)==passed,'check status '+p.stem)
        log=artifact(r['log']).decode()
        if 'test' in r['argv']:
            rows=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',log)
            for i,key in enumerate(['passed_tests','failed_tests','ignored_tests']):d.require(r[key]==sum(int(v[i]) for v in rows),'derived test count')
    for tag,count in [('final-opc-tests',466),('final-pptx-cross-copy-tests',59),('final-harness-tests',361),('final-harness-feature-bins',20)]:
        r=d.load('checks/'+tag+'.json');d.require(r['passed_tests']==count and r['failed_tests']==0,'test count '+tag)
    initial=manifest(d.load('checks/candidate-opc-check.json')['source_after'])
    intermediate=manifest(d.load('checks/candidate-opc-clippy.json')['source_after'])
    d.require(sha(d.raw('draft-history/initial-source_backed.rs.txt'))==initial[files[0]],'initial compile draft')
    d.require(sha(d.raw('draft-history/pre-clippy-source_backed.rs.txt'))==intermediate[files[0]],'lint draft')
    d.require(d.derive()==d.load('measurements.json'),'derived measurements')
    d.require(d.render(d.derive())==d.raw('measurements.md').decode(),'rendered measurements')
    decision=d.load('decision.json')
    d.require(decision['decision']=='retain OPC combined capture' and all(decision[k] is False for k in ['goal_complete','pptx_adopted','latency_claim','memory_peak_claim']),'scoped decision')
    smoke=d.load('checks/fuzz-smoke.json');log=artifact(smoke['log']).decode()
    d.require('Done 1000 runs' in log and '-runs=1000' in smoke['argv'] and '-seed=451' in smoke['argv'],'fuzz completion')
    build=d.load('checks/fuzz-build.json');d.require('RUSTC_BOOTSTRAP=1' in build['argv'] and any('-Z sanitizer=address' in a and '-sanitizer-coverage-level=4' in a for a in build['argv']),'instrumented build')
    prepared=d.load('checks/fuzz-prepared.json');d.require(prepared['status']=='pass' and prepared['task']=='/tmp/litchi-goal-0451-opc-fuzz' and prepared['driver_sha256']==sha(d.raw('fuzz-control.py')),'fuzz preparation')
    inputs={r['path']:r for r in prepared['inputs']};d.require(len(inputs)==18,'16 seeds and two build inputs')
    d.require(inputs['fuzz_targets/parse_opc.rs']['sha256']==final[files[2]],'fuzz source copy')
    d.require(inputs['Cargo.toml']['sha256']==sha(d.raw('checks/fuzz-Cargo.toml.txt')),'fuzz manifest copy')
    seeds=d.load('seed-manifest.json');d.require(len(seeds)==16,'seed count')
    for r in seeds:
        artifact(r);c=inputs['corpus/'+Path(r['path']).name];d.require(c['sha256']==r['sha256'] and c['bytes']==r['bytes'],'seed copy')
    artifacts=d.load('checks/fuzz-artifacts.json');a={r['path']:r for r in artifacts['artifacts']}
    d.require(artifacts['status']=='pass' and artifacts['task']==prepared['task'] and artifacts['files']==len(a) and artifacts['bytes']==sum(r['bytes'] for r in a.values()),'artifact inventory')
    for n,r in inputs.items():d.require(a[n]==r,'preserved fuzz input')
    d.require(a['Cargo.lock']['sha256']==sha(d.raw('checks/fuzz-Cargo.lock.txt')),'retained generated lock')
    d.require(a['target/x86_64-unknown-linux-gnu/release/parse_opc']['bytes']>0,'binary identity')
    if cleanup:
        p=d.load('checks/fuzz-cleanup.json');d.require(p['status']=='pass' and p['temporary_directory_absent'] and p['task']==prepared['task'],'cleanup result')
        d.require(p['files_removed']==artifacts['files'] and p['bytes_removed']==artifacts['bytes'] and p['artifact_manifest_sha256']==sha(d.raw('checks/fuzz-artifacts.json')),'cleanup custody')
        d.require(d.load('checks/precleanup.json')['status']=='pass','precleanup replay')
    if sealed:
        inventory={}
        for line in (ROOT/'SHA256SUMS').read_text().splitlines():
            h,n=line.split('  ',1);d.require(n not in inventory and sha(member(n).read_bytes())==h,'seal member');inventory[n]=h
        d.require(set(inventory)=={str(p.relative_to(ROOT)) for p in ROOT.rglob('*') if p.is_file() and p.name!='SHA256SUMS'},'seal coverage')
    return {'status':'pass','change':451,'opc_tests':466,'pptx_tests':59,'harness_tests':381,'io_cases':8,'fuzz_runs':1000,'sealed':sealed,'cleanup':cleanup}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--sealed',action='store_true');p.add_argument('--cleanup',action='store_true');a=p.parse_args()
    print(json.dumps(check(a.sealed,a.cleanup)))
