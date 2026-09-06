#!/usr/bin/env python3
"""Verify portable test/read-count/fuzz evidence, including source custody."""
import argparse,gzip,hashlib,importlib.util,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
s=importlib.util.spec_from_file_location('derive',ROOT/'derive.py');d=importlib.util.module_from_spec(s);s.loader.exec_module(d)
def sha(b):return hashlib.sha256(b).hexdigest()
def member(name):
    p=Path(name);d.require(not p.is_absolute() and '..' not in p.parts,'safe artifact path')
    result=(ROOT/p).resolve();d.require(result.is_relative_to(ROOT.resolve()),'artifact path escape');return result
def artifact(row):
    member(row['path']);raw=d.raw(row['path']);d.require(len(raw)==row['bytes'] and sha(raw)==row['sha256'],'artifact '+row['path']);return raw
def manifest(row):
    raw=d.raw(row['path']);v=json.loads(raw);d.require(sha(raw)==row['sha256'] and len(v)==row['files'],'source manifest');return v
def check(sealed,cleanup):
    files=d.load('source-files.json');d.require(files==['crates/soapberry-zip/src/office.rs','crates/soapberry-zip/fuzz/fuzz_targets/parse_zip.rs'],'source file set')
    final=manifest(d.load('checks/final-fuzz-strict.json')['source_after'])
    old=manifest(d.load('checks/final-zip-tests.json')['source_after'])
    d.require({k for k in old.keys()|final.keys() if old.get(k)!=final.get(k)}=={files[1]},'only fuzz changed after workspace tests')
    for n in files:
        d.require(sha(d.raw('candidate/after-'+Path(n).name+'.txt'))==final[n],'exact tested source copy')
    d.require(sha(d.raw('candidate/before-parse_zip.rs.txt'))==old[files[1]],'pre-fuzz tests source')
    expected_failed={'candidate-zip-tests','candidate-zip-strict-r2'}
    required=['final-zip-tests','final-opc-tests','final-pptx-cross-copy-tests','final-zip-check','final-zip-doc','final-workspace-check','final-format','final-boundaries','candidate-zip-strict-r3','fuzz-prepare','fuzz-lock','fuzz-build','fuzz-smoke','final-fuzz-format','final-fuzz-strict']
    for name in required:d.require((ROOT/'checks'/f'{name}.json').is_file(),'required '+name)
    common={k:v for k,v in final.items() if k not in files}
    for path in (ROOT/'checks').glob('*.json'):
        r=d.load(str(path.relative_to(ROOT)))
        if 'source_before' not in r:continue
        passed=path.stem not in expected_failed
        d.require(r['change']==450 and r['driver_sha256']==sha(d.raw('check.py')),'check driver')
        d.require(r['source_unchanged'] and r['source_before']==r['source_after'],'source changed during CPU job')
        m=manifest(r['source_after']);d.require({k:v for k,v in m.items() if k not in files}==common,'unrelated source difference')
        d.require(r['status']==('pass' if passed else 'failed') and (r['exit_code']==0)==passed,'check status '+path.stem)
        artifact(r['log'])
        if path.stem in required:d.require(m[files[0]]==final[files[0]],'required final library source')
    for tag,count in [('final-zip-tests',455),('final-opc-tests',436),('final-pptx-cross-copy-tests',59)]:
        r=d.load('checks/'+tag+'.json');d.require(r['passed_tests']==count and r['failed_tests']==0,'test count '+tag)
    d.require(d.derive()==d.load('measurements.json'),'derived measurements')
    d.require(d.render(d.derive())==(ROOT/'measurements.md').read_text(),'rendered measurements')
    decision=d.load('decision.json');d.require(decision['decision']=='retain measured ZIP enabler' and decision['goal_complete'] is False and decision['pptx_adopted'] is False and decision['latency_claim'] is False,'scoped decision')
    smoke=d.load('checks/fuzz-smoke.json');log=artifact(smoke['log']).decode()
    d.require('Done 1000 runs' in log and '-seed=450' in smoke['argv'] and '-runs=1000' in smoke['argv'],'fuzz completion')
    build=d.load('checks/fuzz-build.json');d.require('RUSTC_BOOTSTRAP=1' in build['argv'] and any('-Z sanitizer=address' in a and '-sanitizer-coverage-level=4' in a for a in build['argv']),'sanitizer coverage flags')
    custody=d.load('checks/fuzz-source-custody.json')
    d.require(custody['task_root']=='/tmp/litchi-goal-0450-zip-fuzz' and custody['status'] in ['prepared','cleaned'],'fuzz custody phase')
    d.require(custody['driver']['sha256']==sha(d.raw('fuzz-control.py')),'fuzz controller')
    for row in custody['source_files']:
        a,b=row['source'],row['copy'];d.require(a['sha256']==b['sha256'] and a['bytes']==b['bytes'],'fuzz copy identity')
        if a['path'].endswith(('.rs','.toml')):d.require(final[a['path']]==a['sha256'],'fuzz copy tested source')
    seeds=d.load('seed-manifest.json');d.require(len(seeds)==13,'seed count')
    for r in seeds:artifact(r)
    d.require({r['sha256'] for r in seeds}=={r['source']['sha256'] for r in custody['seed_files']},'fuzz seed identities')
    if cleanup:
        proof=d.load('checks/fuzz-cleanup.json');d.require(proof['status']=='pass' and proof['temporary_directory_absent'] and proof['removed_directory']=='/tmp/litchi-goal-0450-zip-fuzz','cleanup')
        d.require(custody['status']=='cleaned' and d.load('checks/precleanup.json')['status']=='pass','precleanup replay')
        artifact(proof['cargo_lock']['evidence'])
        d.require(proof['custody']['sha256']==sha(d.raw(proof['custody']['path'])),'cleaned custody identity')
    if sealed:
        inventory={}
        for line in (ROOT/'SHA256SUMS').read_text().splitlines():
            h,n=line.split('  ',1);d.require(n not in inventory and sha(member(n).read_bytes())==h,'seal member');inventory[n]=h
        d.require(set(inventory)=={str(p.relative_to(ROOT)) for p in ROOT.rglob('*') if p.is_file() and p.name!='SHA256SUMS'},'seal coverage')
    return {'status':'pass','change':450,'zip_tests':455,'opc_tests':436,'pptx_tests':59,'io_cases':10,'fuzz_runs':1000,'sealed':sealed,'cleanup':cleanup}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--sealed',action='store_true');p.add_argument('--cleanup',action='store_true');a=p.parse_args()
    try:print(json.dumps(check(a.sealed,a.cleanup)))
    except Exception as e:print('INVALID: '+str(e));raise SystemExit(1)
