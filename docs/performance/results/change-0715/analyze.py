#!/usr/bin/env python3
"""Validate current serialization profiles and their complete capture custody."""
import argparse
import importlib.util
import json
from pathlib import Path
P=Path(__file__).resolve().parent

def module(path,name):
    spec=importlib.util.spec_from_file_location(name,path)
    result=importlib.util.module_from_spec(spec);spec.loader.exec_module(result);return result
C=module(P/'capture.py','capture0715analysis')
L=module(P.parent/'change-0709/analyze.py','report0709profile0715')
N=module(P.parent/'change-0713/analyze.py','normalize0713profile0715')

def analyze():
    plan,build,source=[C.read(P/n) for n in ['plan.json','build.json','source.json']]
    assert source==C.read(P.parent/'change-0714/source.json')
    current=C.C.census()
    if current!=source:
        assert current==C.read(P/'source-candidate.json')
        assert {n for n in set(source)|set(current) if source.get(n)!=current.get(n)} <= {
            'crates/litchi-docx/src/package/codec.rs','crates/litchi-docx/src/writer/doc/model.rs','crates/litchi-docx/src/writer/doc/package.rs'}
    for filename in ['capture-freeze.json']:
        for n,h in C.read(P/filename).items():assert C.C.sha(P/n)==h
    for n,h in C.read(P/'constraints.json').items():assert C.C.sha(C.C.ROOT/n)==h
    assert build['exit_code']==0 and build['source_sha256']==C.C.sha(P/'source.json')
    assert build['log_sha256']==C.C.sha(P/'build.log')
    binary=Path(build['binary']['path'])
    if binary.exists():
        assert C.C.sha(binary)==build['binary']['sha256'] and binary.stat().st_size==build['binary']['bytes']
    else:assert build['binary'] in C.read(P/'cleanup.json')['binaries']
    meta=dict(binary_sha256=build['binary']['sha256'],binary_bytes=build['binary']['bytes'])
    T=module(P/'analyze_profiles.py','profile0715parser')
    rows=[]
    for job in C.jobs(plan,'profile'):
        name=job['name']; receipt=C.read(P/(name+'.receipt.json'))
        assert receipt['job']==job and receipt['command']==C.command(plan,job,build)
        assert receipt['exit_code']==0 and receipt['binary']==build['binary']
        assert receipt['source_before']==receipt['source_after']==C.digest(source)
        for k,n in [('source_manifest_sha256','source.json'),('build_sha256','build.json'),('plan_sha256','plan.json'),('script_sha256','capture.py')]:assert receipt[k]==C.C.sha(P/n)
        assert receipt['fixture_before']==receipt['fixture_after']==C.fixture(job['corpus'])
        assert receipt['environment']==dict(LC_ALL='C',LANG='C',TZ='UTC',RUSTFLAGS=None,LD_PRELOAD=None,PERL_HASH_SEED='0',PERL_PERTURB_KEYS='0')
        expected={name+s for s in ['.json','.stdout','.stderr','.callgrind']}
        expected.update(name+'.callgrind.'+str(i) for i in range(1,plan['profile']['numbered_parts'][job['corpus']['id']]+1))
        assert set(receipt['artifacts'])==expected
        for n,h in receipt['artifacts'].items():assert C.C.sha(P/n)==h
        report=C.read(P/(name+'.json')); L.check_report_metadata(report,meta,'native',job['case'],1,0,name)
        result=report['results'][0]; L.validate_elapsed(result['elapsed_ns'],1,name)
        L.validate_operation_metrics(result['operation_metrics'],result['elapsed_ns'],1,'native',name)
        L.validate_ordinary_save(result,job['corpus'],'counting_publish','native',1,name)
        prior=C.read(P.parent/'change-0714'/('native-r1-'+job['corpus']['id']+'-counting_publish.json'))['results'][0]
        assert N.normalized_result(result)==N.normalized_result(prior)
        rows.append(dict(name=name,corpus=job['corpus']['id'],repeat=job['repeat'],report_sha256=C.C.sha(P/(name+'.json')),receipt_sha256=C.C.sha(P/(name+'.receipt.json')),profile=T.analyze_profile(P/(name+'.callgrind'),job['corpus']['id'],plan['profile'])))
    return dict(schema_version=1,revision=plan['revision'],source_sha256=C.C.sha(P/'source.json'),binary=build['binary'],profiles=rows,normalized_parity_with_0714=True,performance_claim='none; current-source guest-instruction attribution',limitations=plan['claims'])

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--check',action='store_true');args=parser.parse_args()
    data=(json.dumps(analyze(),indent=2,sort_keys=True)+'\n').encode();out=P/'profile-analysis.json'
    if args.check:assert out.read_bytes()==data
    else:assert not out.exists();out.write_bytes(data)
    print('PASS: four source-bound serialization profiles and exact semantic parity')
if __name__=='__main__':main()
