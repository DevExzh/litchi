#!/usr/bin/env python3
"""Capture the fixed stack-attribution matrix with immutable child artifacts."""
import datetime,hashlib,importlib.util,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('custody0718capture',P/'custody.py');C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
def read(path):return json.loads(path.read_text())
def digest(value):return hashlib.sha256(json.dumps(value,sort_keys=True,separators=(',',':')).encode()).hexdigest()
def jobs(plan):
    result=[]
    for lane,settings in plan['lanes'].items():
        for repeat in range(1,settings['repeats']+1):
            corpora=plan['corpora'] if repeat==1 else list(reversed(plan['corpora']))
            for corpus in corpora:
                prefix='docx_ordinary_save_' if corpus['id']=='generated' else 'docx_real_file_ordinary_save_'
                result.append(dict(name=f'{lane}-r{repeat}-{corpus["id"]}',lane=lane,repeat=repeat,corpus=corpus,phase=plan['phase'],case=prefix+plan['phase'],samples=settings['samples'],warmup=settings['warmup'],order_index=len(result)))
    assert len(result)==plan['expected_children'];return result
def command(plan,job,build):
    setting=plan['lanes'][job['lane']]
    argv=['taskset','-c',str(plan['cpu']),setting['tool'],*setting['options']]
    if job['lane']=='mapping':argv+=['-o',str(P/(job['name']+'.strace'))]
    else:argv+=['--dhat-out-file='+str(P/(job['name']+'.dhat.json'))]
    argv += [build['binary']['path'],'--warmup',str(job['warmup']),'--samples',str(job['samples']),'--case',job['case'],'--json',str(P/(job['name']+'.json')),'--filesystem-root',plan['filesystem_root']]
    if job['corpus']['path']:argv+=['--ooxml-file',job['corpus']['path']]
    return argv
def fixture(corpus):
    if corpus['path'] is None:return None
    path=C.ROOT/corpus['path'];value=dict(path=corpus['path'],bytes=path.stat().st_size,sha256=C.sha(path));assert value['bytes']==corpus['bytes'] and value['sha256']==corpus['sha256'];return value
def main():
    plan,build,source=[read(P/n) for n in ['plan.json','build.json','source.json']]
    for n,h in read(P/'capture-freeze.json').items():assert C.sha(P/n)==h
    for filename in ['constraints.json','helper-freeze.json']:
        for n,h in read(P/filename).items():assert C.sha(C.ROOT/n)==h
    for n,value in read(P/'tool-identities.json').items():assert C.sha(Path(n))==value['sha256'] and Path(n).stat().st_size==value['bytes']
    binary=Path(build['binary']['path']);assert build['exit_code']==0 and build['source_sha256']==C.sha(P/'source.json') and C.sha(binary)==build['binary']['sha256']
    env=dict(os.environ);env.update(plan['environment_overrides']);environment={k:env.get(k) for k in plan['environment_keys']};assert environment==plan['expected_environment']
    for job in jobs(plan):
        name=job['name'];assert not list(P.glob(name+'.*')),'refusing child replacement'
        before=C.census();assert before==source;fixture_before=fixture(job['corpus']);argv=command(plan,job,build);started=datetime.datetime.now(datetime.timezone.utc).isoformat();start=time.monotonic()
        with (P/(name+'.stdout')).open('xb') as out,(P/(name+'.stderr')).open('xb') as err:r=subprocess.run(argv,cwd=C.ROOT,env=env,stdout=out,stderr=err)
        seconds=time.monotonic()-start;assert C.census()==source and fixture(job['corpus'])==fixture_before and C.sha(binary)==build['binary']['sha256']
        value=dict(job=job,command=argv,exit_code=r.returncode,started_utc=started,seconds=seconds,source_before=digest(before),source_after=digest(source),source_manifest_sha256=C.sha(P/'source.json'),build_sha256=C.sha(P/'build.json'),plan_sha256=C.sha(P/'plan.json'),script_sha256=C.sha(P/'capture.py'),tools_sha256=C.sha(P/'tool-identities.json'),binary=build['binary'],fixture_before=fixture_before,fixture_after=fixture(job['corpus']),environment=environment,artifacts={f.name:C.sha(f) for f in sorted(P.glob(name+'.*')) if f.is_file()})
        (P/(name+'.receipt.json')).write_text(json.dumps(value,indent=2)+'\n');print(name,'exit',r.returncode,flush=True);assert r.returncode==0
if __name__=='__main__':main()
