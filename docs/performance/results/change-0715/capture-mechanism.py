#!/usr/bin/env python3
"""Capture candidate attribution and per-child RSS after an accepted pilot."""
import datetime,importlib.util,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
s=importlib.util.spec_from_file_location('capture0715mechanism',P/'capture.py');C=importlib.util.module_from_spec(s);s.loader.exec_module(C)
def jobs(plan):
    rows=[]
    for repeat in [1,2]:
        corpora=list(plan['corpora']);corpora.reverse() if repeat==2 else None
        for corpus in corpora:
            rows.append(dict(kind='profile',source='candidate',repeat=repeat,corpus=corpus,phase='counting_publish',samples=1,warmup=0))
    for source in ['baseline','candidate']:
        for repeat in [1,2]:
            corpora=list(plan['corpora']);phases=['counting_publish','lifecycle']
            if repeat==2:corpora.reverse();phases.reverse()
            for phase in phases:
                for corpus in corpora:rows.append(dict(kind='rss',source=source,repeat=repeat,corpus=corpus,phase=phase,samples=10,warmup=2))
    for row in rows:
        row['name']=f"mechanism-{row['source']}-{row['kind']}-r{row['repeat']}-{row['corpus']['id']}-{row['phase']}"
        row['case']=('docx_ordinary_save_' if row['corpus']['id']=='generated' else 'docx_real_file_ordinary_save_')+row['phase']
    return rows

def binary(job):
    row=next(r for r in C.read(P/('build-'+job['source']+'.json')) if r['lane']=='native')
    return dict(path=row['binary'],sha256=row['binary_sha256'],bytes=row['binary_bytes'])

def command(plan,job,b):
    prefix=['taskset','-c',str(plan['cpu'])]
    if job['kind']=='profile':prefix += ['/usr/bin/valgrind',*plan['profile']['options'],'--callgrind-out-file='+str(P/(job['name']+'.callgrind'))]
    else:prefix += ['/usr/bin/time','-v','-o',str(P/(job['name']+'.time-v'))]
    prefix += [b['path'],'--warmup',str(job['warmup']),'--samples',str(job['samples']),'--case',job['case'],'--json',str(P/(job['name']+'.json')),'--filesystem-root',plan['filesystem_root']]
    if job['corpus']['path']:prefix += ['--ooxml-file',job['corpus']['path']]
    return prefix

def main():
    plan=C.read(P/'plan.json');candidate=C.read(P/'source-candidate.json')
    decision=C.read(P/'pilot-analysis.json')['decision'];assert decision['accepted']
    for n,h in C.read(P/'mechanism-freeze.json').items():assert C.C.sha(P/n)==h
    env=dict(os.environ,LC_ALL='C',LANG='C',TZ='UTC',PERL_HASH_SEED='0',PERL_PERTURB_KEYS='0')
    for job in jobs(plan):
        name=job['name'];assert not list(P.glob(name+'.*'));assert C.C.census()==candidate
        b=binary(job);assert C.C.sha(Path(b['path']))==b['sha256']
        fixture=C.fixture(job['corpus']);cmd=command(plan,job,b);started=datetime.datetime.now(datetime.timezone.utc).isoformat();t=time.monotonic()
        with (P/(name+'.stdout')).open('wb') as out,(P/(name+'.stderr')).open('wb') as err:r=subprocess.run(cmd,cwd=C.C.ROOT,env=env,stdout=out,stderr=err)
        assert C.C.census()==candidate and C.fixture(job['corpus'])==fixture and C.C.sha(Path(b['path']))==b['sha256']
        C.write(P/(name+'.receipt.json'),dict(job=job,command=cmd,exit_code=r.returncode,started_utc=started,seconds=time.monotonic()-t,binary=b,live_source_sha256=C.C.sha(P/'source-candidate.json'),binary_source_sha256=C.C.sha(P/('source-'+job['source']+'.json')),fixture=fixture,script_sha256=C.C.sha(Path(__file__)),plan_sha256=C.C.sha(P/'plan.json'),pilot_analysis_sha256=C.C.sha(P/'pilot-analysis.json'),artifacts={f.name:C.C.sha(f) for f in sorted(P.glob(name+'.*')) if f.is_file()}))
        assert r.returncode==0;print(name,'PASS',flush=True)
if __name__=='__main__':main()
