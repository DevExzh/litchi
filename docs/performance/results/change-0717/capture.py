#!/usr/bin/env python3
"""Fixed native and opt-in procfs diagnostic capture; never replace a child artifact."""
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import time

P = Path(__file__).resolve().parent

def module(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result

C = module(P/'custody.py', 'custody0717')

def read(path):
    return json.loads(path.read_text())

def write(path, value):
    with path.open('x') as out:
        out.write(json.dumps(value, indent=2)+'\n')

def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(',', ':')).encode()).hexdigest()

def jobs(plan):
    rows=[]
    base=[(corpus,lane) for corpus in plan['corpora'] for lane in plan['lanes']]
    for block in range(plan['blocks']):
        shift=block % len(base)
        for position,(corpus,lane) in enumerate(base[shift:]+base[:shift]):
            prefix='docx_ordinary_save_' if corpus['id']=='generated' else 'docx_real_file_ordinary_save_'
            rows.append(dict(name=f"{lane}-b{block+1}-{corpus['id']}",lane=lane,block=block+1,
                             position=position,order_index=len(rows),corpus=corpus,
                             warmup=plan['warmup'],samples=plan['samples'],phase='counting_publish',
                             case=prefix+'counting_publish'))
    assert len(rows)==plan['expected_children']
    return rows

def command(plan, job, build):
    argv = ['taskset', '-c', str(plan['cpu']), build['binary']['path'],
            '--warmup', str(job['warmup']), '--samples', str(job['samples']),
            '--case', job['case'], '--json', str(P/(job['name']+'.json')),
            '--filesystem-root', plan['filesystem_root']]
    if job['corpus']['path']:
        argv += ['--ooxml-file', job['corpus']['path']]
    return argv

def fixture(corpus):
    if corpus['path'] is None:
        return None
    path=C.ROOT/corpus['path']
    result=dict(path=corpus['path'], bytes=path.stat().st_size, sha256=C.sha(path))
    assert result['bytes']==corpus['bytes'] and result['sha256']==corpus['sha256']
    return result

def context(plan):
    result = dict(monotonic_ns=time.monotonic_ns())
    for name in plan['context_paths']:
        try:
            value=Path(name).read_text()
            if name=='/proc/stat':
                value='\n'.join(line for line in value.splitlines()
                                if line.startswith(('cpu ', f"cpu{plan['cpu']} ", 'ctxt ', 'procs_running ', 'procs_blocked ')))+'\n'
            result[name]=dict(text=value)
        except OSError as error:
            result[name]=dict(unavailable=type(error).__name__, errno=error.errno)
    return result

def main():
    plan, builds, source = [read(P/n) for n in ['plan.json','builds.json','source.json']]
    for n,h in read(P/'capture-freeze.json').items():
        assert C.sha(P/n)==h
    for n,h in read(P/'constraints.json').items():
        assert C.sha(C.ROOT/n)==h
    for n,h in read(P/'helper-freeze.json').items():
        assert C.sha(C.ROOT/n)==h
    assert set(builds)==set(plan['lanes'])
    for build in builds.values():
        assert build['exit_code']==0 and build['source_sha256']==C.sha(P/'source.json')
        binary=Path(build['binary']['path'])
        assert C.sha(binary)==build['binary']['sha256'] and binary.stat().st_size==build['binary']['bytes']
    env=dict(os.environ)
    env.update(plan['environment_overrides'])
    environment={key:env.get(key) for key in plan['environment_keys']}
    assert environment==plan['expected_environment']
    for job in jobs(plan):
        build=builds[job['lane']]
        binary=Path(build['binary']['path'])
        name=job['name']
        # Resume only untouched jobs after validating all earlier completed receipts.
        receipt=P/(name+'.receipt.json')
        if receipt.exists():
            prior=read(receipt)
            assert prior['job']==job and prior['exit_code']==0
            assert prior['plan_sha256']==C.sha(P/'plan.json') and prior['script_sha256']==C.sha(P/'capture.py')
            for n,h in prior['artifacts'].items(): assert C.sha(P/n)==h
            continue
        assert not list(P.glob(name+'.*')), 'partial child must be retained and investigated'
        before=C.census()
        assert before==source
        fixture_before=fixture(job['corpus'])
        argv=command(plan,job,build)
        started=datetime.datetime.now(datetime.timezone.utc).isoformat()
        before_context=context(plan)
        start=time.monotonic()
        with (P/(name+'.stdout')).open('xb') as stdout, (P/(name+'.stderr')).open('xb') as stderr:
            child=subprocess.Popen(argv,cwd=C.ROOT,env=env,stdout=stdout,stderr=stderr)
            pid,status,usage=os.wait4(child.pid,0)
            assert pid==child.pid
            child.returncode=os.waitstatus_to_exitcode(status)
        seconds=time.monotonic()-start
        after_context=context(plan)
        after=C.census()
        assert before==after==source and fixture(job['corpus'])==fixture_before
        assert C.sha(binary)==build['binary']['sha256']
        fields=['ru_utime','ru_stime','ru_maxrss','ru_minflt','ru_majflt','ru_inblock','ru_oublock','ru_nvcsw','ru_nivcsw']
        write(P/(name+'.context.json'),dict(before=before_context,after=after_context,
                                          wait4={k:getattr(usage,k) for k in fields},
                                          scope='Whole child including setup, warmup, untimed opening, verification and reporting; system snapshots are not owner counters.'))
        record=dict(job=job,command=argv,exit_code=child.returncode,started_utc=started,seconds=seconds,
                    binary=build['binary'],source_before=digest(before),source_after=digest(after),
                    fixture_before=fixture_before,fixture_after=fixture(job['corpus']),environment=environment,
                    source_manifest_sha256=C.sha(P/'source.json'),build_sha256=C.sha(P/'builds.json'),
                    plan_sha256=C.sha(P/'plan.json'),script_sha256=C.sha(P/'capture.py'),
                    artifacts={f.name:C.sha(f) for f in sorted(P.glob(name+'.*')) if f.is_file()})
        write(receipt,record)
        assert child.returncode==0, name
        print(name,'PASS',flush=True)

if __name__=='__main__':
    main()
