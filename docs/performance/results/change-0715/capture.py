#!/usr/bin/env python3
"""Capture the frozen native phase matrix, then separately traced children."""
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('custody0715capture', P/'custody.py')
C = importlib.util.module_from_spec(spec)
spec.loader.exec_module(C)

def read(path):
    return json.loads(path.read_text())

def write(path, value):
    assert not path.exists(), path
    path.write_text(json.dumps(value, indent=2)+'\n')

def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(',', ':')).encode()).hexdigest()

def jobs(plan, lane):
    rows = []
    for repeat in range(1, plan[lane]['repeats']+1):
        phases, corpora = list(plan[lane]['phases']), list(plan['corpora'])
        if repeat == 2:
            phases.reverse()
            corpora.reverse()
        for phase in phases:
            for corpus in corpora:
                prefix = 'docx_ordinary_save_' if corpus['id'] == 'generated' else 'docx_real_file_ordinary_save_'
                rows.append(dict(lane=lane, repeat=repeat, phase=phase, corpus=corpus,
                                 name=f'{lane}-r{repeat}-{corpus["id"]}-{phase}',
                                 case=prefix+phase, samples=plan[lane]['samples'],
                                 warmup=plan[lane]['warmup'], order_index=len(rows)))
    assert len(rows) == plan['expected_children'][lane]
    return rows

def command(plan, job, build):
    result = ['taskset', '-c', str(plan['cpu'])]
    if job['lane'] == 'profile':
        result += [plan['profile']['tool'], *plan['profile']['options'], '--callgrind-out-file='+str(P/(job['name']+'.callgrind'))]
    result += [build['binary']['path'], '--warmup', str(job['warmup']),
               '--samples', str(job['samples']), '--case', job['case'],
               '--json', str(P/(job['name']+'.json')),
               '--filesystem-root', plan['filesystem_root']]
    if job['corpus']['path']:
        result += ['--ooxml-file', job['corpus']['path']]
    return result

def fixture(corpus):
    if corpus['path'] is None:
        return None
    path = C.ROOT/corpus['path']
    value = dict(path=corpus['path'], bytes=path.stat().st_size, sha256=C.sha(path))
    assert value['bytes'] == corpus['bytes'] and value['sha256'] == corpus['sha256']
    return value

def main():
    lane = sys.argv[1]
    assert lane == 'profile'
    plan, build, source = read(P/'plan.json'), read(P/'build.json'), read(P/'source.json')
    assert build['exit_code'] == 0 and build['source_sha256'] == C.sha(P/'source.json')
    for name, sha in read(P/'capture-freeze.json').items():
        assert C.sha(P/name) == sha
    for name, sha in read(P/'constraints.json').items():
        assert C.sha(C.ROOT/name) == sha
    binary = Path(build['binary']['path'])
    assert binary.stat().st_size == build['binary']['bytes'] and C.sha(binary) == build['binary']['sha256']
    env = dict(os.environ)
    env.update(LC_ALL='C', LANG='C', TZ='UTC', PERL_HASH_SEED='0', PERL_PERTURB_KEYS='0')
    for job in jobs(plan, lane):
        name = job['name']
        assert not list(P.glob(name+'.*')), name
        before = C.census()
        assert before == source
        fixture_before = fixture(job['corpus'])
        argv = command(plan, job, build)
        started = datetime.datetime.now(datetime.timezone.utc).isoformat()
        start = time.monotonic()
        with (P/(name+'.stdout')).open('wb') as stdout, (P/(name+'.stderr')).open('wb') as stderr:
            result = subprocess.run(argv, cwd=C.ROOT, env=env, stdout=stdout, stderr=stderr)
        after, fixture_after = C.census(), fixture(job['corpus'])
        assert after == before == source and fixture_after == fixture_before
        assert C.sha(binary) == build['binary']['sha256']
        record = dict(job=job, command=argv, exit_code=result.returncode, started_utc=started,
                      seconds=time.monotonic()-start, binary=build['binary'],
                      source_before=digest(before), source_after=digest(after),
                      source_manifest_sha256=C.sha(P/'source.json'),
                      build_sha256=C.sha(P/'build.json'), plan_sha256=C.sha(P/'plan.json'),
                      script_sha256=C.sha(P/'capture.py'), fixture_before=fixture_before,
                      fixture_after=fixture_after,
                      environment={k: env.get(k) for k in ['LC_ALL', 'LANG', 'TZ', 'RUSTFLAGS', 'LD_PRELOAD', 'PERL_HASH_SEED', 'PERL_PERTURB_KEYS']},
                      artifacts={f.name:C.sha(f) for f in sorted(P.glob(name+'.*')) if f.is_file()})
        write(P/(name+'.receipt.json'), record)
        assert result.returncode == 0, name
        print(name, 'PASS', flush=True)

if __name__ == '__main__':
    main()
