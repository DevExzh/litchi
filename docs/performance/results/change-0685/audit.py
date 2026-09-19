#!/usr/bin/env python3
"""Verify 0685 captures and expose every per-phase regression and control drift."""
import hashlib
import json
import math
from pathlib import Path
import statistics
import subprocess

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def stat(v):
    v=sorted(v)
    return dict(n=len(v),p50=statistics.median(v),mean=statistics.mean(v),p95=v[math.ceil(.95*len(v))-1],p99=v[-1])
def delta(a,b):return (a/b-1)*100 if b else None
m={}
for phase in ['baseline','candidate']:
    d=P/'measurements'/phase;m[phase]=read(d/'manifest.json')
    for name,value in m[phase]['raw_sha256'].items():assert sha(d/name)==value,(phase,name)
    for name,value in m[phase]['probe_sha256'].items():assert sha(ROOT/name)==value,(phase,name)
    assert sha(P/'cases.json')==m[phase]['cases_sha256']
    assert sha(P/'corpus-manifest.json')==m[phase]['corpus_manifest_sha256']
assert m['baseline']['probe_sha256']==m['candidate']['probe_sha256']
for path,value in m['candidate']['source_sha256'].items():assert sha(ROOT/path)==value,path
for path,value in read(P/'baseline.json')['constraints_sha256'].items():assert sha(ROOT/path)==value,path
for item in read(P/'corpus-manifest.json'):assert sha(ROOT/item['path'])==item['sha256']
quality=read(P/'final-verified/results.json')
assert {v['name'] for v in quality}=={'fmt','check','clippy','tests','facade','rustdoc'}
for item in quality:
    assert item['exit_code']==0,item['name']
    for path,value in item['source_sha256'].items():assert sha(ROOT/path)==value,(item['name'],path)
B=P/'measurements/baseline';A=P/'measurements/candidate'
for mode in ['owned','file']:
    before=read(B/f'corpus-{mode}.json');after=read(A/f'corpus-{mode}.json')
    assert before==after,mode
    assert after['files_seen']==126 and after['query_mismatches']==0
    assert all(q['agrees'] for f in after['files'] for s in f['worksheets'] for q in s['queries'])

def semantic(record):
    visit=record['visit']
    return (record['open']['outcome'],[(q['label'],q['outcome']) for q in record['query_phases']],
        (visit['outcome'],visit['actual_callbacks'],visit['oracle_callbacks'],visit['oracle_digest']) if visit else None)
def samples(path):
    report=read(path);records=report['records'];n=m['baseline']['samples'];w=m['baseline']['warmups']
    assert len(records)==n and [r['sample'] for r in records]==list(range(w,w+n)),path
    sem=semantic(records[0]);assert all(semantic(r)==sem for r in records),path
    metrics={'open':[r['open']['elapsed_ns'] for r in records]}
    if records[0]['visit']:
        metrics['visit']=[r['visit']['elapsed_ns'] for r in records]
    else:
        for i,q in enumerate(records[0]['query_phases']):metrics[q['label']]=[r['query_phases'][i]['elapsed_ns'] for r in records]
        for count in [2,3]:
            metrics[f'queries-{count}']=[sum(q['elapsed_ns'] for q in r['query_phases'][:count]) for r in records]
            metrics[f'open-and-queries-{count}']=[r['open']['elapsed_ns']+sum(q['elapsed_ns'] for q in r['query_phases'][:count]) for r in records]
    return {key:stat(values) for key,values in metrics.items()},sem
results=[]
for path in sorted(B.glob('a1-*.json')):
    group=path.name[3:-5];legs={};sems=[]
    for label,folder,leg in [('aa1',B,'a1'),('aa2',B,'a2'),('a1',A,'a1'),('b1',A,'b1'),('b2',A,'b2'),('a2',A,'a2')]:
        legs[label],sem=samples(folder/f'{leg}-{group}.json');sems.append(sem)
    assert all(s==sems[0] for s in sems),group
    for metric in legs['aa1']:
        t={label:leg[metric] for label,leg in legs.items()}
        percentages={s:dict(b1_a1=delta(t['b1'][s],t['a1'][s]),b2_a2=delta(t['b2'][s],t['a2'][s]),aa=delta(t['aa2'][s],t['aa1'][s]),abba_a=delta(t['a2'][s],t['a1'][s]),abba_b=delta(t['b2'][s],t['b1'][s])) for s in ['p50','mean','p95','p99']}
        flag=all(percentages['p50'][key]>5 for key in ['b1_a1','b2_a2'])
        results.append(dict(group=group,metric=metric,timing=t,percent=percentages,review_regression=flag))
alloc=[]
metrics=['allocation_calls','allocated_bytes','deallocation_calls','deallocated_bytes','peak_live_delta','retained_live_delta','outcome','callbacks']
for path in sorted(B.glob('alloc-*-0.json')):
    group=path.name[6:-7];rows={}
    for phase,folder in [('before',B),('after',A)]:
        repeated=[read(folder/f'alloc-{group}-{i}.json') for i in range(3)]
        for key in metrics:assert all(v[key]==repeated[0][key] for v in repeated),(phase,group,key)
        rows[phase]={key:repeated[0][key] for key in metrics}
    for key in ['outcome','callbacks']:assert rows['before'][key]==rows['after'][key],(group,key)
    alloc.append(dict(group=group,**rows))
assert len(alloc)==88,len(alloc)
assert len(list(B.glob('a1-*.json')))==48
output=dict(route_groups=48,native_route_samples=48*30*6,allocation_samples=88*3*2,phases=results,allocations=alloc)
(P/'comparison.json').write_text(json.dumps(output,indent=2)+'\n')
print('Verified source/probe/constraints/corpus/quality/raw bindings, 126-fixture corpus parity, 48 route groups and 88 allocation groups')
for r in results:
    if r['review_regression']:print('REVIEW paired >5% p50:',r['group'],r['metric'],r['percent']['p50'])

for name,h in m['baseline']['source_sha256'].items():
    data=subprocess.check_output(['git','show',m['baseline']['head']+':'+name],cwd=ROOT)
    assert hashlib.sha256(data).hexdigest()==h,name
for phase,directory in [('baseline','before'),('candidate','after')]:
    for binary,key in [('xls-index-probe-0684','binary_sha256'),('xls0684-alloc','allocator_binary_sha256')]:
        path=Path('/home/zhuhe/code/litchi-target-0685-'+directory)/'release'/binary
        if path.exists():assert sha(path)==m[phase][key],str(path)
print('Baseline source matches immutable git objects; available binaries match manifests')
def untimed(value):
    if isinstance(value,dict):return {k:untimed(v) for k,v in value.items() if k not in {'elapsed_ns','total_elapsed_ns'}}
    if isinstance(value,list):return [untimed(v) for v in value]
    return value
for path in sorted(B.glob('counts-*.json')):
    assert untimed(read(path))==untimed(read(A/path.name)),path.name
assert len(list(B.glob('counts-*.json')))==24
for item in read(P/'final-builds.json'):
    assert item['exit_code']==0
    for name,h in item['source_sha256'].items():assert sha(ROOT/name)==h,name
print('All 24 counted routes preserve source I/O/fences/outcomes; final builds match live owner sources')
D=P/'timing-diagnostics';manifest=read(D/'manifest.json')
for name,h in manifest['raw_sha256'].items():assert sha(D/name)==h,name
for key in ['source_sha256','probe_sha256']:
    for name,h in manifest[key].items():assert sha(ROOT/name)==h,name
assert manifest['before_binary_sha256']==m['baseline']['binary_sha256']
assert manifest['after_binary_sha256']==m['candidate']['binary_sha256']
assert manifest['cases_sha256']==sha(P/'cases.json')
long=[]
for path in sorted(D.glob('a1-*.json')):
    group=path.name[3:-5];legs={};sems=[]
    for leg in ['a1','b1','b2','a2']:
        records=read(D/f'{leg}-{group}.json')['records']
        assert len(records)==200 and [r['sample'] for r in records]==list(range(10,210))
        sem=semantic(records[0]);assert all(semantic(r)==sem for r in records);sems.append(sem)
        metrics={'open':[r['open']['elapsed_ns'] for r in records]}
        if records[0]['visit']:metrics['visit']=[r['visit']['elapsed_ns'] for r in records]
        else:
            for i,q in enumerate(records[0]['query_phases']):metrics[q['label']]=[r['query_phases'][i]['elapsed_ns'] for r in records]
            metrics['open-and-queries-3']=[r['open']['elapsed_ns']+sum(q['elapsed_ns'] for q in r['query_phases']) for r in records]
        legs[leg]={k:stat(v) for k,v in metrics.items()}
    assert all(sem==sems[0] for sem in sems)
    for metric in legs['a1']:
        t={label:leg[metric] for label,leg in legs.items()}
        changes={key:delta(t[a]['p50'],t[b]['p50']) for key,a,b in [('b1_a1','b1','a1'),('b2_a2','b2','a2'),('abba_a','a2','a1'),('abba_b','b2','b1')]}
        long.append(dict(group=group,metric=metric,timing=t,percent_p50=changes))
assert len(list(D.glob('a1-*.json')))==6
(P/'timing-diagnostics-comparison.json').write_text(json.dumps(long,indent=2)+'\n')
print('Verified 4,800 longer diagnostic route records and bindings')
for r in long:
    if r['metric'] in ['open','q1','q2-build-trigger','q3-warm-repeat-q1','open-and-queries-3']:print('LONG',r['group'],r['metric'],r['percent_p50'])
I=P/'initial-candidate'
for folder in ['measurements/candidate','diagnostics/candidate','timing-diagnostics']:
    d=I/folder;manifest=read(d/'manifest.json')
    for name,h in manifest['raw_sha256'].items():assert sha(d/name)==h,(folder,name)
    for name,h in manifest['source_sha256'].items():
        archived=I/name
        assert sha(archived if archived.exists() else ROOT/name)==h,name
    for name,h in manifest['probe_sha256'].items():assert sha(ROOT/name)==h,name
for item in read(I/'final-verified/results.json'):
    assert item['exit_code']==0
    for name,h in item['source_sha256'].items():
        archived=I/name
        assert sha(archived if archived.exists() else ROOT/name)==h,name
for name,h in read(I/'profile-manifest.json').items():assert sha(I/name)==h,name
print('Verified initial-candidate archive, quality, sources, raw captures and profile text')
