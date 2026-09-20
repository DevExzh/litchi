#!/usr/bin/env python3
"""Independent custody, outcome and raw statistic audit; no retention decision."""
import hashlib,json,statistics,math,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def stats(xs):
 ys=sorted(xs);return dict(n=len(ys),p50=statistics.median(ys),mean=statistics.mean(ys),p95=ys[math.ceil(len(ys)*.95)-1],p99=ys[math.ceil(len(ys)*.99)-1],minimum=ys[0],maximum=ys[-1])
plan=read(P/'plan.json');freeze=read(P/'freeze.json');m=read(P/'captures/manifest.json')
assert m['status']=='complete' and m['freeze_sha256']==sha(P/'freeze.json')
assert m['bindings_start']==m['bindings_end']==freeze['bindings']
for rel,d in freeze['bindings'].items():assert sha(ROOT/rel)==d,rel
for rel,d in read(P/'constraints.json').items():assert sha(ROOT/rel)==d,rel
cleanup=read(P/'cleanup.json') if (P/'cleanup.json').exists() else None
for b in freeze['binaries']:
 path=Path(b['binary'])
 if path.exists():assert sha(path)==b['binary_sha256'] and path.stat().st_size==b['bytes']
 else:assert cleanup and any(v['path']==str(path) and v['sha256']==b['binary_sha256'] and v['bytes']==b['bytes'] for v in cleanup['identities'])
subprocess.run(['python3','-B',str(P/'source-guard.py')],cwd=ROOT,check=True)
expected=[]
for s in plan['stages']:
 for c in plan['cases']:
  expected.append((s['stage'],s['variant'],c['case'],'native',None))
  expected.extend((s['stage'],s['variant'],c['case'],'repeat',i) for i in range(9))
assert [(r['stage'],r['variant'],r['case'],r['lane'],r['sample']) for r in m['runs']]==expected
seen={};outcomes={};result={};rawfiles={'manifest.json'}
for r in m['runs']:
 expected_name=f"{r['stage']}-{r['case']}-{r['lane']}"+(f"-{r['sample']}.tsv" if r['lane']=='repeat' else '.json')
 assert r['output']==expected_name and r['stderr']==expected_name+'.stderr'
 c=next(c for c in plan['cases'] if c['case']==r['case']);path=P/'captures'/r['output'];stderr=P/'captures'/r['stderr']
 assert path.resolve().parent==(P/'captures').resolve() and stderr.resolve().parent==(P/'captures').resolve()
 assert r['exit_code']==0 and sha(path)==r['sha256'] and sha(stderr)==r['stderr_sha256'];rawfiles.update((path.name,stderr.name))
 binary=next(b['binary'] for b in freeze['binaries'] if b['variant']==r['variant'] and Path(b['binary']).name==plan['binaries'][r['lane']])
 cmd=['taskset','-c',str(plan['cpu']),binary]
 if r['lane']=='native':cmd+=['--input',c['path'],'--budget',str(c['budget']),'--mode','owned','--worksheet',str(c['sheet']),'--row',str(c['row']),'--column',str(c['column']),'--queries','8','--warmups','3','--samples','100']
 else:cmd+=['owned',c['path'],str(c['sheet']),str(c['row']),str(c['column']),'50000']
 assert r['command']==cmd
 key=r['stage']+'/'+r['case'];v=result.setdefault(key,{})
 if r['lane']=='native':
  x=read(path);assert x['probe']==plan['native']['probe'] and x['queries']==8 and x['warmups']==3 and x['samples']==100 and x['fresh_owner_per_sample'] is True and x['all_queries_agree'] is True
  assert len(x['records'])==100 and x['input_sha256']==sha(ROOT/c['path'])
  assert x['mode']=='owned' and x['max_query_index_bytes']==c['budget'] and x['row']==c['row'] and x['column']==c['column'] and x['worksheet']==c['sheet']
  vectors={k:[] for k in ('open','q1','q2','q3','q8','q3-to-q8-mean','open-plus-eight')}
  for rec in x['records']:
   assert rec['open']['outcome']['status']=='ok';qs=rec['queries'];assert [q['ordinal'] for q in qs]==list(range(8))
   for q in qs:
    assert q['agrees_with_first'] is True
    assert q['outcome']['status']==c['expected_status']
    prior=outcomes.setdefault(c['case'],q['outcome']);assert prior==q['outcome']
   t=[q['elapsed_ns'] for q in qs];o=rec['open']['elapsed_ns'];assert min(t+[o])>=0
   vals=dict(open=o,q1=t[0],q2=t[1],q3=t[2],q8=t[7],**{'q3-to-q8-mean':sum(t[2:])/6,'open-plus-eight':o+sum(t)})
   for k,n in vals.items():vectors[k].append(n)
  v['native']={k:stats(ns) for k,ns in vectors.items()}
 else:
  x=dict(f.split('=') for f in path.read_text().strip().split('\t'));assert set(x)=={'repeats','found','nanos'} and int(x['repeats'])==50000
  assert int(x['found'])==(0 if c['expected_status']=='missing' else 50000)
  v.setdefault('repeat_raw',[]).append(int(x['nanos'])/50000)
assert rawfiles=={p.name for p in (P/'captures').iterdir() if p.is_file()}
for v in result.values():assert len(v['repeat_raw'])==9;v['repeat']=stats(v.pop('repeat_raw'))
report=dict(status='verified diagnostic evidence; no retention decision',runs=len(m['runs']),native_owners=4800,native_queries=38400,repeat_processes=432,repeat_queries=21600000,stages=result)
(P/'audit.json').write_text(json.dumps(report,indent=2)+'\n');print('PASS 480 runs, exact custody/outcomes; independent raw statistics written')
