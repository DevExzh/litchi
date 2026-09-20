#!/usr/bin/env python3
"""Recompute complete per-process distributions; diagnostics never approve retention."""
import hashlib,json,math,statistics
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def stats(v):
 v=sorted(v);return dict(n=len(v),p50=statistics.median(v),mean=statistics.mean(v),p95=v[math.ceil(.95*len(v))-1],p99=v[math.ceil(.99*len(v))-1],minimum=v[0],maximum=v[-1])
def gate(metric,a,b,plan):
 percent=(b/a-1)*100 if a else (0 if not b else None);delta=b-a;exception=metric in plan['hard_gates']['absolute_exception_metrics'] and delta<=plan['hard_gates']['timing_absolute_ns'];passed=(percent is not None and percent<=plan['hard_gates']['timing_percent']) or exception
 return dict(baseline=a,candidate=b,delta_ns=delta,percent=percent,pass_=passed)
def main():
 plan=read(P/'plan.json');f=read(P/'freeze.json');m=read(P/'captures/manifest.json');assert m['status']=='complete' and m['freeze_sha256']==sha(P/'freeze.json');assert m['bindings_start']==m['bindings_end']==f['bindings']
 for rel,digest in f['bindings'].items():assert sha(ROOT/rel)==digest,rel
 for b in f['binaries']:
  path=Path(b['binary'])
  if path.exists():assert sha(path)==b['binary_sha256'] and path.stat().st_size==b['bytes']
  else:
   c=read(P/'cleanup.json');assert c['removed'];assert any(x['path']==str(path) and x['sha256']==b['binary_sha256'] and x['bytes']==b['bytes'] for x in c['identities'])
 expected={(cy,leg,c['id'],r) for cy in range(plan['cycles']) for leg in plan['legs'] for c in plan['cases'] for r in range(plan['processes_per_cell_leg'])};assert len(m['runs'])==plan['total_processes']==len(expected)
 rows=[];values={};outcomes={}
 for run in m['runs']:
  key=(run['cycle'],run['leg'],run['case'],run['replicate']);assert key in expected and key not in values;assert run['exit_code']==0
  path=P/'captures'/run['output'];assert sha(path)==run['sha256'] and sha(P/'captures'/run['stderr'])==run['stderr_sha256'];x=read(path);c=next(c for c in plan['cases'] if c['id']==run['case'])
  phase='candidate' if run['leg'] in ['b1','b2'] else 'baseline';assert run['phase']==phase
  binary=next(b['binary'] for b in f['binaries'] if b['phase']==phase)
  command=['taskset','-c',str(plan['cpu']),binary,'--input',c['path'],'--budget',str(c['budget']),'--mode',c['mode'],'--worksheet',str(c['sheet']),'--row',str(c['row']),'--column',str(c['column']),'--queries',str(plan['queries']),'--warmups',str(plan['warmups']),'--samples',str(plan['samples'])]
  assert run['command']==command
  assert x['schema_version']==1 and x['probe']=='change-0686-xls-index-budget-retry' and x['input_path']==str(ROOT/c['path'])
  assert x['input_sha256']==sha(ROOT/c['path']) and x['input_bytes']==(ROOT/c['path']).stat().st_size
  assert (x['mode'],x['worksheet'],x['row'],x['column'],x['max_query_index_bytes'],x['queries'],x['warmups'],x['samples'],x['fresh_owner_per_sample'])==(c['mode'],c['sheet'],c['row'],c['column'],c['budget'],plan['queries'],plan['warmups'],plan['samples'],True)
  assert len(x['records'])==plan['samples'];arrays={k:[] for k in plan['timing_metrics']}
  for n,record in enumerate(x['records']):
   assert record['sample']==plan['warmups']+n and record['open']['outcome']['status']=='ok' and record['all_queries_agree']
   qs=record['queries'];assert [q['ordinal'] for q in qs]==list(range(plan['queries']))
   for q in qs:assert q['outcome']['status']==c['expected_status'] and q['outcome']==outcomes.setdefault(c['id'],q['outcome']) and q['agrees_with_first']
   ts=[q['elapsed_ns'] for q in qs];op=record['open']['elapsed_ns'];assert all(isinstance(t,int) and t>=0 for t in ts+[op]);v=[op,ts[0],ts[1],ts[2],ts[7],sum(ts[2:])/6,op+sum(ts)]
   for k,t in zip(plan['timing_metrics'],v):arrays[k].append(t)
  val={k:stats(v) for k,v in arrays.items()};values[key]=val;rows.append(dict(cycle=key[0],leg=key[1],case=key[2],replicate=key[3],metrics=val))
 assert set(values)==expected
 comparisons=[]
 for cy in range(plan['cycles']):
  for c in plan['cases']:
   for r in range(plan['processes_per_cell_leg']):
    for pair,a,b in [('aa2/aa1','aa1','aa2'),('b1/a1','a1','b1'),('b2/a2','a2','b2'),('a2/a1','a1','a2'),('b2/b1','b1','b2')]:
     for metric in plan['timing_metrics']:
      av,bv=values[(cy,a,c['id'],r)][metric],values[(cy,b,c['id'],r)][metric];comparisons.append(dict(cycle=cy,case=c['id'],replicate=r,pair=pair,metric=metric,role=c['role'],focus=metric==c['focus_metric'],statistics={s:gate(metric,av[s],bv[s],plan) for s in ['p50','mean','p95','p99','maximum']}))
 paired=[r for r in comparisons if r['pair'] in ['b1/a1','b2/a2']]
 failed=[dict(cycle=r['cycle'],case=r['case'],replicate=r['replicate'],pair=r['pair'],metric=r['metric'],statistic=s,**r['statistics'][s]) for r in paired for s in ['p50','mean'] if not r['statistics'][s]['pass_']]
 summary=dict(processes=len(rows),owners=len(rows)*plan['samples'],queries=len(rows)*plan['samples']*plan['queries'],failed_central_checks=len(failed),failed_focus_checks=sum(not r['statistics'][s]['pass_'] for r in paired if r['focus'] and r['role']=='failed' for s in ['p50','mean']),paired_tail_flags=sum(r['statistics'][s]['percent']>5 for r in paired for s in ['p95','p99','maximum']))
 result=dict(disposition=plan['disposition'],summary=summary,processes=rows,comparisons=comparisons,failed_checks=failed,outcomes=outcomes);(P/'analysis.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(summary))
if __name__=='__main__':main()
