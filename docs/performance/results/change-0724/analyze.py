#!/usr/bin/env python3
"""Recompute every stage and mirrored contrast; diagnostic evidence, no retention gate."""
import hashlib,json,math,statistics
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def load(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def stats(v):
 v=sorted(v);return dict(n=len(v),p50=statistics.median(v),mean=statistics.mean(v),p95=v[math.ceil(.95*len(v))-1],p99=v[math.ceil(.99*len(v))-1],minimum=v[0],maximum=v[-1])
def delta(a,b):return {k:(b[k]/a[k]-1)*100 if a[k] else None for k in ('p50','mean','p95','p99','minimum','maximum')}
def main():
 plan=load(P/'plan.json');f=load(P/'freeze.json');m=load(P/'captures/manifest.json')
 assert m['status']=='complete' and m['freeze_sha256']==sha(P/'freeze.json')
 assert m['bindings_start']==m['bindings_end']==f['bindings']
 for rel,h in f['bindings'].items():assert sha(ROOT/rel)==h,rel
 assert len(m['runs'])==plan['capture']['runs']
 result={};outcomes={};keys=set()
 for run in m['runs']:
  key=(run['stage'],run['case'],run['lane'],run['sample']);assert key not in keys;keys.add(key)
  assert run['exit_code']==0
  path=P/'captures'/run['output'];assert sha(path)==run['sha256'];assert sha(P/'captures'/run['stderr'])==run['stderr_sha256']
  c=next(c for c in plan['cases'] if c['case']==run['case']);v=result.setdefault(run['stage']+'/'+run['case'],{})
  if run['lane']=='native':
   x=load(path);assert x['input_sha256']==sha(ROOT/c['path']) and len(x['records'])==100
   arrays={k:[] for k in plan['native']['timing_metrics']}
   for record in x['records']:
    assert record['open']['outcome']['status']=='ok';qs=record['queries'];assert [q['ordinal'] for q in qs]==list(range(8))
    for q in qs:
     assert q['outcome']['status']==c['expected_status'] and q['outcome']==outcomes.setdefault(c['case'],q['outcome'])
    times=[q['elapsed_ns'] for q in qs];op=record['open']['elapsed_ns'];assert min(times+[op])>=0
    vals=[op,times[0],times[1],times[2],times[7],sum(times[2:])/6,op+sum(times)]
    for metric,value in zip(plan['native']['timing_metrics'],vals):arrays[metric].append(value)
   v['native']={k:stats(a) for k,a in arrays.items()}
  else:
   parts=dict(t.split('=') for t in path.read_text().strip().split('\t'));assert set(parts)=={'repeats','found','nanos'}
   assert int(parts['repeats'])==50000 and int(parts['found'])==(0 if c['expected_status']=='missing' else 50000)
   v.setdefault('samples',[]).append(int(parts['nanos'])/50000)
 assert len(result)==48
 for v in result.values():assert len(v['samples'])==9;v['repeat']=stats(v.pop('samples'))
 contrasts=[];drift=[]
 stages={(s['variant'],s['round']):s['stage'] for s in plan['stages']}
 def contrast(case,before,after):
  a=result[before+'/'+case];b=result[after+'/'+case]
  return dict(case=case,before=before,after=after,native={k:delta(a['native'][k],b['native'][k]) for k in a['native']},repeat=delta(a['repeat'],b['repeat']))
 for c in plan['cases']:
  for effect in plan['attribution']['effects']:
   for rd in (1,2):contrasts.append(dict(effect=effect['name'],round=rd,**contrast(c['case'],stages[effect['before'],rd],stages[effect['after'],rd])))
  for variant in ('baseline','layout','selection','full'):drift.append(dict(variant=variant,**contrast(c['case'],stages[variant,1],stages[variant,2])))
 report=dict(status='verified diagnostic evidence; no retention decision',stages=result,contrasts=contrasts,drift=drift)
 (P/'analysis.json').write_text(json.dumps(report,indent=2)+'\n')
 lines=['# Four-variant attribution','','All deltas are after/before minus one. No retention gate is applied. Repeat tails are percentiles of nine process loop means, not individual queries.','','| Case | Effect | Round | Repeat p50 | Repeat mean | q2 p50 | q2 mean | q8 p50 |','| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |']
 for r in contrasts:lines.append(f"| {r['case']} | {r['effect']} | {r['round']} | {r['repeat']['p50']:+.2f}% | {r['repeat']['mean']:+.2f}% | {r['native']['q2']['p50']:+.2f}% | {r['native']['q2']['mean']:+.2f}% | {r['native']['q8']['p50']:+.2f}% |")
 lines+=['','## Mirrored stage drift','','| Case | Variant | Repeat p50 | Repeat mean | q2 p50 | q2 mean |','| --- | --- | ---: | ---: | ---: | ---: |']
 for r in drift:lines.append(f"| {r['case']} | {r['variant']} | {r['repeat']['p50']:+.2f}% | {r['repeat']['mean']:+.2f}% | {r['native']['q2']['p50']:+.2f}% | {r['native']['q2']['mean']:+.2f}% |")
 lines+=['','Complete per-stage p50/mean/p95/p99/min/max and every metric contrast remain in analysis.json.']
 (P/'analysis.md').write_text('\n'.join(lines)+'\n');print('PASS 48 native and 48 repeat groups; diagnostic contrasts written')
if __name__=='__main__':main()
