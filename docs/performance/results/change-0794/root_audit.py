"""Independent numerical audit from raw reports; no benchmark execution."""
import json,math,random,statistics,sys
from pathlib import Path
P=Path(__file__).resolve().parent
read=lambda p:json.loads(p.read_text())
plan=read(P/'plan.json');policy=read(P/'adoption-policy.json')
def median(x):return statistics.median(x)
def q(x,f):return sorted(x)[max(0,math.ceil(len(x)*f)-1)]
rows=[];violations=[];benefits=[];memory=[]
for case in plan['cases']:
 shape,mode=case['shape'],case['mode'];p50={}
 for leg in ['before','after']:
  p50[leg]=[q([s['elapsed_ns'] for s in read(P/f'native/{b}-{shape}-{mode}-{leg}.json')['samples']],.5) for b in range(6)]
 ratios=[a/b for a,b in zip(p50['after'],p50['before'])]
 rng=random.Random(policy['latency']['seed']);dist=sorted(median([ratios[rng.randrange(6)] for _ in ratios]) for _ in range(10000))
 ci=[dist[250],dist[9749]];ratio=median(ratios)
 row={**case,'before_p50_ns':median(p50['before']),'after_p50_ns':median(p50['after']),'paired_ratios':ratios,'ratio':ratio,'ci95':ci}
 rows.append(row)
 if ratio>1.05 and ci[0]>1:violations.append(row)
 if mode in ['capture','lifecycle'] and ratio<=.97 and ci[1]<1:benefits.append(row)
 for block in range(2):
  v={}
  for leg in ['before','after']:
   a=[s['allocation'] for s in read(P/f'allocation/{block}-{shape}-{mode}-{leg}.json')['samples']]
   v[leg]={'calls':median(x['allocation_calls'] for x in a),'bytes':median(x['allocated_bytes'] for x in a),'net':median(x['live_bytes_after']-x['live_bytes_before'] for x in a),'peak':median(x['region_peak_live_bytes']-x['live_bytes_before'] for x in a)}
  memory.append({**case,'block':block,**v,'violations':[k for k in v['before'] if v['after'][k]>v['before'][k]]})
result={'native':rows,'latency_violations':violations,'benefits':benefits,'memory':memory,'memory_guard':not any(x['violations'] for x in memory),'adoption_eligible':bool(benefits) and not violations and not any(x['violations'] for x in memory),'scope':'Independent raw numerical audit only; main validator checks custody and identities.'}
out=P/'root-audit.json'
if '--check' in sys.argv:assert read(out)==result;print('Independent raw numerical audit PASS')
else:assert not out.exists();out.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n');print(json.dumps({'adoption_eligible':result['adoption_eligible'],'benefits':len(benefits),'latency_violations':len(violations),'memory_guard':result['memory_guard']}))
