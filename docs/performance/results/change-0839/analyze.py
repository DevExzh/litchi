"""Frozen paired decision; no historical samples or instrumented timing pooled."""
import driver as d
import qualify as q
import readers as r
import json,statistics as st
P=d.P;plan=d.read(P/'plan.json');rows=q.all_rows()
lanes={lane:[x for x in rows if x['lane']==lane] for lane in ['qualification','native','observer','memory','allocation','memory-preflight','allocation-preflight']}
counts={'qualification':120,'native':720,'observer':240,'memory':192,'allocation':108,'memory-preflight':16,'allocation-preflight':18}
for lane,n in counts.items():assert len(lanes[lane])==n,(lane,len(lanes[lane]),n)
for lane in ['qualification','native','observer','memory','allocation']:
 cases=plan[{'memory':'memory_cases','allocation':'allocation_cases'}.get(lane,'cases')]
 blocks=1 if lane=='qualification' else plan[lane]['blocks'];expected=set()
 for b in range(blocks):
  for i in range(len(cases)):
   for leg in ['before','after']:
    for aff in (['all','one'] if lane=='memory' else ['all']):
     for protocol in (range(2) if lane=='memory' else [0]):expected.add((b,i,leg,aff,protocol))
 actual={(x['block'],x['index'],x['leg'],x['affinity'],x['protocol']) for x in lanes[lane]}
 assert actual==expected and len(actual)==len(lanes[lane]),lane
 for x in lanes[lane]:
  assert x['case']==cases[x['index']]
  cfg=plan['qualification'] if lane=='qualification' else plan[lane]
  if lane=='memory':cfg=cfg['protocols'][x['protocol']]
  assert x['samples']==cfg['samples'] and x['warmup']==cfg['warmup']
 # Check each predeclared case/leg schedule against actual start times.
 if lane not in ['qualification']:
  for b in range(blocks):
   observed=sorted([x for x in lanes[lane] if x['block']==b],key=lambda x:x['started_unix'])
   indices=list(range(len(cases)))
   if b%2:indices.reverse()
   schedule=[]
   for i in indices:
    for aff in (['all','one'] if lane=='memory' else ['all']):
     for protocol in (range(2) if lane=='memory' else [0]):
      for leg in plan['native']['orders'][b]:schedule.append((i,leg,aff,protocol))
   assert [(x['index'],x['leg'],x['affinity'],x['protocol']) for x in observed]==schedule

# Every sample's deterministic accounting and source work agrees across legs;
# observed parallel overlap is retained separately and only bounded by width.
for lane in ['native','observer','allocation']:
 cases=plan['allocation_cases'] if lane=='allocation' else plan['cases']
 for b in range(plan[lane]['blocks']):
  for i in range(len(cases)):
   pair={x['leg']:x for x in lanes[lane] if x['block']==b and x['index']==i}
   assert pair['before']['resources']==pair['after']['resources']
   clean=lambda z:[{k:v for k,v in sample.items() if k!='max_simultaneous_reads'} for sample in z]
   assert clean(pair['before']['source_metrics'])==clean(pair['after']['source_metrics'])

flags=[];summaries=[]
def summarize(lane,index,metric,before,after,extra=None,gate=True):
 assert len(before)==len(after)==6
 # Zero/negative baseline values use the frozen conservative absolute guard.
 if any(x<=0 for x in before):
  increase=any(y>x for x,y in zip(before,after));value=dict(estimate=None,ci95_low=None,ci95_high=None,before=before,after=after,absolute_increase=increase)
  flag=increase
 else:
  value=r.bootstrap_ratio(before,after,metric)
  flag=value['estimate']>1.05 and value['ci95_low']>1
 row=dict(lane=lane,index=index,metric=metric,before_median=st.median(before),after_median=st.median(after),**value,**(extra or {}),regression_flag=flag if gate else False,guarded=gate)
 summaries.append(row)
 if gate and flag:flags.append(row)
 return row
for i,case in enumerate(plan['cases']):
 subset=[x for x in lanes['native'] if x['index']==i]
 for metric in ['p50_ns','p95_ns','p99_ns','rss_kib','minor_faults','major_faults']:
  values=lambda leg:[next(x for x in subset if x['leg']==leg and x['block']==b)[metric] for b in range(6)]
  summarize('native',i,metric,values('before'),values('after'),dict(case=case),gate=metric in ['p50_ns','rss_kib'])
for i,case in enumerate(plan['memory_cases']):
 for aff in ['all','one']:
  for protocol in range(2):
   subset=[x for x in lanes['memory'] if x['index']==i and x['affinity']==aff and x['protocol']==protocol]
   for metric in ['after_preload','after_operation','after_batch_drop','after_package_drop','maximum observed RSS']:
    values=lambda leg:[next(x for x in subset if x['leg']==leg and x['block']==b)['memory_metrics'][metric] for b in range(6)]
    summarize('memory',i,metric,values('before'),values('after'),dict(case=case,affinity=aff,protocol=protocol))
for i,case in enumerate(plan['allocation_cases']):
 subset=[x for x in lanes['allocation'] if x['index']==i]
 for metric in plan['allocation']['metrics']:
  values=lambda leg:[st.median(next(x for x in subset if x['leg']==leg and x['block']==b)['allocation_metrics'][metric]) for b in range(6)]
  summarize('allocation',i,metric,values('before'),values('after'),dict(case=case))
benefits=[x for x in summaries if x['lane']=='native' and x['metric']=='p50_ns' and x['case']['state']=='primed' and x['case']['workers']>1 and x['estimate'] is not None and x['estimate']<=0.97 and x['ci95_high']<1]
decision='adopt' if benefits and not flags else 'reject'
result=dict(status='pass',decision=decision,regression_flags=flags,benefit_cases=len(benefits),counts={k:len(v) for k,v in lanes.items()},samples=sum(x['samples'] for x in rows),summaries=summaries,memory_counter_disagreements=[dict(path=x['path'],phase=s['phase'],sample=s['sample'],comparisons=s['comparisons']) for x in lanes['memory'] for s in x['residency'] if not all(s['comparisons'].values())],reader=d.desc(P/'readers.py'),qualifier=d.desc(P/'qualify.py'))
path=P/'analysis.json'
if path.exists():assert d.read(path)==result
else:d.write(path,result)
print('analysis PASS; decision',decision,'flags',len(flags),'benefit cases',len(benefits),'reports',len(rows))
for x in flags:print(x['lane'],x['index'],x['metric'],x.get('affinity'),x.get('protocol'),x['estimate'],x['ci95_low'],x['ci95_high'])
