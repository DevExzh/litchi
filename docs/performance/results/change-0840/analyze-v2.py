"""Replay every raw report, independent artifact comparison and frozen decision."""
import driver as d
import qualify as q
import readers as r
import statistics as st
import math,json
plan=d.read(d.P/'plan.json');rows=q.all_rows();cases=plan['cases']
expected=set()
for lane in ['qualification','observer','allocation-preflight','native','allocation']:
 for b in range(plan.get(lane,{}).get('blocks',1)):
  for case in cases:
   for leg in ['before','after']:expected.add((lane,b,case,leg))
assert {(x['lane'],x['block'],x['case'],x['leg']) for x in rows}==expected and len(rows)==len(expected)
for lane in ['native','allocation']:
 for b in range(6):
  subset=sorted([x for x in rows if x['lane']==lane and x['block']==b],key=lambda x:x['started_unix'])
  order=cases if b%2==0 else list(reversed(cases))
  assert [(x['case'],x['leg']) for x in subset]==[(case,leg) for case in order for leg in plan[lane]['orders'][b]]
summary=[];flags=[];tails=[]
def summarize(lane,case,metric,before,after,guard):
 assert len(before)==len(after)==6
 if any(v<=0 for v in before):
  flag=any(y>x for x,y in zip(before,after));stats=dict(estimate=None,ci95_low=None,ci95_high=None,absolute_increase=flag)
 else:
  stats=r.bootstrap_ratio(before,after);flag=stats['estimate']>1.05 and stats['ci95_low']>1
 stats={key:value for key,value in stats.items() if key not in ['before','after']}
 row=dict(lane=lane,case=case,metric=metric,before=before,after=after,before_median=st.median(before),after_median=st.median(after),**stats,guarded=guard,regression_flag=bool(guard and flag))
 summary.append(row)
 if row['regression_flag']:flags.append(row)
 if metric in ['p95_ns','p99_ns'] and stats['estimate']>1.05:tails.append(row)
for case in cases:
 for lane in ['native','allocation']:
  subset=[x for x in rows if x['case']==case and x['lane']==lane]
  metrics=['p50_ns','p95_ns','p99_ns','rss_kib','minor_faults','major_faults'] if lane=='native' else ['allocation_calls','allocated_bytes','region_peak_live_bytes_minus_entry','retained_live_bytes_delta']
  for metric in metrics:
   def values(leg):
    result=[]
    for b in range(6):
     x=next(x for x in subset if x['leg']==leg and x['block']==b)
     result.append(x[metric] if lane=='native' else st.median(x['allocation_metrics'][metric]))
    return result
   summarize(lane,case,metric,values('before'),values('after'),lane=='allocation' or metric in ['p50_ns','rss_kib'])
benefit=next(x for x in summary if x['lane']=='native' and x['case']=='doc-payload' and x['metric']=='p50_ns')
qualified_benefit=benefit['estimate']<=0.97 and benefit['ci95_high']<1
result=dict(status='pass',decision='adopt' if qualified_benefit and not flags else 'reject',reports=len(rows),samples=sum(plan['qualification' if x['lane']=='allocation-preflight' else x['lane']]['samples'] for x in rows),counts={lane:sum(x['lane']==lane for x in rows) for lane in ['qualification','observer','allocation-preflight','native','allocation']},summary=summary,regression_flags=flags,tail_observations=tails,target_benefit=qualified_benefit,readers=d.desc(d.P/'readers.py'),qualifier=d.desc(d.P/'qualify.py'),analyzer=d.desc(d.P/'analyze-v2.py'))
path=d.P/'analysis.json'
if path.exists():assert d.read(path)==result
else:d.write(path,result)
print('analysis PASS',result['decision'],'flags',len(flags),'reports',len(rows),'samples',result['samples'])
