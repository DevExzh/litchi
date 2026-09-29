"""Strict paired native reduction; artifact and custody gates precede statistics."""
import driver as d
import qualify as q
import math,random,statistics

def bootstrap(before,after):
 assert len(before)==len(after)==6
 result=dict(before=before,after=after,estimate=None,ci95_low=None,ci95_high=None)
 if any(x==0 for x in before):return result
 ratios=[a/b for b,a in zip(before,after)];rng=random.Random(842084)
 draws=sorted(statistics.median([ratios[rng.randrange(6)] for _ in range(6)]) for _ in range(10000))
 return dict(before=before,after=after,ratios=ratios,estimate=statistics.median(ratios),ci95_low=draws[249],ci95_high=draws[9749])

def analyze():
 plan=d.read(d.P/'plan.json');qual=q.qualification();assert qual==d.read(d.P/'qualification.json')
 native=q.validate_items('native');allrows=qual['rows']+native
 assert {x['folder'] for x in allrows}|{str(d.P/plan['initial_failed_qualification'])}=={str(x) for x in (d.P/'runs').iterdir() if x.is_dir()}
 ordered=sorted(allrows,key=lambda x:x['started'])
 assert all(a['finished']<=b['started'] for a,b in zip(ordered,ordered[1:]))
 for p in (d.P/'commands').glob('*/receipt.json'):
  c=d.read(p)
  if c['argv'][0]=='cargo':
   assert all(x['finished']<=c['started_unix'] or x['started']>=c['finished_unix'] for x in native),'native overlaps Cargo'
 summary=[];flags=[];tails=[];spreads=[]
 for case in plan['cases']:
  legs={leg:sorted([x for x in native if x['row']['leg']==leg and all(x['row'][k]==v for k,v in case.items())],key=lambda x:x['row']['block']) for leg in ['before','after']}
  assert all([x['row']['block'] for x in rows]==list(range(6)) for rows in legs.values())
  for metric in ['p50_ns','p95_ns','p99_ns','mean_ns','rss_kib','minor_faults','major_faults']:
   b=[x[metric] for x in legs['before']];a=[x[metric] for x in legs['after']];stats=bootstrap(b,a)
   entry=dict(**case,metric=metric,before_median=statistics.median(b),after_median=statistics.median(a),paired=stats);summary.append(entry)
   if metric in ['p50_ns','rss_kib'] and stats['estimate'] is not None and stats['estimate']>1.05 and stats['ci95_low']>1:flags.append(entry)
   if metric in ['p95_ns','p99_ns'] and stats['estimate'] is not None and stats['estimate']>1.05:tails.append(entry)
  for leg,rows in legs.items():
   values=[x['p50_ns'] for x in rows]
   if max(values)/min(values)>1.2:spreads.append(dict(**case,leg=leg,block_p50_ns=values,ratio=max(values)/min(values)))
 return dict(status='pass',reports=len(allrows),samples=sum(x['row']['samples'] for x in allrows),native=native,summary=summary,regression_flags=flags,tail_review= tails,spread_flags=spreads,analyzer=d.desc(d.P/'analyze.py'),reader=d.desc(d.P/'reader.py'),limits=plan['limits'])
if __name__=='__main__':
 value=analyze();assert value['reports']==168 and value['samples']==4344
 out=d.P/'analysis.json'
 if out.exists():assert d.read(out)==value
 else:d.write(out,value)
 print('analysis PASS',value['reports'],value['samples'],'guard flags',len(value['regression_flags']),'tail points',len(value['tail_review']),'spread flags',len(value['spread_flags']))
