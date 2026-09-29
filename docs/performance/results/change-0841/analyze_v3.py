"""Replay all planned native and guest-instruction evidence; never adopt source."""
import driver as d
import qualify_v3 as q
import readers_v2 as r
import statistics,sys

def analyze():
 plan=d.read(d.P/'plan.json');qual=q.qualification();assert qual==d.read(d.P/'qualification-v2.json')
 native=q.validate_items('native');profiles=q.validate_items('profiles')
 actual={x['folder'] for x in qual['rows']+native+profiles}
 assert actual=={str(x) for x in (d.P/'runs').iterdir() if x.is_dir()}
 allrows=qual['rows']+native+profiles
 ordered=sorted(allrows,key=lambda x:x['started'])
 assert all(a['finished']<=b['started'] for a,b in zip(ordered,ordered[1:])), 'workload intervals overlap'
 cargo=[]
 for p in (d.P/'commands').glob('*/receipt.json'):
  c=d.read(p)
  if c['argv'][0]=='cargo':cargo.append((c['started_unix'],c['finished_unix']))
 assert all(x['finished']<=a or x['started']>=b for x in native+profiles for a,b in cargo),'formal capture overlaps Cargo'
 summary=[];flags=[];spreads=[]
 for case in plan['cases']:
  legs={leg:sorted([x for x in native if x['row']['leg']==leg and all(x['row'][k]==v for k,v in case.items())],key=lambda x:x['row']['block']) for leg in ['before','after']}
  assert all([x['row']['block'] for x in rows]==list(range(6)) for rows in legs.values())
  for metric in ['p50_ns','p95_ns','p99_ns','mean_ns','rss_kib','minor_faults','major_faults']:
   before=[x[metric] for x in legs['before']];after=[x[metric] for x in legs['after']]
   stats=r.bootstrap_ratio(before,after,metric)
   entry=dict(**case,metric=metric,before_median=statistics.median(before),after_median=statistics.median(after),paired=stats)
   summary.append(entry)
   if metric in ['p50_ns','rss_kib'] and stats['estimate'] is not None and stats['estimate']>1.05 and stats['ci95_low']>1:flags.append(entry)
  for leg,rows in legs.items():
   values=[x['p50_ns'] for x in rows]
   if max(values)/min(values)>1.2:spreads.append(dict(**case,leg=leg,block_p50_ns=values,ratio=max(values)/min(values)))
 profile_pairs=[]
 for shape in ['many','unicode','payload','rich']:
  for block in range(2):
   pair={leg:next(x for x in profiles if x['row']['shape']==shape and x['row']['block']==block and x['row']['leg']==leg) for leg in ['before','after']}
   before=pair['before']['profile'];after=pair['after']['profile']
   def selfmap(value):
    result={}
    for row in value['function_rows']:result[row['name']]=result.get(row['name'],0)+row['self_ir']
    assert sum(result.values())==value['summary_ir'];return result
   b=selfmap(before);a=selfmap(after)
   changes=[dict(name=name,before_recorded_self_ir=b.get(name,0),after_recorded_self_ir=a.get(name,0),delta=a.get(name,0)-b.get(name,0),present_before=name in b,present_after=name in a) for name in sorted(set(a)|set(b))]
   changes.sort(key=lambda x:(-abs(x['delta']),x['name']))
   assert sum(x['delta'] for x in changes)==after['summary_ir']-before['summary_ir']
   profile_pairs.append(dict(shape=shape,block=block,before_ir=before['summary_ir'],after_ir=after['summary_ir'],ratio=after['summary_ir']/before['summary_ir'],self_differences=changes))
 return dict(status='pass',decision='diagnostic-only-restore-production',reports=len(allrows),samples=sum(x['row']['samples'] for x in allrows),native=native,profiles=profiles,summary=summary,regression_flags=flags,spread_flags=spreads,profile_pairs=profile_pairs,analyzer=d.desc(d.P/'analyze_v3.py'),reader=d.desc(d.P/'readers_v2.py'),limits=plan['limits'])
if __name__=='__main__':
 value=analyze();assert value['reports']==156 and value['samples']==3636
 out=d.P/'analysis.json'
 if out.exists():assert d.read(out)==value
 else:d.write(out,value)
 print('analysis PASS',value['reports'],value['samples'],'native guard flags',len(value['regression_flags']))
