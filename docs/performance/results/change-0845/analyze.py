"""Paired native comparison and separate full-stack allocation comparison."""
import driver as d
import qualify as q
import heap_reader as h
import random,statistics

def paired(before,after):
 assert len(before)==len(after)==6
 if any(x==0 for x in before):return dict(before=before,after=after,ratios=None,estimate=None,ci95_low=None,ci95_high=None)
 ratios=[a/b for b,a in zip(before,after)];rng=random.Random(845085)
 draws=sorted(statistics.median([ratios[rng.randrange(6)] for _ in range(6)]) for _ in range(10000))
 return dict(before=before,after=after,ratios=ratios,estimate=statistics.median(ratios),ci95_low=draws[249],ci95_high=draws[9749])

def analyze():
 plan=d.read(d.P/'plan.json');qual=q.qualification();assert qual==d.read(d.P/'qualification.json')
 native=q.validate_items('native');heap=q.validate_items('heap');allrows=qual['rows']+native+heap
 for lane in ['native','heap']:
  capture=d.read(d.P/f'capture-{lane}.json')
  assert capture['status']=='pass' and capture['plan']==d.desc(d.P/'plan.json')
  assert capture['receipts']==[d.desc(q.folder_for(lane,row)/'receipt.json') for row in plan[lane]]
 assert {x['folder'] for x in allrows}=={str(p) for p in (d.P/'runs').iterdir() if p.is_dir()}
 ordered=sorted(allrows,key=lambda x:x['started']);assert all(a['finished']<=b['started'] for a,b in zip(ordered,ordered[1:]))
 for lane,rows in [('native',native),('heap',heap)]:assert [x['row'] for x in sorted(rows,key=lambda x:x['started'])]==plan[lane]
 for p in (d.P/'commands').glob('*/receipt.json'):
  c=d.read(p)
  if c['argv'][0]=='cargo':assert all(x['finished']<=c['started_unix'] or x['started']>=c['finished_unix'] for x in native+heap)
 summary=[];flags=[];tails=[];spreads=[]
 for case in plan['cases']:
  legs={leg:sorted([x for x in native if x['row']['leg']==leg and all(x['row'][k]==v for k,v in case.items())],key=lambda x:x['row']['block']) for leg in ['before','after']}
  assert all([x['row']['block'] for x in rows]==list(range(6)) for rows in legs.values())
  for metric in ['p50_ns','p95_ns','p99_ns','mean_ns','rss_kib','minor_faults','major_faults']:
   b=[x[metric] for x in legs['before']];a=[x[metric] for x in legs['after']];stats=paired(b,a)
   row=dict(**case,metric=metric,before_median=statistics.median(b),after_median=statistics.median(a),paired=stats);summary.append(row)
   if metric in ['p50_ns','rss_kib'] and stats['estimate'] is not None and stats['estimate']>1.05 and stats['ci95_low']>1:flags.append(row)
   if metric in ['p95_ns','p99_ns'] and stats['estimate'] is not None and stats['estimate']>1.05:tails.append(row)
  for leg,rows in legs.items():
   values=[x['p50_ns'] for x in rows]
   if max(values)/min(values)>1.2:spreads.append(dict(**case,leg=leg,block_p50_ns=values,ratio=max(values)/min(values)))
 traces=[]
 for item in heap:
  folder=d.Path(item['folder']);paths=list(folder.glob('trace*.zst'));assert len(paths)==1
  value=h.parse_zst(paths[0]);assert value['owner_attribution']['allocation_events']>0
  traces.append(dict(row=item['row'],trace=d.desc(paths[0]),attribution=dict(owner_attribution=value['owner_attribution'])))
 allocations=[];allocation_flags=[]
 for shape in ['shared','root','deep-shared','deep-unique']:
  for parts in [4,64,1024,4096]:
   for metric in ['allocation_events','requested_bytes']:
    legs={leg:[x['attribution']['owner_attribution'][metric] for x in traces if x['row']['leg']==leg and x['row']['shape']==shape and x['row']['parts']==parts] for leg in ['before','after']}
    assert all(len(values)==2 for values in legs.values())
    b=statistics.median(legs['before']);a=statistics.median(legs['after']);row=dict(shape=shape,parts=parts,metric=metric,before=legs['before'],after=legs['after'],ratio=a/b)
    allocations.append(row)
    if a/b>1.05:allocation_flags.append(row)
 return dict(status='pass',reports=len(allrows),samples=sum(x['row']['samples'] for x in allrows),native=native,heap=heap,summary=summary,regression_flags=flags,tail_review=tails,spread_flags=spreads,heap_traces=traces,allocations=allocations,allocation_flags=allocation_flags,analyzer=d.desc(d.P/'analyze.py'),reader=d.desc(d.P/'reader.py'),heap_reader=d.desc(d.P/'heap_reader.py'),limits=plan['limits'])
if __name__=='__main__':
 value=analyze();assert value['reports']==512 and value['samples']==211328
 p=d.P/'analysis.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Analysis PASS',value['reports'],value['samples'],'nativeflags',len(value['regression_flags']),'tailpoints',len(value['tail_review']),'spreadflags',len(value['spread_flags']),'allocationflags',len(value['allocation_flags']))
