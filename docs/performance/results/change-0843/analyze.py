"""Keep native, guest-instruction and allocation populations separate."""
import driver as d
import qualify as q
import reader as r
import heap_reader as h
import statistics,random

def interval(values):
 assert len(values)==6
 rng=random.Random(843083);draws=sorted(statistics.median([values[rng.randrange(6)] for _ in range(6)]) for _ in range(10000))
 return dict(values=values,median=statistics.median(values),ci95_low=draws[249],ci95_high=draws[9749])

def analyze():
 qual=q.qualification();assert qual==d.read(d.P/'qualification.json')
 populations={lane:q.validate_items(lane) for lane in ['native','profiles','heap']}
 allrows=qual['rows']+[x for rows in populations.values() for x in rows]
 assert {x['folder'] for x in allrows}=={str(p) for p in (d.P/'runs').iterdir() if p.is_dir()}
 ordered=sorted(allrows,key=lambda x:x['started']);assert all(a['finished']<=b['started'] for a,b in zip(ordered,ordered[1:]))
 for path in (d.P/'commands').glob('*/receipt.json'):
  c=d.read(path)
  if c['argv'][0]=='cargo':assert all(x['finished']<=c['started_unix'] or x['started']>=c['finished_unix'] for x in populations['native']+populations['profiles']+populations['heap'])
 summary=[];spreads=[]
 for case in d.read(d.P/'plan.json')['cases']:
  rows=[x for x in populations['native'] if all(x['row'][k]==v for k,v in case.items())]
  assert [x['row']['block'] for x in rows]==list(range(6))
  for metric in ['p50_ns','p95_ns','p99_ns','mean_ns','rss_kib','minor_faults','major_faults']:
   summary.append(dict(**case,metric=metric,**interval([x[metric] for x in rows])))
  values=[x['p50_ns'] for x in rows]
  if max(values)/min(values)>1.2:spreads.append(dict(**case,p50_ns=values,ratio=max(values)/min(values)))
 profiles=[]
 for item in populations['profiles']:
  folder=d.Path(item['folder']);assert {p.name for p in folder.glob('profile.callgrind*')}=={'profile.callgrind','profile.callgrind.1'}
  command=d.read(q.bind(d.read(folder/'receipt.json')['command']))['argv']
  profiles.append(dict(row=item['row'],profile=r.validate_profile(folder/'profile.callgrind.1',folder/'profile.callgrind',command=command)))
 heaps=[]
 for item in populations['heap']:
  folder=d.Path(item['folder']);traces=list(folder.glob('trace*.zst'));assert len(traces)==1
  value=h.parse_zst(traces[0]);assert value['owner_attribution']['allocation_events']>0
  heaps.append(dict(row=item['row'],trace=d.desc(traces[0]),attribution=value))
 return dict(status='pass',disposition='diagnostic-only-no-production-change',reports=len(allrows),samples=sum(x['row']['samples'] for x in allrows),populations=populations,native_summary=summary,spread_flags=spreads,profiles=profiles,heaps=heaps,analyzer=d.desc(d.P/'analyze.py'),reader=d.desc(d.P/'reader.py'),heap_reader=d.desc(d.P/'heap_reader.py'),limits=d.read(d.P/'plan.json')['limits'])
if __name__=='__main__':
 value=analyze();assert value['reports']==72 and value['samples']==1464
 p=d.P/'analysis.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Analysis PASS',value['reports'],'reports;',value['samples'],'samples;',len(value['spread_flags']),'native spread flags')
