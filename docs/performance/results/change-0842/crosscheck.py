"""Recompute all native reductions directly from raw reports/time counters."""
import driver as d
import math,random

def median(xs):
 ys=sorted(xs);n=len(ys)
 return ys[n//2] if n%2 else (ys[n//2-1]+ys[n//2])/2
analysis=d.read(d.P/'analysis.json');quantiles=0;intervals=0
for row in analysis['summary']:
 values={}
 for leg in ['before','after']:
  xs=[]
  for block in range(6):
   folder=d.P/'runs'/f"native-{block:02}-{row['shape']}-{row['mode']}-{'leg-a' if leg=='before' else 'leg-b'}"
   raw=d.read(folder/'report.json');ns=sorted(x['elapsed_ns'] for x in raw['samples']);assert len(ns)==30
   counter=list(map(int,(folder/'time.txt').read_text().split()));metric=row['metric']
   if metric.startswith('p') and metric.endswith('_ns'):
    percentile=int(metric[1:-3]);value=ns[(30*percentile+99)//100-1]
   elif metric=='mean_ns':value=sum(ns)/len(ns)
   else:value=counter[['rss_kib','minor_faults','major_faults'].index(metric)]
   xs.append(value)
  values[leg]=xs;assert row['paired'][leg]==xs;assert row[leg+'_median']==median(xs)
 if row['metric'] in ['p50_ns','p95_ns','p99_ns']:quantiles+=1
 b,a=values['before'],values['after']
 if not all(b):
  assert row['paired']['estimate'] is None and row['paired']['ci95_low'] is None and row['paired']['ci95_high'] is None
  continue
 ratios=[a[i]/b[i] for i in range(6)];assert row['paired']['estimate']==median(ratios)
 rng=random.Random(842084);draws=[]
 for _ in range(10000):
  indices=[rng.randrange(0,6) for _ in range(6)]
  ys=sorted(ratios[i] for i in indices);draws.append((ys[2]+ys[3])/2)
 draws.sort();assert row['paired']['ci95_low']==draws[249] and row['paired']['ci95_high']==draws[9749];intervals+=1
value=dict(status='pass',quantile_rows=quantiles,metric_rows=len(analysis['summary']),paired_intervals=intervals,analysis=d.desc(d.P/'analysis.json'),checker=d.desc(d.P/'crosscheck.py'))
p=d.P/'crosscheck.json'
if p.exists():assert d.read(p)==value
else:d.write(p,value)
print('Independent arithmetic PASS',quantiles,'quantiles;',intervals,'intervals')
