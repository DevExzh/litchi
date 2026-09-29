"""Independent arithmetic from raw native reports and process counters."""
import driver as d
import math,statistics,random

def verify():
 a=d.read(d.P/'analysis.json');count=0
 for row in a['summary']:
  legs={}
  for leg in ['before','after']:
   values=[]
   for block in range(6):
    folder=d.P/'runs'/f"native-{block:02}-{row['shape']}-{row['parts']:04}-{row['mode']}-{leg}"
    samples=sorted(x['elapsed_ns'] for x in d.read(folder/'report.json')['samples']);metric=row['metric']
    if metric=='mean_ns':value=sum(samples)/len(samples)
    elif metric in ['p50_ns','p95_ns','p99_ns']:value=samples[math.ceil(len(samples)*{'p50_ns':.5,'p95_ns':.95,'p99_ns':.99}[metric])-1]
    else:value=int((folder/'time.txt').read_text().split()[['rss_kib','minor_faults','major_faults'].index(metric)])
    values.append(value)
   assert values==row['paired'][leg] and statistics.median(values)==row[leg+'_median'];legs[leg]=values
  if all(legs['before']):
   ratios=[y/x for x,y in zip(legs['before'],legs['after'])];rng=random.Random(844084)
   boot=sorted(statistics.median([ratios[rng.randrange(len(ratios))] for _ in ratios]) for _ in range(10000))
   assert ratios==row['paired']['ratios'] and [statistics.median(ratios),boot[249],boot[9749]]==[row['paired'][k] for k in ['estimate','ci95_low','ci95_high']]
  else:assert row['paired']['estimate'] is None
  count+=1
 assert count==224
 return dict(status='pass',metric_rows=count,raw_native_reports=384,analysis=d.desc(d.P/'analysis.json'),checker=d.desc(d.P/'crosscheck.py'))
if __name__=='__main__':
 v=verify();p=d.P/'crosscheck.json'
 if p.exists():assert d.read(p)==v
 else:d.write(p,v)
 print('Independent arithmetic PASS',v['metric_rows'])
