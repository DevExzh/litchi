"""Independently replay all confirmation quantiles and paired intervals."""
import driver as d
import math,statistics,random

def verify():
 a=d.read(d.P/'tail-analysis.json');count=0
 for row in a['summary']:
  legs={}
  for leg in ['before','after']:
   values=[]
   for block in range(6):
    p=d.P/'tail-runs'/f"{block:02}-{row['shape']}-{row['parts']:04}-{row['mode']}-{leg}"
    samples=sorted(x['elapsed_ns'] for x in d.read(p/'report.json')['samples']);assert len(samples)==1000;metric=row['metric']
    if metric=='mean_ns':value=sum(samples)/1000
    elif metric in ['p50_ns','p95_ns','p99_ns']:value=samples[{'p50_ns':499,'p95_ns':949,'p99_ns':989}[metric]]
    else:value=int((p/'time.txt').read_text().split()[['rss_kib','minor_faults','major_faults'].index(metric)])
    values.append(value)
   assert values==row['paired'][leg] and statistics.median(values)==row[leg+'_median'];legs[leg]=values
  if all(legs['before']):
   ratios=[y/x for x,y in zip(legs['before'],legs['after'])];rng=random.Random(844084)
   boot=sorted(statistics.median([ratios[rng.randrange(6)] for _ in range(6)]) for _ in range(10000))
   assert ratios==row['paired']['ratios'] and [statistics.median(ratios),boot[249],boot[9749]]==[row['paired'][k] for k in ['estimate','ci95_low','ci95_high']]
  else:assert row['paired']['estimate'] is None
  count+=1
 assert count==42
 return dict(status='pass',metric_rows=count,reports=72,samples=72000,analysis=d.desc(d.P/'tail-analysis.json'),checker=d.desc(d.P/'tail_crosscheck.py'))
if __name__=='__main__':
 v=verify();p=d.P/'tail-crosscheck.json'
 if p.exists():assert d.read(p)==v
 else:d.write(p,v)
 print('Confirmation arithmetic PASS',v['metric_rows'])
