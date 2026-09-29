"""Independent native arithmetic replay from raw reports and process counters."""
import driver as d
import math,random,statistics

def verify():
 a=d.read(d.P/'analysis.json'); checked=0
 for row in a['native_summary']:
  values=[]
  for block in range(6):
   folder=d.P/'runs'/f"native-{block:02}-{row['parts']:04}-{row['mode']}"
   samples=sorted(x['elapsed_ns'] for x in d.read(folder/'report.json')['samples'])
   metric=row['metric']
   if metric=='mean_ns':value=sum(samples)/len(samples)
   elif metric in ['p50_ns','p95_ns','p99_ns']:
    percentile={'p50_ns':.5,'p95_ns':.95,'p99_ns':.99}[metric]
    value=samples[max(0,math.ceil(len(samples)*percentile)-1)]
   else:value=int((folder/'time.txt').read_text().split()[['rss_kib','minor_faults','major_faults'].index(metric)])
   values.append(value)
  assert values==row['values'],(row,values)
  rng=random.Random(843083)
  boot=sorted(statistics.median([values[rng.randrange(len(values))] for _ in values]) for _ in range(10000))
  assert [statistics.median(values),boot[249],boot[9749]]==[row['median'],row['ci95_low'],row['ci95_high']]
  checked+=1
 assert checked==56
 return dict(status='pass',metric_rows=checked,raw_native_reports=48,analysis=d.desc(d.P/'analysis.json'),checker=d.desc(d.P/'crosscheck.py'))
if __name__=='__main__':
 value=verify();p=d.P/'crosscheck.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Independent arithmetic PASS',value['metric_rows'],'rows')
