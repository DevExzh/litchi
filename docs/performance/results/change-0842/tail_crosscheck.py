"""Independent raw tiny-tail quantiles and matched-block bootstrap arithmetic."""
import driver as d
import random

def median(xs):
 ys=sorted(xs);return (ys[2]+ys[3])/2
v=d.read(d.P/'tail-analysis.json')
for row in v['summary']:
 values={}
 for leg in ['before','after']:
  values[leg]=[]
  for block in range(6):
   folder=d.P/'tails'/f"tail-{block:02}-tiny-{row['mode']}-{'leg-a' if leg=='before' else 'leg-b'}"
   if row['metric']=='rss_kib':x=int((folder/'time.txt').read_text().split()[0])
   else:
    ns=sorted(x['elapsed_ns'] for x in d.read(folder/'report.json')['samples']);assert len(ns)==1000
    x=ns[int(row['metric'][1:-3])*10-1]
   values[leg].append(x)
  assert values[leg]==row['paired'][leg] and median(values[leg])==row[leg+'_median']
 ratios=[a/b for a,b in zip(values['after'],values['before'])];assert median(ratios)==row['paired']['estimate']
 rng=random.Random(842084);draws=sorted(median([ratios[rng.randrange(6)] for _ in range(6)]) for _ in range(10000))
 assert draws[249]==row['paired']['ci95_low'] and draws[9749]==row['paired']['ci95_high']
value=dict(status='pass',metric_rows=8,paired_intervals=8,analysis=d.desc(d.P/'tail-analysis.json'),checker=d.desc(d.P/'tail_crosscheck.py'))
p=d.P/'tail-crosscheck.json'
if p.exists():assert d.read(p)==value
else:d.write(p,value)
print('Independent tail arithmetic PASS;8rows/intervals')
