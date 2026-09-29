"""Independent native arithmetic and instruction partition cross-check."""
import driver as d
import math,statistics
analysis=d.read(d.P/'analysis.json');checked=0
for case in d.read(d.P/'plan.json')['cases']:
 for quantile in [50,95,99]:
  metric=f'p{quantile}_ns';vectors={}
  for leg in ['before','after']:
   vals=[]
   for block in range(6):
    folder=d.P/'runs'/f"native-{block:02}-{case['shape']}-{case['mode']}-{'leg-a' if leg=='before' else 'leg-b'}"
    data=d.read(folder/'report.json');xs=sorted(s['elapsed_ns'] for s in data['samples'])
    assert len(xs)==30;vals.append(xs[math.ceil(len(xs)*quantile/100)-1])
   vectors[leg]=vals
  row=next(x for x in analysis['summary'] if x['metric']==metric and all(x[k]==v for k,v in case.items()))
  assert row['paired']['before']==vectors['before'] and row['paired']['after']==vectors['after']
  assert row['paired']['estimate']==statistics.median([a/b for b,a in zip(vectors['before'],vectors['after'])])
  checked+=1
for row in analysis['profile_pairs']:
 assert sum(x['delta'] for x in row['self_differences'])==row['after_ir']-row['before_ir']
result=dict(status='pass',native_quantile_rows=checked,profile_partitions=len(analysis['profile_pairs']),analysis=d.desc(d.P/'analysis.json'))
p=d.P/'arithmetic-crosscheck.json'
if p.exists():assert d.read(p)==result
else:d.write(p,result)
print('Independent arithmetic PASS',checked,len(analysis['profile_pairs']))
