"""Independent point arithmetic from raw native sample arrays."""
import driver as d
import statistics as st
import math
rows=d.read(d.P/'analysis.json')['summary'];checks=[]
for case in d.read(d.P/'plan.json')['cases']:
 for metric,quantile in [('p50_ns',.5),('p95_ns',.95),('p99_ns',.99)]:
  pair={}
  for leg,label in [('before','leg-a'),('after','leg-b')]:
   values=[]
   for block in range(6):
    report=d.read(d.P/'runs/native'/f'{block:02}-{case}-{label}'/'report.json')
    times=sorted(sample['wall_ns'] for sample in report['samples'])
    values.append(times[math.ceil(quantile*len(times))-1])
   pair[leg]=values
  row=next(row for row in rows if row['case']==case and row['metric']==metric and row['lane']=='native')
  assert row['before']==pair['before'] and row['after']==pair['after']
  assert row['estimate']==st.median(after/before for after,before in zip(pair['after'],pair['before']))
  checks.append(dict(case=case,metric=metric,estimate=row['estimate']))
result=dict(status='pass',raw_quantile_rows=checks,scope='Independent raw native nearest-rank and paired-point recomputation; bootstrap covered by frozen reader statistical tests.')
assert d.read(d.P/'arithmetic-crosscheck.json')==result
print('crosscheck PASS',len(checks),'raw quantile rows')
