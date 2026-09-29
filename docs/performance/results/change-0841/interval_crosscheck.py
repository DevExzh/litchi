"""Recompute paired percentile endpoints without importing the main reducer."""
import driver as d
import random
v=d.read(d.P/'analysis.json');count=0
for row in v['summary']:
 p=row['paired'];ratios=[after/before for before,after in zip(p['before'],p['after']) if before>0]
 if not ratios:
  assert p['ci95_low'] is None and p['ci95_high'] is None;continue
 rng=random.Random(841084);medians=[]
 for _ in range(10000):
  sample=sorted(ratios[rng.randrange(len(ratios))] for _ in range(len(ratios)))
  middle=len(sample)//2
  medians.append(sample[middle] if len(sample)%2 else (sample[middle-1]+sample[middle])/2)
 medians.sort();assert (medians[249],medians[9749])==(p['ci95_low'],p['ci95_high'])
 count+=1
result=dict(status='pass',paired_intervals=count,analysis=d.desc(d.P/'analysis.json'),checker=d.desc(d.P/'interval_crosscheck.py'))
p=d.P/'interval-crosscheck.json'
if p.exists():assert d.read(p)==result
else:d.write(p,result)
print('Independent paired intervals PASS',count)
