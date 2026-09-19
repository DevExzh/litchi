#!/usr/bin/env python3
"""Retain every repeated startup-subtracted counter, without tail pooling."""
import json
from pathlib import Path
P=Path(__file__).resolve().parent
rows=[]
events=['cycles','instructions','branches','branch-misses','cache-misses','page-faults','task-clock']
for case in ['all','presentation','slides']:
 for repeat in range(3):
  parsed={}
  for count in [10,210]:
   values={}
   for line in (P/'profile'/f'{case}-{repeat}-{count}.stderr').read_text().splitlines():
    fields=line.split('\t')
    if len(fields)>4 and fields[2] in events:
     assert float(fields[4])>=99.0,('multiplexed counter',fields)
     values[fields[2]]=float(fields[0])
   assert set(values)==set(events),values
   parsed[count]=values
  slope={k:(parsed[210][k]-parsed[10][k])/2000 for k in events}
  assert slope['cycles']>0 and slope['instructions']>0
  rows.append(dict(case=case,repeat=repeat,counts=parsed,per_sequence=slope,ipc=slope['instructions']/slope['cycles'],task_clock_unit='milliseconds'))
(P/'counter-summary.json').write_text(json.dumps(rows,indent=2)+'\n')
print('Nine per-sequence slopes retained; each subtracts 10 from 210 samples, batch 10.')
