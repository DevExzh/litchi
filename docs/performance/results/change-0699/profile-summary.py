#!/usr/bin/env python3
"""Extract counter slopes and marker-loop instruction sample counts."""
import json,re
from pathlib import Path
P=Path(__file__).resolve().parent
EVENTS=['cycles','instructions','branches','branch-misses','cache-misses','page-faults','task-clock']
def counters(path):
 values={}
 for line in path.read_text().splitlines():
  fields=line.split('\t')
  if len(fields)>2 and fields[2] in EVENTS:values[fields[2]]=float(fields[0])
 assert set(values)==set(EVENTS)
 return values
rows=[]
for repeat in range(3):
 for phase in ['baseline','candidate']:
  a=counters(P/f'profiles/{phase}-counter-{repeat}-1000.stderr');b=counters(P/f'profiles/{phase}-counter-{repeat}-11000.stderr')
  rows.append(dict(phase=phase,repeat=repeat,per_iteration={name:(b[name]-a[name])/10000 for name in EVENTS}))
comparisons=[]
for repeat in range(3):
 by={r['phase']:r['per_iteration'] for r in rows if r['repeat']==repeat}
 comparisons.append(dict(repeat=repeat,delta_pct={k:((by['candidate'][k]/by['baseline'][k]-1)*100 if by['baseline'][k]!=0 else None) for k in EVENTS}))
(P/'counter-summary.json').write_text(json.dumps(dict(unit='counts except task-clock milliseconds',rows=rows,comparisons=comparisons),indent=2)+'\n')
rows=[]
for phase in ['baseline','candidate']:
 record=next(r for r in json.loads((P/'annotations.json').read_text()) if r['phase']==phase and r['name']=='samples')
 text=(P/f'annotations/{phase}-samples.stdout').read_text()
 total=int(re.search(r'\((\d+) samples,',text).group(1))
 samples=[(int(n),int(address,16)) for n,address in re.findall(r'^\s*(\d+)\s*:\s*([0-9a-f]+):',text,re.M)]
 assert sum(n for n,_ in samples)==total
 start=record['address']+0xc0;end=record['address']+0xdf
 loop=sum(n for n,address in samples if start<=address<=end)
 rows.append(dict(phase=phase,symbol=record['symbol'],symbol_samples=total,loop_first_address=start,loop_last_address=end,loop_samples=loop,loop_fraction_of_symbol_samples=loop/total))
(P/'instruction-summary.json').write_text(json.dumps(rows,indent=2)+'\n')
print('Counter slopes and bounded marker-loop sample summaries retained.')
