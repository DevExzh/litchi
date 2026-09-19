#!/usr/bin/env python3
"""Summarize each process leg without pooling across process-order effects."""
import json,math,statistics
from pathlib import Path
P=Path(__file__).resolve().parent
FIELDS=['p50_ns','mean_ns','p95_ns','p99_ns']
def parse(path):
 header={};cases={};current=None
 for line in path.read_text().splitlines():
  f=line.split('\t');key=f[0]
  if key=='case':
   current=f[1];assert current not in cases
   cases[current]={'metadata':{},'samples':[]}
  elif key.isdigit():
   assert current is not None and int(key)==len(cases[current]['samples']) and len(f)==2
   cases[current]['samples'].append(int(f[1]))
  elif key=='sample_ns':continue
  elif key=='all_iterations_passed' or current is None:header[key]=f[1:]
  else:cases[current]['metadata'][key]=f[1:]
 return header,cases
rows=[]
for run in json.loads((P/'runs.json').read_text()):
 header,cases=parse(P/run['stdout'])
 assert header['samples']==['300'] and header['warmups']==['10'] and header['all_iterations_passed']==['true']
 assert len(cases)==(1 if run['mode']=='case' else 10)
 if run['mode']=='case':assert list(cases)==[run['case']]
 for case,item in cases.items():
  values=item['samples'];assert len(values)==300
  ordered=sorted(values)
  stats=dict(samples=len(values),p50_ns=statistics.median(values),mean_ns=statistics.mean(values),p95_ns=ordered[math.ceil(.95*len(values))-1],p99_ns=ordered[math.ceil(.99*len(values))-1],min_ns=min(values),max_ns=max(values))
  rows.append(dict(mode=run['mode'],case=case,leg=run['leg'],phase=run['phase'],metadata=item['metadata'],stats=stats))
rows.sort(key=lambda r:(r['mode'],r['case'],r['leg']))
identities={}
for row in rows:
 case=row['case'];meta=row['metadata']
 assert case not in identities or identities[case]==meta,case
 identities[case]=meta
 if 'observed_error_debug' in meta:assert meta['expected_error_debug']==meta['observed_error_debug']
comparisons=[]
for mode in ['case','matrix']:
 pairs=[(1,0),(2,3),(4,5),(7,6),(9,8),(10,11)] if mode=='case' else [(1,0),(2,3)]
 for case in sorted({r['case'] for r in rows if r['mode']==mode}):
  by={r['leg']:r['stats'] for r in rows if r['mode']==mode and r['case']==case}
  for b,a in pairs:
   comparisons.append(dict(mode=mode,case=case,candidate_leg=b,baseline_leg=a,delta_pct={f:(by[b][f]/by[a][f]-1)*100 for f in FIELDS},delta_ns={f:by[b][f]-by[a][f] for f in FIELDS}))
triggers=[dict(mode=r['mode'],case=r['case'],candidate_leg=r['candidate_leg'],baseline_leg=r['baseline_leg'],metric=f,delta_pct=r['delta_pct'][f],delta_ns=r['delta_ns'][f]) for r in comparisons for f in FIELDS if r['delta_pct'][f]>5]
for name,data in [('summary.json',rows),('comparisons.json',comparisons),('triggers.json',triggers)]:
 (P/name).write_text(json.dumps(data,indent=2)+'\n')
print(len(rows),'leg/case summaries;',len(comparisons),'pairs;',len(triggers),'review triggers')
