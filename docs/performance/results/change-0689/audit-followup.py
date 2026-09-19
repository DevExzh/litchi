#!/usr/bin/env python3
"""Retain all follow-up phases, tails and control drift without replacing primary data."""
import gzip,hashlib,json,math,statistics
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];out=P/'followup'
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
m=read(out/'manifest.json');assert len(m['commands'])==24 and all(c['exit_code']==0 for c in m['commands'])
for n,h in m['raw_sha256'].items():assert sha(out/n)==h,n
for field in ['current_source_sha256','corpus','probe_sha256']:
 for n,h in m[field].items():assert sha(ROOT/n)==h,n
assert m['cases_sha256']==sha(P/'cases.json')
for phase,h in m['binary_sha256'].items():assert h==next(b['binary_sha256'] for b in read(P/(phase+'-builds.json')) if b['binary']=='xls-index-retry-probe-0686')
def stats(v):
 v=sorted(v);return dict(p50=statistics.median(v),mean=statistics.mean(v),p95=v[math.ceil(len(v)*.95)-1],p99=v[math.ceil(len(v)*.99)-1],minimum=v[0],maximum=v[-1])
rows=[];total=0
for case,mode in [('54016-missing-1048576','owned'),('45365-first','file'),('45365-late','file'),('synthetic-70000-default','owned')]:
 data={};expected=None
 for leg in ['aa1','aa2','a1','b1','b2','a2']:
  j=json.loads(gzip.decompress((out/f'{case}-{mode}-{leg}.json.gz').read_bytes()));assert j['samples']==1000 and j['warmups']==10 and len(j['records'])==1000
  for r in j['records']:
   signature=[r['open']['outcome'],*[q['outcome'] for q in r['queries']]]
   if expected is None:expected=signature
   assert signature==expected and r['all_queries_agree'];total+=1
  records=j['records'];data[leg]={'open':[r['open']['elapsed_ns'] for r in records]}
  for i in range(8):data[leg]['q'+str(i+1)]=[r['queries'][i]['elapsed_ns'] for r in records]
  data[leg]['q3-to-q8-mean']=[statistics.mean(q['elapsed_ns'] for q in r['queries'][2:]) for r in records]
  data[leg]['open-plus-eight']=[r['open']['elapsed_ns']+sum(q['elapsed_ns'] for q in r['queries']) for r in records]
 for metric in data['a1']:
  t={leg:stats(v[metric]) for leg,v in data.items()};pct={stat:{name:(t[a][stat]/t[b][stat]-1)*100 for name,a,b in [('b1_a1','b1','a1'),('b2_a2','b2','a2'),('aa','aa2','aa1'),('abba_a','a2','a1'),('abba_b','b2','b1')]} for stat in ['p50','mean','p95','p99']};rows.append(dict(case=case,mode=mode,metric=metric,timing=t,percent=pct))
assert total==24000
(P/'followup-comparison.json').write_text(json.dumps(rows,indent=2)+'\n')
lines=['# Supplementary drift and tail controls','','1,000 fresh owners per leg; A/A then A/B/B/A. All 24,000 owners and 192,000','queries retain outcome parity. This supplements the original captures. Tails','remain descriptive; the original 100-sample results are not overwritten.','','| Case | Source | Phase | Statistic | A1 | B1 | B2 | A2 | B1/A1 % | B2/A2 % | A/A % | ABBA A % |','|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|']
for r in rows:
 for stat in ['p50','p95','p99']:
  if r['metric'] in ['open','q1','q2','q8','q3-to-q8-mean','open-plus-eight']:
   t=r['timing'];v=r['percent'][stat];lines.append('| '+r['case']+' | '+r['mode']+' | '+r['metric']+' | '+stat+' | '+' | '.join(f'{t[k][stat]:g}' for k in ['a1','b1','b2','a2'])+' | '+' | '.join(f'{v[k]:+.2f}' for k in ['b1_a1','b2_a2','aa','abba_a'])+' |')
(P/'followup-summary.md').write_text('\n'.join(lines)+'\n');print('PASS supplementary 24,000 owners / 192,000 queries, all raw/source/probe/binary bindings and outcomes.')
