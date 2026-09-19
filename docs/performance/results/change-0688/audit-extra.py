#!/usr/bin/env python3
"""Audit longer controls and native-child counter/RSS captures."""
import hashlib,json,math,re,statistics,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def values(p):return {k:int(v) for k,v in (x.split('=') for x in p.read_text().strip().split('\t'))}
def pct(a,b):return (a/b-1)*100
base=read(P/'baseline.json')
for phase in ['baseline','candidate']:
 builds=read(P/(phase+'-builds.json'));binary=next(r['binary_sha256'] for r in builds if r['binary']=='xls0684-repeat')
 for folder in ['repeat','diagnostics']:
  d=P/folder/phase;m=read(d/'manifest.json')
  assert m['source_sha256']==builds[0]['source_sha256']
  for n,h in m['source_sha256'].items():
   actual=sha(ROOT/n) if phase=='candidate' else hashlib.sha256(subprocess.check_output(['git','show',base['baseline_head']+':'+n],cwd=ROOT)).hexdigest()
   assert actual==h,n
  for n,h in m['raw_sha256'].items():assert sha(d/n)==h,n
  for field in ['probe_sha256','corpus']:
   for n,h in m.get(field,{}).items():assert sha(ROOT/n)==h,n
  assert sha(P/'cases.json')==m['cases_sha256']
  if folder=='repeat':assert m['binary_sha256'][phase]==binary
  else:assert m['binary_sha256']==binary
  assert all(c['exit_code']==0 for c in m['commands'])
rows=[]
for case in ['54016-stored-2097152','54016-late','54016-missing-1048576','Plan1-stored-2097152','Simple-stored-2097152','45365-first']:
 for mode in ['owned','file']:
  data={}
  for leg in ['aa1','aa2','a1','b1','b2','a2']:
   phase='baseline' if leg.startswith('aa') else 'candidate';raw=[values(P/'repeat'/phase/f'{case}-{mode}-{leg}-{sample}.tsv') for sample in range(9)]
   assert all(v['repeats']==50000 and v['found']==(0 if 'missing' in case else 50000) for v in raw)
   ns=[v['nanos']/50000 for v in raw];data[leg]=dict(samples_ns_per_query=ns,median=statistics.median(ns),mean=statistics.mean(ns),minimum=min(ns),maximum=max(ns))
  changes={k:pct(data[a]['median'],data[b]['median']) for k,a,b in [('b1_a1','b1','a1'),('b2_a2','b2','a2'),('aa','aa2','aa1'),('abba_a','a2','a1'),('abba_b','b2','b1')]}
  rows.append(dict(case=case,mode=mode,max_query_index_bytes=2097152,repeats_per_sample=50000,samples_per_leg=9,timing=data,percent=changes))
(P/'repeat-comparison.json').write_text(json.dumps(rows,indent=2)+'\n')
def counters(p):
 d={}
 for line in p.read_text().splitlines():
  parts=line.split(',')
  if len(parts)>2 and parts[0].isdigit():d[parts[2]]=int(parts[0]);assert float(parts[4])>=99.9
 assert set(d)=={'cycles','instructions','branches','branch-misses','cache-misses','page-faults'}
 return d
rows=[]
for case in ['54016-stored-2097152','54016-late','Simple-stored-2097152','Plan1-stored-2097152']:
 for mode in ['owned','file']:
  phases={}
  for phase in ['baseline','candidate']:
   d=P/'diagnostics'/phase;samples=[]
   for repeat in range(3):
    legs={};rss={}
    for n in [10,100010]:
     stem=f'{case}-{mode}-{repeat}-{n}';v=values(d/(stem+'.tsv'));assert v['repeats']==v['found']==n
     legs[n]=counters(d/(stem+'.csv'));rss[n]=int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',(d/(stem+'.rss.txt')).read_text())[1])
    samples.append(dict(per_extra_query={k:(legs[100010][k]-legs[10][k])/100000 for k in legs[10]},rss_kib=rss))
   phases[phase]=dict(samples=samples,median_per_extra_query={k:statistics.median(s['per_extra_query'][k] for s in samples) for k in samples[0]['per_extra_query']},median_rss_kib={str(n):statistics.median(s['rss_kib'][n] for s in samples) for n in [10,100010]})
  b=phases['baseline'];a=phases['candidate'];rows.append(dict(case=case,mode=mode,**phases,percent={k:pct(a['median_per_extra_query'][k],v) if v>0 and a['median_per_extra_query'][k]>=0 else None for k,v in b['median_per_extra_query'].items()},rss_percent={k:pct(a['median_rss_kib'][k],v) for k,v in b['median_rss_kib'].items()}))
(P/'diagnostics-comparison.json').write_text(json.dumps(rows,indent=2)+'\n')
print('PASS longer controls: 12 groups x 6 legs x 9 samples x 50,000 queries; default 2 MiB in every long-loop case.')
print('PASS counters/RSS: 8 groups x 2 lengths x 3 repeats x 2 binaries; subtract 100,000 extra queries. Counters include whole-process work.')
