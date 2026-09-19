#!/usr/bin/env python3
import hashlib,json,re,statistics,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def counters(p):
 result={}
 for line in p.read_text().splitlines():
  bits=line.split(',')
  if len(bits)>2 and bits[0].isdigit():result[bits[2]]=int(bits[0])
 return result
base=read(P/'baseline.json');rows=[]
for phase in ['baseline','candidate']:
 d=P/'diagnostics'/phase;m=read(d/'manifest.json')
 for n,h in m['raw_sha256'].items():assert sha(d/n)==h,n
 for field in ['probe_sha256','corpus']:
  for n,h in m[field].items():assert sha(ROOT/n)==h,n
 for n,h in m['source_sha256'].items():
  actual=sha(ROOT/n) if phase=='candidate' else hashlib.sha256(subprocess.check_output(['git','show',base['baseline_head']+':'+n],cwd=ROOT)).hexdigest()
  assert actual==h,n
 assert m['cases_sha256']==sha(P/'cases.json')
 assert all(c['exit_code']==0 for c in m['commands']) and len(m['commands'])==30
 nm=read(P/'native'/phase/'manifest.json')
 assert m['binary_sha256']==nm['before_binary_sha256' if phase=='baseline' else 'after_binary_sha256']
 for case in ['54016-stored-1048576','54016-stored-524288','Simple-stored-2097152','synthetic-70000-default','synthetic-100000-default']:
  samples=[]
  for repeat in range(3):
   legs={};rss={}
   for n in [10,1010]:
    stem=f'{case}-{repeat}-{n}';j=read(d/(stem+'.json'));assert j['warmups']==1 and j['samples']==1 and j['queries']==n
    assert j['records'][0]['all_queries_agree'] and len(j['records'][0]['queries'])==n
    legs[n]=counters(d/(stem+'.csv'))
    rss[n]=int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',(d/(stem+'.rss.txt')).read_text())[1])
   # Perf measures both the warmup owner and measured owner plus reporting.
   samples.append(dict(per_extra_query={k:(legs[1010][k]-legs[10][k])/2000 for k in legs[10]},rss_kib=rss))
  rows.append(dict(phase=phase,case=case,samples=samples,median_per_extra_query={k:statistics.median(s['per_extra_query'][k] for s in samples) for k in samples[0]['per_extra_query']},median_rss_kib={str(n):statistics.median(s['rss_kib'][n] for s in samples) for n in [10,1010]}))
comparison=[]
for b in rows[:5]:
 a=next(r for r in rows[5:] if r['case']==b['case'])
 comparison.append(dict(case=b['case'],baseline=b,candidate=a,percent={k:(a['median_per_extra_query'][k]/v-1)*100 if v>0 and a['median_per_extra_query'][k]>=0 else None for k,v in b['median_per_extra_query'].items()},rss_percent={k:(a['median_rss_kib'][k]/v-1)*100 for k,v in b['median_rss_kib'].items()}))
(P/'diagnostics-comparison.json').write_text(json.dumps(comparison,indent=2)+'\n')
for r in comparison:print(r['case'],r['percent'],r['rss_percent'])
print('PASS whole-process bindings; extra-query estimates include semantic projection/reporting and both owners, not native API latency')
