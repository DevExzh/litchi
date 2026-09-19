#!/usr/bin/env python3
"""Validate native diagnostic bindings and report process costs without hiding RSS."""
import hashlib,json,re,statistics
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def counters(p):
 out={}
 for line in p.read_text().splitlines():
  bits=line.split(',')
  if len(bits)>2 and bits[0].isdigit():out[bits[2]]=int(bits[0])
 return out
rows=[]
for phase in ['baseline','candidate']:
 d=P/'diagnostics'/phase;m=read(d/'manifest.json')
 for name,h in m['raw_sha256'].items():assert sha(d/name)==h,(phase,name)
 for name,h in m['probe_sha256'].items():assert sha(P/name)==h,name
 for name,h in m['corpus'].items():assert sha(ROOT/name)==h,name
 if phase=='candidate':
  for name,h in m['source_sha256'].items():assert sha(ROOT/name)==h,name
 commands=read(d/'commands.json');assert len(commands)==36
 assert all(c['exit_code']==0 for c in commands)
 for case in ['54016-stored','WithCustomViews-stored','Simple-stored']:
  for mode in ['owned','file']:
   samples=[]
   for repeat in range(3):
    legs={}
    for n in [10,1010]:
     stem=d/f'{case}-{mode}-{repeat}-{n}'
     values=dict(s.split('=') for s in stem.with_suffix('.tsv').read_text().strip().split('\t'))
     assert int(values['repeats'])==n and int(values['found'])==n
     c=counters(stem.with_suffix('.csv'))
     c['loop_ns']=int(values['nanos'])
     c['max_rss_kib']=int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',stem.with_suffix('.rss.txt').read_text())[1])
     legs[n]=c
    samples.append(dict(per_extra_query={k:(legs[1010][k]-legs[10][k])/1000 for k in legs[10] if k!='max_rss_kib'},rss_kib={str(n):legs[n]['max_rss_kib'] for n in legs}))
   rows.append(dict(phase=phase,case=case,mode=mode,samples=samples,median_per_extra_query={k:statistics.median(s['per_extra_query'][k] for s in samples) for k in samples[0]['per_extra_query']},median_rss_kib={str(n):statistics.median(s['rss_kib'][str(n)] for s in samples) for n in [10,1010]}))
comparison=[]
for b in rows[:6]:
 a=next(r for r in rows[6:] if (r['case'],r['mode'])==(b['case'],b['mode']))
 comparison.append(dict(case=b['case'],mode=b['mode'],before=b,after=a,percent={k:(a['median_per_extra_query'][k]/v-1)*100 if v > 0 and a['median_per_extra_query'][k] >= 0 else None for k,v in b['median_per_extra_query'].items()},rss_percent={k:(a['median_rss_kib'][k]/v-1)*100 for k,v in b['median_rss_kib'].items()}))
(P/'diagnostics-comparison.json').write_text(json.dumps(comparison,indent=2)+'\n')
for r in comparison:print(r['case'],r['mode'],r['percent'],r['rss_percent'])
for phase,m in read(P/'profiles-manifest.json').items():
 for name,h in m['raw_sha256'].items():assert sha(P/name)==h,name
 d=read(P/'diagnostics'/phase/'manifest.json')
 for key in ['binary_sha256','source_sha256','probe_sha256','corpus']:assert m[key]==d[key],(phase,key)
 assert read(P/f'{phase}-profile-command.json')['exit_code']==0
print('Verified profile text, source, binary, repeat-probe and corpus bindings')
for phase,suffix in [('baseline','before'),('candidate','after')]:
 path=Path('/home/zhuhe/code/litchi-target-0684-'+suffix)/'release/xls0684-repeat'
 if path.exists():assert sha(path)==read(P/'diagnostics'/phase/'manifest.json')['binary_sha256']
path=Path('/home/zhuhe/code/litchi-target-0684-after/release/xls-index-budget-probe-0684')
if path.exists():assert sha(path)==read(P/'budgets/manifest.json')['binary_sha256']
print('Available repeat/budget binaries match retained hashes')
