#!/usr/bin/env python3
import hashlib,json,math,random,statistics,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
base=read(P/'baseline.json')
for phase in ['baseline','candidate']:
 d=P/'native'/phase;m=read(d/'manifest.json')
 for n,h in m['raw_sha256'].items():assert sha(d/n)==h,n
 for field in ['probe_sha256','corpus']:
  for n,h in m[field].items():assert sha(ROOT/n)==h,n
 assert sha(P/'cases.json')==m['cases_sha256']
 for n,h in m['source_sha256'].items():
  actual=sha(ROOT/n) if phase=='candidate' else hashlib.sha256(subprocess.check_output(['git','show',base['baseline_head']+':'+n],cwd=ROOT)).hexdigest()
  assert actual==h,n
 assert all(r['exit_code']==0 for r in read(d/'commands.json'))
b=read(P/'native/baseline/manifest.json');a=read(P/'native/candidate/manifest.json');assert b['before_binary_sha256']==a['before_binary_sha256']
for phase,key in [('before','before_binary_sha256'),('after','after_binary_sha256')]:
 f=Path('/home/zhuhe/code/litchi-target-0688-'+phase)/'release/xls-index-retry-probe-0686'
 if f.exists():assert sha(f)==a[key]
def stats(v):
 s=sorted(v);return dict(n=len(s),p50=statistics.median(s),mean=statistics.mean(s),p95=s[math.ceil(.95*len(s))-1],p99=s[math.ceil(.99*len(s))-1])
def pct(a,b):return (a/b-1)*100
rng=random.Random(688)
def ci(a,b):
 draws=sorted(pct(statistics.median(rng.choices(a,k=len(a))),statistics.median(rng.choices(b,k=len(b)))) for _ in range(1000))
 return [draws[24],draws[974]]
rows=[]
for c in read(P/'cases.json'):
 for mode in ['owned','file']:
  values={};outcomes=[]
  for leg in ['aa1','aa2','a1','b1','b2','a2']:
   phase='baseline' if leg.startswith('aa') else 'candidate'
   j=read(P/'native'/phase/f"{leg}-{c['case']}-{mode}.json")
   assert j['input_sha256']==a['corpus'][c['path']]
   assert j['max_query_index_bytes']==c['budget'] and j['mode']==mode
   assert (j['worksheet'],j['row'],j['column'])==(c['sheet'],c['row'],c['column'])
   assert j['samples']==100 and j['warmups']==3 and j['queries']==8
   r=j['records'];assert [x['sample'] for x in r]==list(range(3,103))
   assert all(x['all_queries_agree'] and x['open']['outcome']['status']=='ok' and [q['ordinal'] for q in x['queries']]==list(range(8)) for x in r)
   outcome=r[0]['queries'][0]['outcome'];assert all(q['outcome']==outcome for x in r for q in x['queries']);outcomes.append(outcome)
   metrics={'open':[x['open']['elapsed_ns'] for x in r]}
   for k in [1,2,3,8]:metrics[f'q{k}']=[x['queries'][k-1]['elapsed_ns'] for x in r]
   metrics['q3-to-q8-mean']=[statistics.mean(q['elapsed_ns'] for q in x['queries'][2:]) for x in r]
   metrics['open-plus-eight']=[x['open']['elapsed_ns']+sum(q['elapsed_ns'] for q in x['queries']) for x in r]
   values[leg]=metrics
  assert all(o==outcomes[0] for o in outcomes)
  for metric in values['a1']:
   t={leg:stats(v[metric]) for leg,v in values.items()}
   changes={key:pct(t[x]['p50'],t[y]['p50']) for key,x,y in [('b1_a1','b1','a1'),('b2_a2','b2','a2'),('aa','aa2','aa1'),('abba_a','a2','a1'),('abba_b','b2','b1')]}
   row=dict(case=c['case'],mode=mode,metric=metric,timing=t,percent_p50=changes,paired_p50_bootstrap95={'b1_a1':ci(values['b1'][metric],values['a1'][metric]),'b2_a2':ci(values['b2'][metric],values['a2'][metric])},review_regression=changes['b1_a1']>5 and changes['b2_a2']>5)
   rows.append(row)
(P/'native-comparison.json').write_text(json.dumps(rows,indent=2)+'\n')
print('PASS native source/probe/corpus/raw bindings and exact semantic parity; 24 groups, 14,400 fresh owners, 115,200 query records')
for r in rows:
 if r['metric'] in ['q8','open-plus-eight'] or r['review_regression']:print('REVIEW' if r['review_regression'] else 'RESULT',r['case'],r['mode'],r['metric'],{k:round(v,2) for k,v in r['percent_p50'].items()})
