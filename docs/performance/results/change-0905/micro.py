"""Frozen paired native XML-audit guards, separate from Office benefit gates."""
import argparse,hashlib,json,os,random,statistics,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
if not __debug__: raise RuntimeError('Run without Python optimization; custody assertions are mandatory')
PROTOCOL={'cpu':12,'blocks':6,'samples':30,'warmup':5,'iterations':1000,'bootstrap_draws':10000,'seed':905019,'p50_ratio_threshold':1.05,'lower_95_threshold':1.0,'review_ratio':1.05,'spread_ratio':1.2}
CASES=['cardinality-'+str(n) for n in [0,1,2,8,32,33,128]]+['xml-space-valid','xml-space-invalid','duplicate','malformed-tail','malformed-second']
JOBS=[case+'/'+policy for case in CASES for policy in ['source','compact','authored']]

def desc(q):
 q=Path(q);b=q.read_bytes();return {'path':str(q.resolve()),'bytes':len(b),'sha256':hashlib.sha256(b).hexdigest()}
def write(q,v):
 with q.open('x') as f:json.dump(v,f,indent=2,sort_keys=True);f.write('\n')
def read(q):return json.loads(q.read_text())
def checked(d):
 actual=desc(d['path']);assert actual==d,(actual,d);return actual

def load_plan():
 plan=read(P/'micro-plan.json');assert plan['schema']=='litchi.xml-attribute-micro-plan-0905.v1' and plan['status']=='frozen';assert plan['protocol']==PROTOCOL and plan['jobs']==JOBS;checked(plan['helper']);checked(plan['probe']);checked(plan['method']);return plan

def expected_start(plan,block,role):
 binary=checked(plan['binaries'][role]);argv=['taskset','-c','12',binary['path'],'--samples','30','--warmup','5','--iterations','1000']
 if block%2:argv.append('--reverse')
 return {'block':block,'role':role,'argv':argv,'binary':binary,'plan':desc(P/'micro-plan.json'),'helper':desc(__file__)}

def capture():
 plan=load_plan()
 dest=P/'micro';dest.mkdir(exist_ok=False)
 for block in range(6):
  roles=['baseline','candidate'] if block%2==0 else ['candidate','baseline']
  for role in roles:
   stem=dest/f'b{block:02}-{role}';start=expected_start(plan,block,role);argv=start['argv'];binary=start['binary'];write(stem.with_suffix('.start.json'),start)
   with stem.with_suffix('.tsv').open('x') as out,stem.with_suffix('.stderr').open('x') as err:
    r=subprocess.run(argv,stdout=out,stderr=err,env={'PATH':'/usr/bin:/bin','LC_ALL':'C','LANG':'C','TZ':'UTC'},cwd=P.parents[3])
   start.update(exit_code=r.returncode,stdout=desc(stem.with_suffix('.tsv')),stderr=desc(stem.with_suffix('.stderr')));write(stem.with_suffix('.receipt.json'),start);assert r.returncode==0
   checked(binary)
   print('micro',block,role,'pass',flush=True)

def reduce():
 plan=load_plan();rows={};identity=None
 expected={f'b{block:02}-{role}{suffix}' for block in range(6) for role in ['baseline','candidate'] for suffix in ['.start.json','.tsv','.stderr','.receipt.json']}
 assert {q.name for q in (P/'micro').iterdir()}==expected
 for block in range(6):
  for role in ['baseline','candidate']:
   stem=P/'micro'/f'b{block:02}-{role}';receipt=read(stem.with_suffix('.receipt.json'));start=expected_start(plan,block,role);assert read(stem.with_suffix('.start.json'))==start
   for key,value in start.items():assert receipt[key]==value
   assert receipt['exit_code']==0 and receipt['stdout']==desc(stem.with_suffix('.tsv')) and receipt['stderr']==desc(stem.with_suffix('.stderr'))
   comments=[];cases={}
   for line in stem.with_suffix('.tsv').read_text().splitlines():
    if line.startswith(('# input\t','# outcome\t')):comments.append(line)
    if line.startswith('#') or line.startswith('case\t'):continue
    case,policy,sample,ns,iterations=line.split('\t');key=case+'/'+policy;sample=int(sample);ns=int(ns);assert int(iterations)==1000 and ns>0;cases.setdefault(key,[]).append((sample,ns))
   if identity is None:identity=comments
   assert comments==identity
   assert list(cases)==(list(reversed(JOBS)) if block%2 else JOBS)
   for key,samples in cases.items():
    assert [s[0] for s in samples]==list(range(30));values=sorted(s[1] for s in samples);rows.setdefault(key,{}).setdefault(role,[]).append({'block':block,'p50':statistics.median(values),'p95':values[28],'p99':values[29]})
 results={};rng=random.Random(905019)
 for key,roles in sorted(rows.items()):
  b=roles['baseline'];c=roles['candidate'];assert [x['block'] for x in b]==list(range(6)) and [x['block'] for x in c]==list(range(6));ratios=[y['p50']/x['p50'] for x,y in zip(b,c)];boot=sorted(statistics.median([ratios[rng.randrange(6)] for _ in range(6)]) for _ in range(10000));med=statistics.median(ratios);low=boot[249];high=boot[9749]
  adverse=med>1.05 and low>1.0
  results[key]={'baseline':b,'candidate':c,'paired_p50_ratios':ratios,'median_paired_p50_ratio':med,'bootstrap_95':[low,high],'material_adverse_regression':adverse,'tail_review_flags':{metric:statistics.median(y[metric]/x[metric] for x,y in zip(b,c))>1.05 for metric in ['p95','p99']},'spread_flags':{role:max(x['p50'] for x in values)/min(x['p50'] for x in values)>1.2 for role,values in roles.items()}}
 return {'schema':'litchi.xml-attribute-micro-0905.v1','status':'pass','plan':desc(P/'micro-plan.json'),'cases':results,'children':12,'samples':12*36*30,'material_adverse_cases':[key for key,v in results.items() if v['material_adverse_regression']],'adoption_allowed':not any(v['material_adverse_regression'] for v in results.values()),'scope':'per-sample elapsed is1000complete slice audits; output formatting and last-outcome check excluded; separate from Office benefit estimates'}

if __name__=='__main__':
 a=argparse.ArgumentParser();a.add_argument('mode',choices=['capture','reduce']);a.add_argument('--output',type=Path);args=a.parse_args()
 if args.mode=='capture':capture()
 else:
  assert args.output;write(args.output,reduce())
