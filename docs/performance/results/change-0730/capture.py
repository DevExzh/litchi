#!/usr/bin/env python3
"""Serial, immutable-schedule ordinary DOC comparison; root runs all natives."""
import hashlib,json,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def bindings():
 names=['hypothesis.md','plan.json','cases.json','constraints.json','capture.py','probe-design.md','baseline-builds.json','candidate-builds.json','candidate-source.json','source-guard.py','oracle-contract.json','ancestry.json','analyze.py','audit.py','negative-checks.py','quality.py','qualification.json','retention.py','retention-builds.json','before-qualification.json','environment.json','hypothesis.md','preflight.py','qualify.py']
 result={str(P/n):sha(P/n) for n in names}
 for variant in ['baseline','candidate','retention']:
  b=read(P/(variant+'-builds.json'))
  for rel in b['probe_sha256']:result[str(P/rel)]=sha(P/rel)
  if 'receipt' in b:
   for rel in [b['receipt'],b['receipt'].replace('.json','.log')]:result[str(P/rel)]=sha(P/rel)
  for row in b['binaries']:assert sha(Path(row['path']))==row['sha256'];result[row['path']]=row['sha256']
 for rel,expected in read(P/'candidate-builds.json')['source_sha256'].items():
  archived=P/'candidate-source'/rel;path=archived if archived.exists() else ROOT/rel;assert sha(path)==expected;result[str(path)]=expected
 for c in read(P/'cases.json'):assert sha(ROOT/c['path'])==c['sha256'];result[str(ROOT/c['path'])]=c['sha256']
 for rel,expected in read(P/'qualification.json')['files'].items():assert sha(P/rel)==expected;result[str(P/rel)]=expected
 for directory in ['baseline-source','candidate-source']:
  for path in (P/directory).rglob('*'):
   if path.is_file():result[str(path)]=sha(path)
 return result
mode=sys.argv[1];assert mode in ['before','freeze','run']
if mode in ['freeze','run']:subprocess.run([sys.executable,str(P/'source-guard.py')],cwd=ROOT,check=True)
if mode=='freeze':
 assert not (P/'freeze.json').exists();write(P/'freeze.json',bindings());print('frozen');sys.exit()
if mode=='run':
 assert read(P/'freeze.json')==bindings()
 preflight=read(P/'preflight.json');assert preflight['status']=='passed'
 for key,name in [('analyzer_sha256','analyze.py'),('auditor_sha256','audit.py'),('contract_sha256','oracle-contract.json'),('script_sha256','preflight.py')]:assert preflight[key]==sha(P/name)
out=P/('before' if mode=='before' else 'captures');out.mkdir(exist_ok=False)
plan=read(P/'plan.json');cases=read(P/'cases.json');runs=[];manifest={'mode':mode,'status':'running','runs':runs}
if mode=='run':manifest['freeze_sha256']=sha(P/'freeze.json');manifest['preflight_sha256']=sha(P/'preflight.json')
if mode=='before':schedule=[('native',0,c,v,0) for c in cases for v in ['baseline']]+[('allocation',0,c,'baseline',0) for c in cases]
else:
 schedule=[]
 for cycle in range(3):
  for c in cases[::1 if cycle%2==0 else -1]:
   for slot,v in enumerate(['baseline','baseline','baseline','candidate','candidate','baseline']):schedule.append(('native',cycle,c,v,slot))
 for c in cases:
  for slot,v in enumerate(['baseline','candidate','candidate','baseline']):schedule.append(('allocation',0,c,v,slot))
for lane,cycle,c,variant,slot in schedule:
 b=read(P/(variant+'-builds.json'));suffix='ole_format_save_probe'+('_alloc' if lane=='allocation' else '')
 row=next(r for r in b['binaries'] if Path(r['path']).name==variant+'-'+suffix);assert sha(Path(row['path']))==row['sha256']
 name=f'{lane}-c{cycle}-{c["case"]}-{slot}-{variant}.json'
 cmd=['taskset','-c',str(plan['cpu']),row['path'],'--case',c['case'],'--input',c['path'],'--operation','format','--samples',str(plan['samples'] if lane=='native' else 1),'--warmups',str(plan['warmups'] if lane=='native' else 0)]
 start=time.monotonic()
 with (out/name).open('wb') as f,(out/(name+'.stderr')).open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=e)
 runs.append(dict(lane=lane,cycle=cycle,case=c['case'],variant=variant,slot=slot,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,output=name,sha256=sha(out/name),stderr_sha256=sha(out/(name+'.stderr'))))
 write(out/'manifest.json',manifest);assert r.returncode==0
 print(name,flush=True)
if mode=='run':
 assert read(P/'freeze.json')==bindings()
 preflight=read(P/'preflight.json');assert preflight['status']=='passed'
 for key,name in [('analyzer_sha256','analyze.py'),('auditor_sha256','audit.py'),('contract_sha256','oracle-contract.json'),('script_sha256','preflight.py')]:assert preflight[key]==sha(P/name)
manifest['status']='complete';write(out/'manifest.json',manifest)
