#!/usr/bin/env python3
"""Serial four-route qualification and prospective native phase capture."""
import hashlib,json,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def guard():
 b=read(P/'build.json');assert all(sha(ROOT/f)==h for f,h in b['source'].items());assert all(sha(P/f)==h for f,h in b['probe'].items());assert all(sha(ROOT/f)==h for f,h in read(P/'constraints.json').items());assert sha(P/b['quality'])==b['quality_sha256'];q=read(P/b['quality']);assert len(q['runs'])==12 and all(r['exit_code']==0 and sha(P/r['output'])==r['sha256'] for r in q['runs']);assert q['source']==b['source'] and q['probe']==b['probe'];r=b['binary'];assert sha(Path(r['path']))==r['sha256'];return b
plan=read(P/'plan.json');case=read(P/'case.json');mode=sys.argv[1];assert mode in ['qualify','freeze','capture'];b=guard()
def bindings():
 paths=[P/n for n in ['run.py','quality.py','analyze.py','audit.py','preflight.py','source-guard.py','plan.json','hypothesis.md','environment.json','constraints.json','case.json','oracle.json','ancestry.json','build.json','source-review.json','qualification.json']]
 paths+=list((P/'probe').rglob('*'))+list((P/'qualification').rglob('*'))+list((P/Path(b['quality']).parent).rglob('*'))+list((P/'source-archive').rglob('*'));paths+=[ROOT/f for f in b['source']]+[ROOT/case['path'],Path(b['binary']['path'])]
 return {str(p):sha(p) for p in paths if p.is_file()}
def command(route,samples,warmups):return ['taskset','-c','12',b['binary']['path'],'--route',route,'--input',case['path'],'--samples',str(samples),'--warmups',str(warmups)]
def execute(cmd,out):
 start=time.monotonic()
 with out.open('wb') as f,Path(str(out)+'.stderr').open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=e)
 return dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,output=out.name,sha256=sha(out),stderr_sha256=sha(Path(str(out)+'.stderr')))
if mode=='qualify':
 out=P/'qualification';out.mkdir();rows=[];expected=read(P/'oracle.json')['expected']
 for route in plan['routes']:
  row=execute(command(route,1,1),out/(route+'.json'));row['route']=route;rows.append(row);write(out/'manifest.json',dict(status='running',runs=rows));assert row['exit_code']==0
  x=read(out/row['output'])
  for k in ['case','format','operation','input','source_sha256','expected_output_sha256','replacements_sha256','source_inventory','expected_output_inventory','replacements','changed_length_proof','expected_oracle','oracle_controls']:assert x[k]==expected[k],k
  assert len(x['samples'])==1
  for sample in x['samples']:assert sample['output_sha256']==expected['expected_output_sha256'] and sample['oracle']==expected['expected_oracle'] and sample['output_inventory']==expected['expected_output_inventory']
 write(out/'manifest.json',dict(status='passed',runs=rows));write(P/'qualification.json',dict(status='passed',build_sha256=sha(P/'build.json'),files={str(p.relative_to(P)):sha(p) for p in out.iterdir() if p.is_file()}));print('PASS four exact0728oracle route smoke processes')
elif mode=='freeze':
 assert read(P/'qualification.json')['status']=='passed';assert not (P/'freeze.json').exists();write(P/'freeze.json',bindings());print('frozen')
else:
 assert read(P/'freeze.json')==bindings()
 preflight=read(P/'preflight.json');assert preflight['status']=='passed';assert preflight['freeze_sha256']==sha(P/'freeze.json')
 for name in ['preflight.py','analyze.py','audit.py']:assert preflight['scripts'][name]==sha(P/name)
 preflight_sha=sha(P/'preflight.json');out=P/'captures';out.mkdir();rows=[]
 for spec in plan['schedule']:
  name=f'c{spec["cycle"]}-r{spec["repeat"]}-{spec["route"]}.json';r=execute(command(spec['route'],plan['samples'],plan['warmups']),out/name);r.update(spec);rows.append(r);write(out/'manifest.json',dict(status='running',freeze_sha256=sha(P/'freeze.json'),preflight_sha256=preflight_sha,runs=rows));print(name,r['exit_code'],flush=True);assert r['exit_code']==0
 assert read(P/'freeze.json')==bindings();guard();write(out/'manifest.json',dict(status='complete',freeze_sha256=sha(P/'freeze.json'),preflight_sha256=preflight_sha,runs=rows))
