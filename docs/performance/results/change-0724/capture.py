#!/usr/bin/env python3
"""Freeze and capture the mirrored diagnostic matrix; never overwrite observations."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def bindings(plan):
 paths={P/n for n in ('plan.json','analyze.py','capture.py','build.py','constraints.json','environment.json','builds.json','build-restoration.json')}
 paths.update(P.glob('source-*.json'));paths.update(p for p in (P/'sources').rglob('*') if p.is_file())
 for owner in ('litchi-cfb','litchi-xls'):paths.update(p for p in (ROOT/'crates'/owner).rglob('*') if p.is_file() and p.suffix in ('.rs','.toml'))
 for folder in ('change-0684/repeat-probe','change-0686/probe'):
  d=P.parent/folder;paths.update([d/'Cargo.toml',d/'Cargo.lock']);paths.update((d/'src').rglob('*.rs'))
 paths.update(ROOT/c['path'] for c in plan['cases'])
 return {str(p.relative_to(ROOT)):sha(p) for p in sorted(paths)}
def binaries():
 result=[]
 for b in json.loads((P/'builds.json').read_text()):
  assert b['exit_code']==0
  path=Path(b['binary']);assert path.stat().st_size==b['bytes'] and sha(path)==b['binary_sha256']
  result.append({k:b[k] for k in ('variant','binary','binary_sha256','bytes')})
 return result
plan=json.loads((P/'plan.json').read_text());command=sys.argv[1];f=P/'freeze.json'
if command=='freeze':
 assert not f.exists();assert json.loads((P/'build-restoration.json').read_text())['exact']
 write(f,dict(bindings=bindings(plan),binaries=binaries(),baseline_head=subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip()));print('frozen');sys.exit(0)
assert command=='run';frozen=json.loads(f.read_text());assert frozen['bindings']==bindings(plan);assert frozen['binaries']==binaries()
out=P/'captures';out.mkdir(exist_ok=False)
m=dict(status='running',freeze_sha256=sha(f),bindings_start=bindings(plan),runs=[]);write(out/'manifest.json',m)
stages=[('b0','baseline'),('l0','layout'),('s0','selection'),('f0','full'),('f1','full'),('s1','selection'),('l1','layout'),('b1','baseline')]
for stage,variant in stages:
 for c in plan['cases']:
  for lane,count in [('native',1),('repeat',plan['repeat']['samples_per_leg'])]:
   binary=next(b['binary'] for b in frozen['binaries'] if b['variant']==variant and Path(b['binary']).name==('xls0684-repeat' if lane=='repeat' else 'xls-index-retry-probe-0686'))
   for sample in range(count):
    name=f"{stage}-{c['case']}-{lane}"+(f'-{sample}.tsv' if lane=='repeat' else '.json')
    cmd=['taskset','-c',str(plan['cpu']),binary]
    if lane=='native':
     cmd+=['--input',c['path'],'--budget',str(c['budget']),'--mode','owned','--worksheet',str(c['sheet']),'--row',str(c['row']),'--column',str(c['column']),'--queries',str(plan['native']['queries']),'--warmups',str(plan['native']['warmups']),'--samples',str(plan['native']['samples'])]
    else:cmd+=['owned',c['path'],str(c['sheet']),str(c['row']),str(c['column']),str(plan['repeat']['repetitions'])]
    with (out/name).open('wb') as stdout,(out/(name+'.stderr')).open('wb') as stderr:r=subprocess.run(cmd,cwd=ROOT,stdout=stdout,stderr=stderr)
    m['runs'].append(dict(stage=stage,variant=variant,case=c['case'],lane=lane,sample=sample if lane=='repeat' else None,command=cmd,exit_code=r.returncode,output=name,sha256=sha(out/name),stderr=name+'.stderr',stderr_sha256=sha(out/(name+'.stderr'))));write(out/'manifest.json',m)
    assert r.returncode==0,m['runs'][-1]
  print(stage,c['case'],'complete',flush=True)
assert binaries()==frozen['binaries'];m['bindings_end']=bindings(plan);assert m['bindings_end']==m['bindings_start'];m['status']='complete';write(out/'manifest.json',m)
