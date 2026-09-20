#!/usr/bin/env python3
"""Freeze and run the predeclared diagnostic without overwriting any sample."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def bindings(plan):
 paths={P/n for n in ['plan.json','build.py','capture.py','analyze.py','constraints.json','environment.json','builds.json','build-restoration.json','hypothesis.md']}
 paths.update(P.glob('source-*.json'));paths.update(p for p in (P/'sources').rglob('*') if p.is_file())
 for owner in ['litchi-cfb','litchi-xls']:paths.update(p for p in (ROOT/'crates'/owner).rglob('*') if p.is_file() and p.suffix in ('.rs','.toml'))
 probe=P.parent/'change-0686/probe';paths.update([probe/'Cargo.toml',probe/'Cargo.lock']);paths.update(probe.glob('src/**/*.rs'));paths.update(ROOT/c['path'] for c in plan['cases'])
 for rel in read(P/'constraints.json'):paths.add(ROOT/rel)
 return {str(p.relative_to(ROOT)):sha(p) for p in sorted(paths)}
def binaries():
 rows=read(P/'builds.json');assert len(rows)==2
 for r in rows:assert r['exit_code']==0 and Path(r['binary']).stat().st_size==r['bytes'] and sha(Path(r['binary']))==r['binary_sha256']
 return [{k:r[k] for k in ['phase','binary','binary_sha256','bytes']} for r in rows]
plan=read(P/'plan.json');freeze=P/'freeze.json'
if sys.argv[1]=='freeze':
 assert not freeze.exists() and read(P/'build-restoration.json')['exact'];write(freeze,dict(bindings=bindings(plan),binaries=binaries(),baseline_head=plan['baseline_revision']));print('frozen');sys.exit(0)
assert sys.argv[1]=='run';f=read(freeze);assert f['bindings']==bindings(plan) and f['binaries']==binaries()
out=P/'captures';out.mkdir(exist_ok=False);m=dict(status='running',freeze_sha256=sha(freeze),bindings_start=bindings(plan),runs=[]);write(out/'manifest.json',m)
for cycle in range(plan['cycles']):
 for leg in plan['legs']:
  phase='candidate' if leg in ['b1','b2'] else 'baseline';binary=next(x['binary'] for x in f['binaries'] if x['phase']==phase)
  for c in plan['cases']:
   for replicate in range(plan['processes_per_cell_leg']):
    name=f"c{cycle}-{leg}-{c['id']}-r{replicate}.json";cmd=['taskset','-c',str(plan['cpu']),binary,'--input',c['path'],'--budget',str(c['budget']),'--mode',c['mode'],'--worksheet',str(c['sheet']),'--row',str(c['row']),'--column',str(c['column']),'--queries',str(plan['queries']),'--warmups',str(plan['warmups']),'--samples',str(plan['samples'])]
    with (out/name).open('wb') as o,(out/(name+'.stderr')).open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=o,stderr=e)
    m['runs'].append(dict(cycle=cycle,leg=leg,phase=phase,case=c['id'],replicate=replicate,command=cmd,exit_code=r.returncode,output=name,sha256=sha(out/name),stderr=name+'.stderr',stderr_sha256=sha(out/(name+'.stderr'))));write(out/'manifest.json',m);assert r.returncode==0
  print('cycle',cycle,leg,'complete',flush=True)
assert f['binaries']==binaries();m['bindings_end']=bindings(plan);assert m['bindings_end']==m['bindings_start'];m['status']='complete';write(out/'manifest.json',m)
