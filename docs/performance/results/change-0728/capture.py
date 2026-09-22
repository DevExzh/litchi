#!/usr/bin/env python3
"""Run fixed policy baseline matrix only after frozen source/oracle qualification."""
import hashlib,json,subprocess,sys,time
from datetime import datetime,timezone
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def bindings():
 paths=[P/n for n in ['plan.json','cases.json','builds.json','capture.py','analyze.py','audit.py','hypothesis.md','constraints.json','environment.json','qualification.json','oracle-contract.json','qualify.py','quality.py','prepare.py','preflight.py','preflight.json','negative-checks.py','report-tables.py']]
 paths.extend(P/rel for rel in read(P/'qualification.json')['files'])
 b=read(P/'builds.json');paths.extend(ROOT/rel for rel in b['source_sha256']);paths.extend(P/rel for rel in b['probe_sha256']);paths.extend(ROOT/c['path'] for c in read(P/'cases.json'));paths.extend(ROOT/rel for rel in read(P/'constraints.json'))
 return {str(p.relative_to(ROOT)):sha(p) for p in sorted(set(paths))}
def binaries():
 rows=read(P/'builds.json')['binaries']
 for r in rows:assert sha(Path(r['path']))==r['sha256'] and Path(r['path']).stat().st_size==r['bytes']
 return rows
if sys.argv[1]=='freeze':
 assert not (P/'freeze.json').exists();write(P/'freeze.json',dict(bindings=bindings(),binaries=binaries()));print('frozen');sys.exit(0)
assert sys.argv[1]=='run';plan=read(P/'plan.json');f=read(P/'freeze.json');assert f['bindings']==bindings() and f['binaries']==binaries();out=P/'captures';out.mkdir(exist_ok=False);m=dict(started_utc=datetime.now(timezone.utc).isoformat(),status='running',freeze_sha256=sha(P/'freeze.json'),bindings_start=bindings(),runs=[]);write(out/'manifest.json',m)
cases=read(P/'cases.json');routes=[('format','reuse'),('container','reuse'),('container','rewrite')]
for lane in ['native','allocation']:
 binary=next(b['path'] for b in f['binaries'] if Path(b['path']).name==('ole_format_save_probe' if lane=='native' else 'ole_format_save_probe_alloc'))
 for cycle in range(plan['cycles'] if lane=='native' else 1):
  cells=[(c,op,policy) for c in cases for op,policy in routes]
  if cycle%2:cells.reverse()
  for c,op,policy in cells:
   for repeat in range(3):
    name=f"{lane}-c{cycle}-{c['case']}-{op}-{policy}-r{repeat}.json";cmd=['taskset','-c',str(plan['cpu']),binary,'--case',c['case'],'--input',c['path'],'--operation',op,'--policy',policy,'--samples',str(plan['samples'] if lane=='native' else 1),'--warmups',str(plan['warmups'] if lane=='native' else 0)]
    started=time.monotonic();started_utc=datetime.now(timezone.utc).isoformat()
    with (out/name).open('wb') as o,(out/(name+'.stderr')).open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=o,stderr=e)
    m['runs'].append(dict(started_utc=started_utc,elapsed_seconds=time.monotonic()-started,lane=lane,cycle=cycle,case=c['case'],operation=op,policy=policy,repeat=repeat,command=cmd,exit_code=r.returncode,output=name,sha256=sha(out/name),stderr=name+'.stderr',stderr_sha256=sha(out/(name+'.stderr'))));write(out/'manifest.json',m);assert r.returncode==0
   print(lane,cycle,c['case'],op,policy,flush=True)
assert f['binaries']==binaries();m['bindings_end']=bindings();assert m['bindings_start']==m['bindings_end'];m['finished_utc']=datetime.now(timezone.utc).isoformat();m['status']='complete';write(out/'manifest.json',m)
