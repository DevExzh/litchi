#!/usr/bin/env python3
"""Freeze and execute the fixed public DOC attribution matrix."""
import hashlib,json,subprocess,sys,time
from datetime import datetime,timezone
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def bindings():
 names=['plan.json','cases.json','builds.json','capture.py','analyze.py','hypothesis.md','constraints.json','environment.json','qualification.json','oracle-contract.json','qualify.py','quality.py','prepare.py','preflight.py','preflight.json','negative-checks.py','report-tables.py','audit.py','prior-oracle-contract.json','ancestry.json']
 paths=[P/n for n in names];paths.extend(P.parent/'change-0728'/n for n in ['artifact-manifest.json','oracle-contract.json','builds.json']);b=read(P/'builds.json')
 paths.extend(P/rel for rel in read(P/'qualification.json')['files']);paths.extend(ROOT/rel for rel in b['source_sha256']);paths.extend(P/rel for rel in b['probe_sha256']);paths.extend(ROOT/c['path'] for c in read(P/'cases.json'));paths.extend(ROOT/rel for rel in read(P/'constraints.json'))
 return {str(p.relative_to(ROOT)):sha(p) for p in sorted(set(paths))}
def binaries():
 rows=read(P/'builds.json')['binaries'];assert len(rows)==1
 for r in rows:assert sha(Path(r['path']))==r['sha256'] and Path(r['path']).stat().st_size==r['bytes']
 return rows
if sys.argv[1]=='freeze':
 assert not (P/'freeze.json').exists();write(P/'freeze.json',dict(bindings=bindings(),binaries=binaries()));print('frozen');sys.exit(0)
assert sys.argv[1]=='run';plan=read(P/'plan.json');f=read(P/'freeze.json');assert f['bindings']==bindings() and f['binaries']==binaries();out=P/'captures';out.mkdir(exist_ok=False)
m=dict(status='running',started_utc=datetime.now(timezone.utc).isoformat(),freeze_sha256=sha(P/'freeze.json'),bindings_start=bindings(),runs=[]);write(out/'manifest.json',m)
cases={c['case']:c for c in read(P/'cases.json')};binary=f['binaries'][0]['path']
for row in plan['schedule']:
 c=cases[row['case']];name=f"c{row['cycle']}-{row['case']}-{row['route']}-r{row['repeat']}.json"
 cmd=['taskset','-c',str(plan['cpu']),binary,'--case',row['case'],'--input',c['path'],'--route',row['route'],'--samples',str(plan['samples']),'--warmups',str(plan['warmups'])]
 start=time.monotonic();started_utc=datetime.now(timezone.utc).isoformat()
 with (out/name).open('wb') as o,(out/(name+'.stderr')).open('wb') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=o,stderr=e)
 m['runs'].append(dict(row,command=cmd,started_utc=started_utc,elapsed_seconds=time.monotonic()-start,exit_code=r.returncode,output=name,sha256=sha(out/name),stderr=name+'.stderr',stderr_sha256=sha(out/(name+'.stderr'))));write(out/'manifest.json',m);assert r.returncode==0
 print(row['cycle'],row['repeat'],row['case'],row['route'],flush=True)
assert f['binaries']==binaries();m['bindings_end']=bindings();assert m['bindings_start']==m['bindings_end'];m['finished_utc']=datetime.now(timezone.utc).isoformat();m['status']='complete';write(out/'manifest.json',m)
