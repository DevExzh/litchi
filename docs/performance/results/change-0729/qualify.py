#!/usr/bin/env python3
"""Retain every public DOC route qualification before measurement freeze."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
i=0
while (P/f'qualification-{i}').exists():i+=1
out=P/f'qualification-{i}';out.mkdir();build=read(P/'builds.json');assert len(build['binaries'])==1;b=build['binaries'][0];rows=[];identities={};prior=read(P/'prior-oracle-contract.json')
assert sha(Path(b['path']))==b['sha256']
for c in read(P/'cases.json'):
 assert sha(ROOT/c['path'])==c['sha256'] and (ROOT/c['path']).stat().st_size==c['bytes']
 for route in read(P/'plan-draft.json')['routes']:
  name=f"{c['case']}-{route}";cmd=['taskset','-c','12',b['path'],'--case',c['case'],'--input',c['path'],'--route',route,'--samples','1','--warmups','0']
  with (out/(name+'.json')).open('wb') as stdout,(out/(name+'.stderr')).open('wb') as stderr:r=subprocess.run(cmd,cwd=ROOT,stdout=stdout,stderr=stderr)
  rows.append(dict(command=cmd,exit_code=r.returncode,output=name+'.json',sha256=sha(out/(name+'.json')),stderr=name+'.stderr',stderr_sha256=sha(out/(name+'.stderr'))));write(out/'manifest.json',dict(builds_sha256=sha(P/'builds.json'),script_sha256=sha(Path(__file__)),runs=rows))
  print(name,r.returncode,flush=True)
  assert r.returncode==0
  x=read(out/(name+'.json'));assert (x['case'],x['route'],x['samples_requested'],x['warmups'])==(c['case'],route,1,0)
  identity={k:x[k] for k in prior[c['case']]['identity']};assert identity==prior[c['case']]['identity'] and identity==identities.setdefault(c['case'],identity)
  assert x['expected_oracle']['semantic_witness']==prior[c['case']]['semantic_witness']
  assert x['oracle_controls'] and all(o['rejected'] and o['status']=='rejected' and o['failure_reasons'] for o in x['oracle_controls'])
  assert len(x['samples'])==1
  for oracle in [x['expected_oracle']]+[s['oracle'] for s in x['samples']]:assert all(v for v in oracle.values() if isinstance(v,bool)) and not oracle['failure_reasons']
print('PASS eight routes; exact inherited DOC identity/oracles')
