#!/usr/bin/env python3
"""Longer paired controls for flagged tiny phases and baseline drift."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];before,after=sys.argv[1:];D=P/'timing-diagnostics';D.mkdir(exist_ok=True)
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
commands=[]
groups=[('54016-missing','visit'),('Simple-missing','prepared'),('Simple-stored','prepared'),('15228-stored','prepared'),('WithCustomViews-stored','prepared'),('54016-stored','prepared')]
cases=json.loads((P/'cases.json').read_text())
for name,route in groups:
 c=next(c for c in cases if c['case']==name)
 for leg,binary in [('a1',before),('b1',after),('b2',after),('a2',before)]:
  command=['taskset','-c','12',binary,'route','--input',c['path'],'--route',route,'--mode','owned-native','--worksheet',str(c['sheet']),'--row',str(c['row']),'--column',str(c['column']),'--second-row',str(c['second_row']),'--second-column',str(c['second_column']),'--warmups','10','--samples','200']
  r=subprocess.run(command,cwd=ROOT,capture_output=True,text=True);out=f'{leg}-{name}-{route}.json';(D/out).write_text(r.stdout);commands.append(dict(command=command,output=out,exit_code=r.returncode));assert r.returncode==0,r.stderr
 print(name,route,flush=True)
(D/'commands.json').write_text(json.dumps(commands,indent=2)+'\n')
base=json.loads((P/'measurements/candidate/manifest.json').read_text());m=dict(source_sha256=base['source_sha256'],probe_sha256=base['probe_sha256'],before_binary_sha256=sha(before),after_binary_sha256=sha(after),cases_sha256=sha(P/'cases.json'),samples=200,warmups=10,raw_sha256={p.name:sha(p) for p in D.iterdir() if p.is_file() and p.name!='manifest.json'});(D/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
