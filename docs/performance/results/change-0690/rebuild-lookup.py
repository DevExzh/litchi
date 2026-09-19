#!/usr/bin/env python3
"""Rebuild the corrected supplemental probe against both frozen source revisions."""
import hashlib,json,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
base=json.loads((P/'baseline.json').read_text())['baseline_head']
paths=[ROOT/n for n in ['crates/litchi-cfb/src/directory_name.rs','crates/litchi-cfb/src/file.rs','crates/litchi-cfb/src/shared.rs']]
original={p:p.read_bytes() for p in paths};commands=[]
def run(name,*args):
 cmd=['python3',str(P/name),*args];start=time.monotonic()
 with (P/('lookup-final-'+name.removesuffix('.py')+('-'+args[0] if args else '')+'.log')).open('w') as f:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=subprocess.STDOUT)
 commands.append(dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start));(P/'lookup-final-commands.json').write_text(json.dumps(commands,indent=2)+'\n');print(name,args,r.returncode,flush=True);assert r.returncode==0
expected=json.loads((P/'candidate-builds.json').read_text())[0]['source_sha256']
assert all(hashlib.sha256((ROOT/n).read_bytes()).hexdigest()==h for n,h in expected.items())
try:
 for p in paths:p.write_bytes(subprocess.check_output(['git','show',base+':'+str(p.relative_to(ROOT))],cwd=ROOT))
 run('build-lookup.py','baseline')
 run('measure-lookup.py','baseline')
finally:
 for p,data in original.items():p.write_bytes(data)
 assert all(hashlib.sha256((ROOT/n).read_bytes()).hexdigest()==h for n,h in expected.items())
run('build-lookup.py','candidate')
run('measure-lookup.py','candidate')
run('audit-lookup.py')
