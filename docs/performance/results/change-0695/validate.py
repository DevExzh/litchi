#!/usr/bin/env python3
"""Probe quality and applicable repository evidence gates; no production edits."""
import hashlib,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
out=P/'validation';out.mkdir(exist_ok=True)
manifest=str(P/'probe/Cargo.toml');target=str(ROOT.parent/'litchi-target-0695')
rows=[]
commands=[('probe-fmt',['cargo','fmt','--manifest-path',manifest,'--','--check']),('probe-clippy',['cargo','clippy','--locked','--offline','--manifest-path',manifest,'--target-dir',target,'-j','2','--','-D','warnings']),('probe-doc',['cargo','doc','--no-deps','--locked','--offline','--manifest-path',manifest,'--target-dir',target,'-j','2'])]
commands += [(r['name'],r['command']) for r in json.loads((P.parent/'change-0694/evidence/results.json').read_text())]
for name,command in commands:
 path=out/(name+'.log');started=time.monotonic()
 with path.open('w') as log:
  r=subprocess.run(command,cwd=ROOT,env={**os.environ,'RUSTFLAGS':'-D warnings','RUSTDOCFLAGS':'-D warnings','CARGO_TARGET_DIR':target,'CARGO_BUILD_JOBS':'2'},stdout=log,stderr=subprocess.STDOUT)
 rows.append(dict(name=name,command=command,exit_code=r.returncode,seconds=time.monotonic()-started,log_sha256=hashlib.sha256(path.read_bytes()).hexdigest()))
 (P/'validation.json').write_text(json.dumps(rows,indent=2)+'\n')
 print(name,r.returncode,flush=True)
 assert r.returncode==0,path.read_text()
