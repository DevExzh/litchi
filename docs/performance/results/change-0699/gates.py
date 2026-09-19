#!/usr/bin/env python3
"""Check the standalone diagnostic and unchanged-production documentation gates."""
import hashlib,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
WORK=ROOT.parent/'litchi-0699-work'
TARGET=ROOT.parent/'litchi-target-0699'
out=P/'gates';out.mkdir(exist_ok=True)
commands=[('workspace-fmt',['cargo','fmt','--all','--check']),('probe-fmt',['cargo','fmt','--manifest-path',str(P/'probe/Cargo.toml'),'--check']),('probe-clippy',['cargo','clippy','--release','--locked','--manifest-path',str(WORK/P.relative_to(ROOT)/'probe/Cargo.toml'),'--target-dir',str(TARGET),'--no-deps','--','-D','warnings'])]
commands += [(r['name'],r['command']) for r in json.loads((P.parent/'change-0698/evidence/results.json').read_text())]
records=[]
for name,command in commands:
 start=time.monotonic();log=out/(name+'.log')
 with log.open('w') as f:
  r=subprocess.run(command,cwd=ROOT,env={**os.environ,'RUSTFLAGS':'-D warnings','CARGO_BUILD_JOBS':'2','CARGO_TARGET_DIR':str(TARGET)},stdout=f,stderr=subprocess.STDOUT)
 records.append(dict(name=name,command=command,exit_code=r.returncode,seconds=time.monotonic()-start,log_sha256=hashlib.sha256(log.read_bytes()).hexdigest()))
 (P/'gates.json').write_text(json.dumps(records,indent=2)+'\n')
 print(name,r.returncode,flush=True)
 assert r.returncode==0,log.read_text()[-3000:]
