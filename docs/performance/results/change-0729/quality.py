#!/usr/bin/env python3
"""Serial focused probe quality checks, retaining each attempt."""
import hashlib,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
i=0
while (P/f'quality-{i}').exists():i+=1
out=P/f'quality-{i}';out.mkdir()
manifest=str(P/'probe/Cargo.toml');env=dict(os.environ,CARGO_TARGET_DIR=str(ROOT.parent/'litchi-target-0729'),CARGO_BUILD_JOBS='2',RUSTDOCFLAGS='-D warnings')
commands=[['cargo','fmt','--manifest-path',manifest,'--','--check'],['cargo','test','--manifest-path',manifest,'--release','--offline','--locked','--lib'],['cargo','clippy','--manifest-path',manifest,'--release','--offline','--locked','--all-targets','--','-D','warnings'],['cargo','doc','--manifest-path',manifest,'--release','--offline','--locked','--no-deps']]
rows=[]
for n,cmd in enumerate(commands):
 started=time.monotonic()
 with (out/f'{n}.log').open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append(dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-started,log=f'{n}.log',sha256=sha(out/f'{n}.log')))
 (out/'manifest.json').write_text(json.dumps(dict(probe_sha256={str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()},runs=rows),indent=2)+'\n')
 print(n,r.returncode,flush=True)
 assert r.returncode==0
