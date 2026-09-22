#!/usr/bin/env python3
"""Serial affected-owner/probe qualification; every failed attempt is retained."""
import hashlib,json,os,subprocess,time,shutil
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];TARGET=ROOT.parent/'litchi-target-0734';BIN=ROOT.parent/'litchi-0734-bin'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def source():return {str(p.relative_to(ROOT)):sha(p) for p in sorted(list((ROOT/'crates').rglob('*.rs'))+list((ROOT/'crates').rglob('Cargo.toml'))+[ROOT/'Cargo.toml',ROOT/'Cargo.lock'])}
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
i=0
while (P/f'quality-{i}').exists():i+=1
out=P/f'quality-{i}';out.mkdir()
for rel in ['crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs']:
 dest=out/'source'/rel;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(ROOT/rel,dest)
shutil.copytree(P/'probe',out/'probe')
before=source();probe={str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()};env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2',RUSTDOCFLAGS='-D warnings');m=str(P/'probe/Cargo.toml');owner=['-p','litchi-ppt','--release','--offline','--locked'];feature=['--features','performance-diagnostics'];common=['--manifest-path',m,'--release','--offline','--locked'];commands=[['cargo','fmt','-p','litchi-ppt','--','--check'],['cargo','check',*owner,'--no-default-features'],['cargo','test',*owner,*feature,'--all-targets'],['cargo','clippy',*owner,*feature,'--all-targets','--','-D','warnings'],['cargo','test',*owner,*feature,'--doc'],['cargo','doc',*owner,*feature,'--no-deps'],['python3','tools/check_crate_boundaries.py']];runs=[]
for n,cmd in enumerate(commands):
 start=time.monotonic()
 with (out/f'{n}.log').open('wb') as f:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 runs.append(dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,output=f'quality-{i}/{n}.log',sha256=sha(out/f'{n}.log')));write(out/'manifest.json',dict(source=before,probe=probe,runs=runs));print(n,r.returncode,flush=True)
 assert r.returncode==0
assert source()==before;assert probe=={str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()}
write(P/'quality.json',dict(source=before,probe=probe,runs=runs));print('PASS seven serial owner gates')
