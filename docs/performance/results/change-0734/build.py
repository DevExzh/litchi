#!/usr/bin/env python3
"""Serial exact-source before/after probe builds with preserved attempts."""
import hashlib,json,os,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
TARGET=ROOT.parent/'litchi-target-0734';BIN=ROOT.parent/'litchi-0734-bin'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def source():return {str(p.relative_to(ROOT)):sha(p) for p in sorted(list((ROOT/'crates').rglob('*.rs'))+list((ROOT/'crates').rglob('Cargo.toml'))+[ROOT/'Cargo.toml',ROOT/'Cargo.lock'])}
variant=sys.argv[1];assert variant in ['baseline','candidate'];before=source();probe={str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()};i=0
while (P/f'build-{i}').exists():i+=1
out=P/f'build-{i}';out.mkdir();shutil.copytree(P/'probe',out/'probe');f=read(P/'base.json')['owned_file'];d=out/'source'/f;d.parent.mkdir(parents=True);shutil.copy2(ROOT/f,d)
assert all(sha(ROOT/f)==h for f,h in read(P/'constraints.json').items())
if variant=='baseline':assert before[f]==read(P/'base.json')['before_sha256']
env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2',RUSTDOCFLAGS='-D warnings');m=str(P/'probe/Cargo.toml');common=['--manifest-path',m,'--release','--offline','--locked'];commands=[['cargo','fmt','--manifest-path',m,'--','--check'],['cargo','test',*common,'--lib'],['cargo','clippy',*common,'--all-targets','--','-D','warnings'],['cargo','doc',*common,'--no-deps'],['cargo','build',*common,'--bins']];rows=[]
for n,cmd in enumerate(commands):
 start=time.monotonic()
 with (out/f'{n}.log').open('wb') as handle:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=handle,stderr=subprocess.STDOUT)
 rows.append(dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,output=f'build-{i}/{n}.log',sha256=sha(out/f'{n}.log')));write(out/'manifest.json',dict(variant=variant,source=before,probe=probe,runs=rows));print(variant,n,r.returncode,flush=True);assert r.returncode==0
assert source()==before and probe=={str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()}
BIN.mkdir(exist_ok=True);bins={}
for name in ['ole_format_save_probe','ole_format_save_probe_alloc']:
 dest=BIN/(variant+'-'+name);assert not dest.exists();shutil.copy2(TARGET/'release'/name,dest);bins['allocation' if name.endswith('_alloc') else 'native']=dict(path=str(dest),bytes=dest.stat().st_size,sha256=sha(dest))
write(P/(variant+'-build.json'),dict(source=before,probe=probe,quality=f'build-{i}/manifest.json',quality_sha256=sha(out/'manifest.json'),binaries=bins));print('PASS five probe gates and exact-source '+variant+' binaries')
