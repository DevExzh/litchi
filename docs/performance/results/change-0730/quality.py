#!/usr/bin/env python3
"""Affected owner checks, strictly serial with every attempt retained."""
import hashlib,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
i=0
while (P/f'quality-{i}').exists():i+=1
out=P/f'quality-{i}';out.mkdir()
pkgs=['-p','litchi-doc','-p','litchi-ole-common'];common=['--release','--offline','--locked']
manifest=str(P/'probe/Cargo.toml')
commands=[['cargo','fmt',*pkgs,'--','--check'],['cargo','test',*pkgs,*common,'--all-features','--all-targets'],['cargo','clippy',*pkgs,*common,'--all-features','--all-targets','--','-D','warnings'],['cargo','test',*pkgs,*common,'--all-features','--doc'],['cargo','doc',*pkgs,*common,'--all-features','--no-deps'],['cargo','fmt','--manifest-path',manifest,'--','--check'],['cargo','test','--manifest-path',manifest,*common,'--lib'],['cargo','clippy','--manifest-path',manifest,*common,'--all-targets','--','-D','warnings'],['cargo','doc','--manifest-path',manifest,*common,'--no-deps']]
retention_manifest=str(P/'retention-probe/Cargo.toml')
commands.extend([['cargo','fmt','--manifest-path',retention_manifest,'--','--check'],['cargo','clippy','--manifest-path',retention_manifest,*common,'--all-targets','--','-D','warnings'],['cargo','doc','--manifest-path',retention_manifest,*common,'--no-deps'],['python3','tools/check_crate_boundaries.py']])
env=dict(os.environ,CARGO_TARGET_DIR=str(ROOT.parent/'litchi-target-0730'),CARGO_BUILD_JOBS='2',RUSTDOCFLAGS='-D warnings');rows=[]
source={str(p.relative_to(ROOT)):sha(p) for p in (ROOT/'crates').rglob('*') if p.is_file() and p.suffix in ['.rs','.toml']}
for n,cmd in enumerate(commands):
 start=time.monotonic()
 with (out/f'{n}.log').open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append(dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,log=f'{n}.log',sha256=sha(out/f'{n}.log')))
 (out/'manifest.json').write_text(json.dumps(dict(source_sha256=source,runs=rows),indent=2)+'\n');print(n,r.returncode,flush=True)
 if r.returncode:raise SystemExit(r.returncode)
assert all(sha(ROOT/rel)==h for rel,h in source.items())
print('PASS all owner/probe checks; source unchanged')
