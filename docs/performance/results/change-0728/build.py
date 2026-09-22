#!/usr/bin/env python3
"""Build only packet probes; bind complete workspace source and preserve failures."""
import hashlib,json,os,shutil,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];TARGET=ROOT.parent/'litchi-target-0728';BIN=ROOT.parent/'litchi-0728-bin'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def census():
 paths=[p for p in (ROOT/'crates').rglob('*') if p.is_file() and p.suffix in ('.rs','.toml')]
 paths.extend([ROOT/'Cargo.toml',ROOT/'Cargo.lock']);return {str(p.relative_to(ROOT)):sha(p) for p in sorted(paths)}
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
source=census();attempt=0
while (P/f'build-{attempt}.json').exists():attempt+=1
snapshot=P/f'build-{attempt}-probe';shutil.copytree(P/'probe',snapshot)
cmd=['cargo','build','--manifest-path',str(P/'probe/Cargo.toml'),'--release','--offline','--bins'];cmd.extend(['--locked'] if (P/'probe/Cargo.lock').exists() else []);start=time.monotonic()
with (P/f'build-{attempt}.log').open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2'),stdout=f,stderr=subprocess.STDOUT)
assert source==census();probe={str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()};rows=[]
if r.returncode==0:
 BIN.mkdir(exist_ok=True)
 for name in ['ole_format_save_probe','ole_format_save_probe_alloc']:
  dest=BIN/name;shutil.copy2(TARGET/'release'/name,dest);rows.append(dict(path=str(dest),bytes=dest.stat().st_size,sha256=sha(dest)))
write(P/f'build-{attempt}.json',dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,source_sha256=source,probe_sha256=probe,binaries=rows));assert r.returncode==0
write(P/'builds.json',dict(receipt=f'build-{attempt}.json',binaries=rows,source_sha256=source,probe_sha256=probe));print('PASS release probes; source unchanged')
