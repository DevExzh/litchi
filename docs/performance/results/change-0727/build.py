#!/usr/bin/env python3
"""Build exact baseline and archived candidate serially, then restore baseline."""
import hashlib,json,os,shutil,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];REL='crates/litchi-xls/src/workbook/source.rs'
TARGET=ROOT.parent/'litchi-target-0727';BIN=ROOT.parent/'litchi-0727-bin';BINARY='xls-index-retry-probe-0686'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def census():return {str(p.relative_to(ROOT)):sha(p) for c in ('litchi-cfb','litchi-xls') for p in sorted((ROOT/'crates'/c).rglob('*')) if p.is_file() and p.suffix in ('.rs','.toml')}
plan=json.loads((P/'plan.json').read_text());original=(ROOT/REL).read_bytes();assert original==subprocess.check_output(['git','show',plan['baseline_revision']+':'+REL],cwd=ROOT)
assert not (P/'builds.json').exists();baseline=census();rows=[]
try:
 for phase in ['baseline','candidate']:
  data=original if phase=='baseline' else (P.parent/'change-0726/candidate-source'/REL).read_bytes()
  (ROOT/REL).write_bytes(data);archive=P/'sources'/phase/REL;archive.parent.mkdir(parents=True,exist_ok=True);archive.write_bytes(data)
  source=census();write(P/f'source-{phase}.json',source)
  cmd=['cargo','build','--manifest-path','docs/performance/results/change-0686/probe/Cargo.toml','--release','--locked','--offline'];t=time.monotonic();log=P/f'{phase}.build.log'
  with log.open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2'),stdout=f,stderr=subprocess.STDOUT)
  assert source==census();row=dict(phase=phase,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-t,source_manifest=f'source-{phase}.json',log=log.name)
  if r.returncode==0:
   dest=BIN/phase/BINARY;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(TARGET/'release'/BINARY,dest);row.update(binary=str(dest),binary_sha256=sha(dest),bytes=dest.stat().st_size)
  rows.append(row);write(P/'builds.json',rows);assert r.returncode==0;print(phase,'built',flush=True)
finally:
 (ROOT/REL).write_bytes(original);assert census()==baseline;write(P/'build-restoration.json',dict(exact=True,source_sha256=baseline))
