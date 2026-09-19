#!/usr/bin/env python3
"""Freeze the separate refusal/control probe without changing the phase probe."""
import hashlib,json,os,shutil,subprocess,sys,time,tomllib
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
phase=sys.argv[1]
assert phase in ['baseline','candidate']
def sha(f):return hashlib.sha256(f.read_bytes()).hexdigest()
base=json.loads((P/'baseline.json').read_text())
for key in ['constraints_sha256','build_inputs_sha256']:
 for name,digest in base[key].items():assert sha(ROOT/name)==digest,name
source={str(f.relative_to(ROOT)):sha(f) for owner in ['litchi-pptx','litchi-ooxml-common','litchi-opc'] for f in (ROOT/'crates'/owner).rglob('*.rs')}
if phase=='baseline':assert source==base['source_sha256']
probe={str(f.relative_to(P)):sha(f) for f in (P/'oracle').rglob('*') if f.is_file()}
if phase=='candidate':assert probe==json.loads((P/'build-oracle-baseline.json').read_text())['probe_sha256']
manifest=P/'oracle/Cargo.toml';meta=tomllib.loads(manifest.read_text());name=meta.get('bin',[{'name':meta['package']['name']}])[0]['name']
target=ROOT.parent/'litchi-target-0700'
command=['cargo','build','--release','--locked','--manifest-path',str(manifest),'--target-dir',str(target),'-j','2']
start=time.monotonic()
with (P/f'build-oracle-{phase}.log').open('w') as log:
 result=subprocess.run(command,cwd=ROOT,env={**os.environ,'RUSTFLAGS':'-D warnings'},stdout=log,stderr=subprocess.STDOUT)
assert result.returncode==0
assert all(sha(ROOT/n)==h for n,h in source.items())
assert all(sha(P/n)==h for n,h in probe.items())
binary=ROOT.parent/'litchi-0700-bin'/(phase+'-oracle');shutil.copy2(target/'release'/name,binary)
(P/f'build-oracle-{phase}.json').write_text(json.dumps(dict(phase=phase,command=command,exit_code=result.returncode,seconds=time.monotonic()-start,source_sha256=source,probe_sha256=probe,binary=str(binary),binary_sha256=sha(binary)),indent=2)+'\n')
print(phase,'oracle built',flush=True)
