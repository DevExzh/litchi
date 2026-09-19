#!/usr/bin/env python3
"""Freeze the oracle; baseline temporarily restores exact HEAD MCE sources."""
import hashlib,json,os,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
phase=sys.argv[1]
assert phase in ['baseline','candidate']
def sha(f):return hashlib.sha256(f.read_bytes()).hexdigest()
def sources():return {str(f.relative_to(ROOT)):sha(f) for owner in ['litchi-pptx','litchi-ooxml-common','litchi-opc'] for f in (ROOT/'crates'/owner).rglob('*.rs')}
def inputs():return {str(f.relative_to(P)):sha(f) for f in (P/'oracle').rglob('*') if f.is_file() and (f.name in ['Cargo.toml','Cargo.lock'] or f.suffix=='.rs')}
base=json.loads((P/'baseline.json').read_text())
paths=[ROOT/'crates/litchi-ooxml-common/src/mce'/name for name in ['codec.rs','tests.rs']]
saved={f:f.read_bytes() for f in paths}
receipt=dict(phase=phase,restore_required=phase=='baseline')
try:
 if phase=='baseline':
  archive=P/'oracle-baseline-restoration';archive.mkdir(exist_ok=True)
  for f,b in saved.items():
   assert not (archive/f.name).exists() or (archive/f.name).read_bytes()==b
   (archive/f.name).write_bytes(b)
   f.write_bytes(subprocess.check_output(['git','show',base['baseline_head']+':'+str(f.relative_to(ROOT))],cwd=ROOT))
  assert sources()==base['source_sha256']
 bound=sources();probe=inputs()
 command=['cargo','build','--release','--locked','--manifest-path',str(P/'oracle/Cargo.toml'),'--target-dir',str(ROOT.parent/'litchi-target-0694'),'-j','2']
 started=time.monotonic()
 with (P/('build-oracle-'+phase+'.log')).open('w') as log:
  result=subprocess.run(command,cwd=ROOT,env={**os.environ,'RUSTFLAGS':'-D warnings'},stdout=log,stderr=subprocess.STDOUT)
 receipt.update(command=command,exit_code=result.returncode,seconds=time.monotonic()-started,source_sha256=bound,probe_sha256=probe)
 assert result.returncode==0
 assert sources()==bound and inputs()==probe
 binary=ROOT.parent/'litchi-0694-bin'/(phase+'-oracle')
 shutil.copy2(ROOT.parent/'litchi-target-0694/release/mce-oracle-0694',binary)
 receipt.update(binary=str(binary),binary_sha256=sha(binary))
finally:
 if phase=='baseline':
  for f,b in saved.items():f.write_bytes(b)
 receipt['restored_exact']=all(f.read_bytes()==b for f,b in saved.items())
 receipt['restored_sha256']={str(f.relative_to(ROOT)):sha(f) for f in saved}
 (P/('build-oracle-'+phase+'.json')).write_text(json.dumps(receipt,indent=2)+'\n')
print(phase,'oracle frozen',flush=True)
