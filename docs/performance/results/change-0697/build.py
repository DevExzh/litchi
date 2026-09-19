#!/usr/bin/env python3
"""Build and bind the standalone attribution probe; never edit production."""
import hashlib,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
TARGET=ROOT.parent/'litchi-target-0697'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def bindings():
 paths=subprocess.check_output(['git','ls-files','crates'],cwd=ROOT,text=True).splitlines()
 return {f:sha(ROOT/f) for f in paths if (ROOT/f).is_file() and ((ROOT/f).suffix=='.rs' or (ROOT/f).name=='Cargo.toml')}
assert not subprocess.check_output(['git','status','--porcelain','--','crates','Cargo.toml','.cargo'],cwd=ROOT)
source=bindings()
probe={str(f.relative_to(P)):sha(f) for f in (P/'probe').rglob('*') if f.is_file()}
head=subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip()
command=['cargo','build','--release','--locked','--offline','--manifest-path',str(P/'probe/Cargo.toml'),'--target-dir',str(TARGET),'-j','2']
started=time.monotonic()
with (P/'build.log').open('w') as out:r=subprocess.run(command,cwd=ROOT,env={**os.environ,'RUSTFLAGS':'-D warnings'},stdout=out,stderr=subprocess.STDOUT)
record=dict(command=command,head=head,source_sha256=source,probe_sha256=probe,root_lock_sha256=sha(ROOT/'Cargo.lock'),rustflags='-D warnings',exit_code=r.returncode,seconds=time.monotonic()-started,log_sha256=sha(P/'build.log'))
assert bindings()==source
assert {str(f.relative_to(P)):sha(f) for f in (P/'probe').rglob('*') if f.is_file()}==probe
if r.returncode==0:
    binary=TARGET/'release/mce-attribution-0697'
    record.update(binary=str(binary),binary_sha256=sha(binary))
(P/'build.json').write_text(json.dumps(record,indent=2)+'\n')
assert r.returncode==0,(P/'build.log').read_text()
print('Probe built; source and probe hashes unchanged.',flush=True)
