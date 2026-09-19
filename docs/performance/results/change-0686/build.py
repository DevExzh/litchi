#!/usr/bin/env python3
"""Build the frozen probes serially, recording source/binary identities."""
import hashlib,json,os,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1];assert phase in ['baseline','candidate']
root=Path('/home/zhuhe/code/litchi-0686-before') if phase=='baseline' else ROOT
target=Path('/home/zhuhe/code/litchi-target-0686-'+('before' if phase=='baseline' else 'after'))
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def sources():return {str(f.relative_to(root)):sha(f) for owner in ['litchi-cfb','litchi-xls'] for f in sorted((root/'crates'/owner).rglob('*.rs'))}
rows=[]
for folder,binary in [('change-0686/probe','xls-index-retry-probe-0686'),('change-0686/allocation-probe','xls0686-alloc'),('change-0684/budget-probe','xls-index-budget-probe-0684'),('change-0684/probe','xls-index-probe-0684')]:
 hashes=sources();cmd=['cargo','build','--manifest-path','docs/performance/results/'+folder+'/Cargo.toml','--release','--locked','--offline'];start=time.monotonic()
 with (P/(phase+'-final-build-'+binary+'.log')).open('w') as log:r=subprocess.run(cmd,cwd=root,env=dict(os.environ,CARGO_BUILD_JOBS='2',CARGO_TARGET_DIR=str(target)),stdout=log,stderr=subprocess.STDOUT)
 assert r.returncode==0,(binary,r.returncode);assert hashes==sources(),'source changed during build'
 rows.append(dict(command=cmd,cwd=str(root),target=str(target),exit_code=r.returncode,seconds=time.monotonic()-start,source_sha256=hashes,binary=binary,binary_sha256=sha(target/'release'/binary)))
 (P/(phase+'-builds.json')).write_text(json.dumps(rows,indent=2)+'\n');print(phase,binary,flush=True)
