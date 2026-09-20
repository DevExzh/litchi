#!/usr/bin/env python3
"""Build immutable probes in one serial lane; copy executables per revision."""
import hashlib,json,os,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1];assert phase in ['baseline','candidate']
target=Path('/home/zhuhe/code/litchi-target-0725');dest=Path('/home/zhuhe/code/litchi-0725-bin')/phase;dest.mkdir(parents=True,exist_ok=True)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def sources():return {str(f.relative_to(ROOT)):sha(f) for owner in ['litchi-cfb','litchi-xls'] for f in sorted((ROOT/'crates'/owner).rglob('*.rs'))}
rows=[]
for folder,binary in [('change-0684/repeat-probe','xls0684-repeat'),('change-0686/probe','xls-index-retry-probe-0686'),('change-0686/allocation-probe','xls0686-alloc'),('change-0684/budget-probe','xls-index-budget-probe-0684'),('change-0684/probe','xls-index-probe-0684')]:
 hashes=sources();cmd=['cargo','build','--manifest-path','docs/performance/results/'+folder+'/Cargo.toml','--release','--locked','--offline'];start=time.monotonic()
 with (P/(phase+'-build-'+binary+'.log')).open('w') as log:r=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,CARGO_BUILD_JOBS='2',CARGO_TARGET_DIR=str(target)),stdout=log,stderr=subprocess.STDOUT)
 assert r.returncode==0,(binary,r.returncode);assert hashes==sources(),'source changed during build'
 shutil.copy2(target/'release'/binary,dest/binary)
 rows.append(dict(command=cmd,cwd=str(ROOT),target=str(target),exit_code=r.returncode,seconds=time.monotonic()-start,source_sha256=hashes,binary=binary,binary_sha256=sha(dest/binary)))
 (P/(phase+'-builds.json')).write_text(json.dumps(rows,indent=2)+'\n');print(phase,binary,flush=True)
