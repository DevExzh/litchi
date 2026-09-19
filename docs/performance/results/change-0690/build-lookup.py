#!/usr/bin/env python3
import hashlib,json,os,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1];assert phase in ['baseline','candidate'];probe=P/'lookup-probe';target=Path('/home/zhuhe/code/litchi-target-0690')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def sources():return {str(f.relative_to(ROOT)):sha(f) for owner in ['litchi-cfb','litchi-xls'] for f in sorted((ROOT/'crates'/owner).rglob('*.rs'))}
h={str(f.relative_to(ROOT)):sha(f) for f in [probe/'Cargo.toml',probe/'Cargo.lock',probe/'src/main.rs']};s=sources();cmd=['cargo','build','--manifest-path',str(probe.relative_to(ROOT)/'Cargo.toml'),'--release','--offline','--locked'];start=time.monotonic()
with (P/(phase+'-lookup-build.log')).open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,CARGO_TARGET_DIR=str(target),CARGO_BUILD_JOBS='2',RUSTFLAGS='-D warnings'),stdout=f,stderr=subprocess.STDOUT)
assert sources()==s and all(sha(ROOT/n)==v for n,v in h.items())
row=dict(command=cmd,rustflags='-D warnings',exit_code=r.returncode,seconds=time.monotonic()-start,probe_sha256=h,source_sha256=s);(P/(phase+'-lookup-build.json')).write_text(json.dumps(row,indent=2)+'\n');print('build',r.returncode,flush=True);assert r.returncode==0
b=target/'release/cfb-lookup-probe-0690';dest=Path('/home/zhuhe/code/litchi-target-0690-'+('before' if phase=='baseline' else 'after'))/'release'/b.name;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(b,dest);row['binary_sha256']=sha(dest);(P/(phase+'-lookup-build.json')).write_text(json.dumps(row,indent=2)+'\n')
cmd=['taskset','-c','12',str(dest),'--case','all','--format','json','--repetitions','10000'];r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);(P/(phase+'-lookup-smoke.json')).write_text(r.stdout);(P/(phase+'-lookup-smoke.stderr')).write_text(r.stderr);print('smoke',r.returncode);assert r.returncode==0,r.stderr
