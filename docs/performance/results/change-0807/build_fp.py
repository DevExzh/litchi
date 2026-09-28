"""Separate frame-pointer diagnostic build; ordinary binaries are retained."""
import os,subprocess,shutil,time
import custody as c
out=c.P/'build-fp';assert not out.exists();out.mkdir()
source=c.source();assert source==c.read(c.P/'build/source.json')
assert not os.environ.get('RUSTFLAGS'),'retain and compose existing flags before proceeding'
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','RUSTFLAGS':'-C force-frame-pointers=yes'}
cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(c.P/'probe-src/Cargo.toml'),'--features','capture-profile']
start=time.time();log=out/'build.log'
with log.open('w') as f:r=subprocess.run(cmd,env=env,stdout=f,stderr=subprocess.STDOUT)
receipt={'command':cmd,'started':start,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log),'source':c.artifact(c.P/'build/source.json'),'plan':c.artifact(c.P/'perf-fp-plan.json'),'environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS']}}
c.write(out/'receipt.json',receipt)
assert r.returncode==0 and c.source()==source
dest=c.TARGET/'profile-fp';assert not dest.exists();shutil.copy2(c.TARGET/'release/pptx-capture-probe',dest)
receipt['binary']=c.artifact(dest);c.write(out/'receipt.json',receipt)
print('Frame-pointer diagnostic built',flush=True)
