"""Diagnostic-only frame-pointer build after ordinary DWARF qualification failed."""
import json,os,subprocess,time
from contract import P,ROOT,guard,sha
if __name__=='__main__':
    guard(); target=ROOT.parent/'litchi-target-0740-fp';assert not target.exists()
    assert not os.environ.get('RUSTFLAGS') and not os.environ.get('CARGO_ENCODED_RUSTFLAGS')
    env=os.environ|{'CARGO_TARGET_DIR':str(target),'CARGO_BUILD_JOBS':'2','RUSTFLAGS':'-C force-frame-pointers=yes -C force-unwind-tables=yes'}
    cmd=['cargo','build','--release','--offline','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--bin','litchi-perf-baseline']
    start=time.time()
    with (P/'build-fp.log').open('w') as log:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=log,stderr=subprocess.STDOUT)
    row={'command':cmd,'started':start,'ended':time.time(),'exit':r.returncode,'rustflags':env['RUSTFLAGS'],'target':str(target),'log_sha256':sha(P/'build-fp.log')}
    if r.returncode==0:
        b=target/'release'/'litchi-perf-baseline';row|={'binary':str(b),'binary_sha256':sha(b),'binary_bytes':b.stat().st_size}
    (P/'build-fp.json').write_text(json.dumps(row,indent=2)+'\n');assert r.returncode==0;guard();print('PASS diagnostic frame-pointer build')
