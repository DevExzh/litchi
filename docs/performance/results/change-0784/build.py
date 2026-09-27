"""Serial evidence-only builds; frozen source and executable custody."""
import os, shutil, subprocess, time
import custody as c
assert not (c.P/'build').exists()
assert not c.TARGET.exists()
out=c.P/'build';out.mkdir()
manifest=c.P/'probe-src/Cargo.toml'
manifest.write_text((manifest.with_suffix('.toml.template')).read_text().replace('@SRC@',str(c.ROOT)))
source=c.source();c.write(out/'source.json',source)
c.write(out/'frozen-inputs.json',{n:c.sha(c.P/n) for n in ['plan.json','build.py','capture.py','origin.json','architecture-inputs.json']})
probe={str(p.relative_to(c.P)):c.sha(p) for p in (c.P/'probe-src').rglob('*') if p.is_file()}
c.write(out/'probe.json',probe)
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
rows=[];binaries={}
for leg,features in [('control',[]),('profile',['--features','capture-profile'])]:
 cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(manifest),*features]
 log=out/f'{leg}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'log':c.artifact(log)})
 c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 assert c.source()==source
 assert {n:c.sha(c.P/n) for n in probe}==probe
 dest=c.TARGET/leg;shutil.copy2(c.TARGET/'release/pptx-capture-probe',dest)
 binaries[leg]=c.artifact(dest)
 print(leg,'built',flush=True)
c.write(out/'build.json',{'binaries':binaries,'probe':probe,'source':c.artifact(out/'source.json'),'commands':rows,'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTUP_TOOLCHAIN']}})
