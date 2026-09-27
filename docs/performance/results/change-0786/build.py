"""Root-only serial build. Retains every attempt; never overwrites evidence."""
import os, shutil, subprocess, time
import custody as c
attempt=0
while (c.P/f'build-{attempt}').exists():attempt+=1
out=c.P/f'build-{attempt}';out.mkdir()
frozen={'production':c.source(),'tool':c.tool_source()};c.write(out/'source.json',frozen)
assert subprocess.run(['git','diff','--quiet',c.read(c.P/'origin.json')['base'],'--','crates','Cargo.toml','clippy.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=c.ROOT).returncode==0
c.write(out/'frozen-inputs.json',{n:c.sha(c.P/n) for n in ['plan.json','build.py','capture.py','quality.py','custody.py','architecture-inputs.json','host.json','cgroup-limits.json','origin.json']})
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
rows=[];binaries={}
for kind,features in [('native',[]),('observer',['--features','source-metrics'])]:
 cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(c.TOOL/'Cargo.toml'),*features]
 log=out/f'{kind}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'exit_code':r.returncode,'log':c.artifact(log),'started':start,'ended':time.time()});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 c.unchanged(frozen)
 dest=c.TARGET/f'attempt-{attempt}-{kind}';assert not dest.exists()
 shutil.copy2(c.TARGET/'release/litchi-perf-execution',dest);binaries[kind]=c.artifact(dest)
 print(kind,'built',flush=True)
result={'source':c.artifact(out/'source.json'),'frozen_inputs':c.artifact(out/'frozen-inputs.json'),'binaries':binaries,'rows':rows,'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTUP_TOOLCHAIN']}}
c.write(out/'build.json',result);assert not (c.P/'build.json').exists();c.write(c.P/'build.json',result)
