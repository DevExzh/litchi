"""Root-only serial paired builds, original standalone tool frozen throughout."""
import os,shutil,subprocess,sys,time
import custody as c
leg=sys.argv[1];assert leg in ('before','after')
if leg=='before':assert subprocess.run(['git','diff','--quiet',c.read(c.P/'origin.json')['base'],'--','crates','Cargo.toml','clippy.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=c.ROOT).returncode==0
out=c.P/f'build-{leg}';assert not out.exists();out.mkdir()
frozen={'production':c.source(),'tool':c.tool_source()};c.write(out/'source.json',frozen)
c.write(out/'frozen-inputs.json',{n:c.sha(c.P/n) for n in ['plan.json','adoption-policy.json','build.py','capture.py','quality.py','custody.py','architecture-inputs.json','host.json','origin.json','profile-plan.json']})
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
rows=[];binaries={}
for kind,features in [('native',[]),('observer',['--features','source-metrics'])]:
 cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(c.TOOL/'Cargo.toml'),*features]
 log=out/f'{kind}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'exit_code':r.returncode,'log':c.artifact(log),'started':start,'ended':time.time()});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 c.unchanged(frozen)
 dest=c.TARGET/f'{leg}-{kind}';assert not dest.exists();shutil.copy2(c.TARGET/'release/litchi-perf-execution',dest);binaries[kind]=c.artifact(dest)
 print(leg,kind,'built',flush=True)
c.write(out/'build.json',{'source':c.artifact(out/'source.json'),'frozen_inputs':c.artifact(out/'frozen-inputs.json'),'binaries':binaries,'rows':rows,'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTUP_TOOLCHAIN']}})
