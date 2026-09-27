"""Root-only serial builds of an unchanged native tool and phase probe."""
import os,shutil,subprocess,sys,time
import custody as c
leg=sys.argv[1];assert leg in ('before','after')
out=c.P/f'build-{leg}';assert not out.exists();out.mkdir()
frozen={'production':c.source(),'tool':c.tool_source(),'probe':{p.name:c.sha(p) for p in (c.P/'probe-src').iterdir() if p.is_file()}}
history=c.read(c.P.parent/f'change-0787/build-{leg}/source.json')
assert frozen['production']['files']==history['production']['files']
assert frozen['tool']==history['tool']
for n,h in c.read(c.P/'architecture-inputs.json').items():assert c.sha(c.ROOT/n)==h,n
c.write(out/'source.json',frozen)
inputs=['plan.json','build.py','capture.py','quality.py','custody.py','architecture-inputs.json','host.json','origin.json','candidate-binding.json','quality.json','references.json','candidate.py']
c.write(out/'frozen-inputs.json',{n:c.sha(c.P/n) for n in inputs})
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
rows=[];binaries={};sections={}
for kind,manifest,features,exe in [('native',c.TOOL/'Cargo.toml',[],'litchi-perf-execution'),('memory',c.P/'probe-src/Cargo.toml',['--features','source-metrics'],'cached-part-memory')]:
 cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(manifest),*features]
 log=out/f'{kind}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'exit_code':r.returncode,'log':c.artifact(log),'started':start,'ended':time.time()});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 c.unchanged(frozen)
 assert frozen['probe']=={p.name:c.sha(p) for p in (c.P/'probe-src').iterdir() if p.is_file()}
 dest=c.TARGET/f'{leg}-{kind}';assert not dest.exists();shutil.copy2(c.TARGET/'release'/exe,dest);binaries[kind]=c.artifact(dest)
 sections[kind]={}
 for name,cmd in [('sections',['readelf','-W','-S',str(dest)]),('segments',['readelf','-W','-l',str(dest)]),('size',['size','-A',str(dest)])]:
  f=out/f'{kind}-{name}.txt';f.write_bytes(subprocess.check_output(cmd));sections[kind][name]=c.artifact(f)
 print(leg,kind,'built',flush=True)
receipt={'source':c.artifact(out/'source.json'),'frozen_inputs':c.artifact(out/'frozen-inputs.json'),'binaries':binaries,'sections':sections,'rows':rows,'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTUP_TOOLCHAIN']}}
c.write(out/'build.json',receipt)
(c.P/'builds').mkdir(exist_ok=True);c.write(c.P/'builds'/f'{leg}.json',receipt)
