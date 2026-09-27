"""Root-only source-bound control/instrumented builds; no adoption."""
import os,subprocess,sys,time,shutil
import custody as c
p=c.P;leg=sys.argv[1];assert leg in ['before','after'];out=p/('build-'+leg);assert not out.exists();out.mkdir()
source=c.source();base=c.read(p/'source.json');changed={n:h for n,h in source['files'].items() if h!=base['files'][n]}
assert source['revision']==base['revision']
assert set(changed)==(set() if leg=='before' else {'crates/litchi-opc/src/xml_attributes.rs'})
if leg=='after':assert c.sha(c.ROOT/'crates/litchi-opc/src/xml_attributes.rs')==c.sha(p/'instrumentation/after/xml_attributes.rs')
c.write(out/'source.json',source)
frozen={n:c.sha(p/n) for n in ['lock-generation.json','lock-generation.log','probe-inheritance.json','quality.py','quality/complete.json','plan.json','build.py','capture.py','custody.py','origin.json','inheritance.json','architecture-inputs.json','host.json']}
probe={str(f.relative_to(p/'probe-src')):c.sha(f) for f in (p/'probe-src').rglob('*') if f.is_file()}
instrumentation={str(f.relative_to(p/'instrumentation')):c.sha(f) for f in (p/'instrumentation').rglob('*') if f.is_file()}
c.write(out/'inputs.json',{'frozen':frozen,'probe':probe,'instrumentation':instrumentation})
assert not os.environ.get('RUSTFLAGS')
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
manifest=str(p/'probe-src/Cargo.toml');features=[] if leg=='before' else ['--features','attribute-census']
commands=[['rustfmt','--check','--edition','2024','--config','skip_children=true',str(p/'probe-src/src/main.rs')]]
for verb in ['build','check','clippy']:
 commands.append(['cargo',verb,'--offline','--locked','--release','--manifest-path',manifest,'--bin','namespace-uri-probe',*features]+(['--','-D','warnings'] if verb=='clippy' else []))
rows=[]
for i,cmd in enumerate(commands):
 log=out/f'{i}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'started':start,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 assert c.source()==source and all(c.sha(p/'probe-src'/n)==h for n,h in probe.items())
 print(leg,'gate',i+1,'PASS',flush=True)
binary=c.TARGET/(leg+'-probe');assert not binary.exists();shutil.copy2(c.TARGET/'release/namespace-uri-probe',binary)
c.write(out/'build.json',{'binary':c.artifact(binary),'source':c.artifact(out/'source.json'),'inputs':c.artifact(out/'inputs.json'),'rows':rows,'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS']}})
