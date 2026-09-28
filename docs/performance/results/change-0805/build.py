"""Root-only quality gates and release build of the direct helper probe."""
import os,subprocess,shutil,time
import custody as c
p=c.P;out=p/'build';assert not out.exists();out.mkdir();source=c.source();assert source==c.read(p/'source.json')
probe={str(f.relative_to(p/'probe-src')):c.sha(f) for f in (p/'probe-src').rglob('*') if f.is_file()}
assert c.sha(p/'probe-src/src/baseline.rs')==c.read(p/'inheritance.json')['baseline_helper']['sha256']
assert c.sha(p/'probe-src/src/candidate.rs')==c.read(p/'inheritance.json')['candidate_helper']['sha256']
c.write(out/'inputs.json',{'probe':probe,'frozen':{n:c.sha(p/n) for n in ['candidate.patch','decision.py','quality.py','quality/archive-inputs.json','quality/complete.json','fixtures.json','fixture_check.py','lock-generation.json','lock-generation.log','plan.json','build.py','capture.py','custody.py','origin.json','inheritance.json','architecture-inputs.json','host.json']}})
assert not os.environ.get('RUSTFLAGS')
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
manifest=str(p/'probe-src/Cargo.toml')
commands=[['rustfmt','--check','--edition','2024','--config','skip_children=true',str(p/'probe-src/src/main.rs')],['cargo','build','--offline','--locked','--release','--manifest-path',manifest,'--bin','attribute-boundary-probe'],['cargo','check','--offline','--locked','--release','--manifest-path',manifest,'--bin','attribute-boundary-probe'],['cargo','clippy','--offline','--locked','--release','--manifest-path',manifest,'--bin','attribute-boundary-probe','--','-D','warnings']]
rows=[]
for i,cmd in enumerate(commands):
 log=out/f'{i}.log';started=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'started':started,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 assert probe=={n:c.sha(p/'probe-src'/n) for n in probe}
 print('gate',i+1,'PASS',flush=True)
binary=c.TARGET/'attribute-boundary-probe';assert not binary.exists();shutil.copy2(c.TARGET/'release/attribute-boundary-probe',binary)
for option,name in [('--list-cases','cases.json'),('--self-check','self-check.json')]:
 cmd=[str(binary),option];log=out/(name+'.log');dest=p/name;started=time.time()
 with dest.open('w') as f,log.open('w') as err:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=err)
 rows.append({'command':cmd,'started':started,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log),'output':c.artifact(dest)});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
from fixture_check import check
check()
assert c.source()==source
symbol_command=['nm','-C',str(binary)]
symbols=subprocess.run(symbol_command,cwd=c.ROOT,text=True,capture_output=True)
(out/'symbols.txt').write_text(symbols.stdout)
(out/'symbols.stderr').write_text(symbols.stderr)
c.write(out/'symbols.json',{'command':symbol_command,'exit_code':symbols.returncode,'stdout':c.artifact(out/'symbols.txt'),'stderr':c.artifact(out/'symbols.stderr'),'binary':c.artifact(binary)})
assert symbols.returncode==0
for leg in ['before','after']:
 for mode in ['construct','consume']:
  owner='attribute_boundary_probe::'+leg+'_'+mode
  assert sum(line.endswith(' '+owner) for line in symbols.stdout.splitlines())==1,owner
c.write(out/'build.json',{'binary':c.artifact(binary),'inputs':c.artifact(out/'inputs.json'),'lock':c.artifact(p/'probe-src/Cargo.lock'),'rows':rows,'environment':{n:env.get(n) for n in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS']}})
print('Probe quality and build PASS',flush=True)
