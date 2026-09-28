"""Root-only exact-helper minimal workspace tests, both legs run serially."""
import os,subprocess,time
from pathlib import Path
import custody as c
p=c.P;owners=['litchi-opc','litchi-ole-common','litchi-sign','litchi-xldm','xml-minifier']
assert not (p/'quality').exists();(p/'quality').mkdir()
for owner in owners:assert (p/'candidate/before'/(owner+'-xml_attributes.rs')).read_bytes()==(p.parent/'change-0803/candidate/after'/(owner+'-xml_attributes.rs')).read_bytes()
assert (p/'candidate/before/litchi-opc-xml_attributes-tests.rs').read_bytes()==(p.parent/'change-0803/candidate/after/litchi-opc-xml_attributes-tests.rs').read_bytes()
c.write(p/'quality/archive-inputs.json',{str(f.relative_to(p/'candidate')):c.sha(f) for f in (p/'candidate').rglob('*') if f.is_file()})
rows=[];inputs={};env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET/'helper-tests'),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
for leg in ['before','after']:
 project=p/'test-src'/leg;project.mkdir(parents=True)
 (project/'Cargo.toml').write_text('[workspace]\nresolver="3"\nmembers=['+','.join('"crates/'+o+'"' for o in owners)+']\n')
 for owner in owners:
  crate=project/'crates'/owner;src=crate/'src';src.mkdir(parents=True)
  (crate/'Cargo.toml').write_text('[package]\nname="'+owner+'"\nversion="0.0.0"\nedition="2024"\npublish=false\n[dependencies]\nquick-xml="=0.41.0"\n')
  (src/'lib.rs').write_text('#![forbid(unsafe_code)]\n#![allow(dead_code)]\nmod xml_attributes;\n')
  (src/'xml_attributes.rs').write_bytes((p/'candidate'/leg/(owner+'-xml_attributes.rs')).read_bytes())
 tests=project/'crates/litchi-opc/src/xml_attributes/tests.rs';tests.parent.mkdir()
 tests.write_bytes((p/'candidate'/leg/'litchi-opc-xml_attributes-tests.rs').read_bytes())
 manifest=str(project/'Cargo.toml')
 commands=[['cargo','generate-lockfile','--offline','--manifest-path',manifest],['cargo','test','--offline','--locked','--manifest-path',manifest,'--workspace','--','--test-threads=2'],['cargo','clippy','--offline','--locked','--manifest-path',manifest,'--workspace','--all-targets','--','-D','warnings']]
 for i,cmd in enumerate(commands):
  log=p/'quality'/f'{leg}-{i}.log';start=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
  rows.append({'leg':leg,'command':cmd,'started':start,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)});c.write(p/'quality/receipts.json',rows)
  assert r.returncode==0,log
  print(leg,i,'PASS',flush=True)
 inputs[leg]={str(f.relative_to(project)):c.sha(f) for f in project.rglob('*') if f.is_file()}
 assert c.source()==c.read(p/'source.json')
 assert all(c.sha(p/'candidate'/n)==h for n,h in c.read(p/'quality/archive-inputs.json').items())
c.write(p/'quality/complete.json',{'rows':rows,'inputs':inputs,'scope':'Exact helper modules and shared tests in minimal mirror crates; not full production-crate checks'})
