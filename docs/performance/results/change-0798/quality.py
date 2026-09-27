"""Root-only isolated instrumented-helper tests; source copies remain exact."""
import os,subprocess,time
from pathlib import Path
import custody as c
p=c.P;out=p/'quality';assert not out.exists();out.mkdir();project=p/'hook-test-src';assert not project.exists();src=project/'src';src.mkdir(parents=True)
(src/'lib.rs').write_text('#![forbid(unsafe_code)]\n#![allow(dead_code)]\nmod xml_attributes;\n')
(src/'xml_attributes.rs').write_bytes((p/'instrumentation/after/xml_attributes.rs').read_bytes())
(src/'xml_attributes').mkdir();(src/'xml_attributes/tests.rs').write_bytes((p/'instrumentation/after/canonical_tests.rs').read_bytes())
(src/'census_tests_0798.rs').write_bytes((p/'instrumentation/after/census_tests_0798.rs').read_bytes())
(project/'Cargo.toml').write_text('[package]\nname="attribute-census-hook-tests"\nversion="0.0.0"\nedition="2024"\npublish=false\n[dependencies]\nquick-xml="=0.41.0"\n[workspace]\n')
manifest=str(project/'Cargo.toml');env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET/'hook-tests'),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
commands=[['cargo','generate-lockfile','--offline','--manifest-path',manifest],['cargo','test','--offline','--locked','--manifest-path',manifest,'--','--test-threads=2','--skip','every_copy_of_the_module_is_the_canonical_code'],['cargo','clippy','--offline','--locked','--manifest-path',manifest,'--all-targets','--','-D','warnings']]
rows=[]
for i,cmd in enumerate(commands):
 log=out/f'{i}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'started':start,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)});c.write(out/'receipts.json',rows)
 assert r.returncode==0,log
 print('hook gate',i+1,'PASS',flush=True)
assert c.source()==c.read(p/'source.json')
c.write(out/'complete.json',{'rows':rows,'inputs':{str(f.relative_to(project)):c.sha(f) for f in project.rglob('*') if f.is_file()},'scope':'Exact OPC helper/parser plus diagnostic hook tests; excludes canonical-copy filesystem test because instrumentation is OPC-only'})
