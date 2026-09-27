"""Root-only separate count/layout diagnostic; supplemental, no adoption change."""
import os,shutil,subprocess,sys,time
import custody as c
leg=sys.argv[1];assert leg in ['before','after']
out=c.P/f'attribute-{leg}';assert not out.exists();out.mkdir()
source=c.source();c.write(out/'source.json',source)
probe=c.P/'attribute-probe-src';manifest=probe/'Cargo.toml'
manifest.write_text((probe/'Cargo.toml.template').read_text().replace('@SRC@',str(c.ROOT)))
frozen={str(f.relative_to(c.P)):c.sha(f) for f in probe.rglob('*') if f.is_file() and f.name not in ['Cargo.lock','Cargo.toml']}
frozen['attribute_diagnostic.py']=c.sha(c.P/'attribute_diagnostic.py')
c.write(out/'frozen-inputs.json',frozen)
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
command=['cargo','build','--release','--offline','--manifest-path',str(manifest)]
if (probe/'Cargo.lock').exists():command+=['--locked']
started=time.time();log=out/'build.log'
with log.open('w') as f:r=subprocess.run(command,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
build={'command':command,'started':started,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)}
c.write(out/'build.json',build);assert r.returncode==0
binary=c.TARGET/f'attribute-{leg}';assert not binary.exists();shutil.copy2(c.TARGET/'release/attribute-allocation-probe-0794',binary)
identity=c.artifact(binary);report=out/'report.json';log=out/'run.log'
command=['taskset','-c','12',str(binary),str(report)];started=time.time()
with log.open('w') as f:r=subprocess.run(command,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
row={'command':command,'started':started,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log),'binary':identity,'source':c.artifact(out/'source.json'),'lock':c.artifact(probe/'Cargo.lock')}
if report.exists():row['report']=c.artifact(report)
c.write(out/'receipt.json',row);assert r.returncode==0
assert c.source()==source and c.artifact(binary)==identity
binary.unlink();c.write(out/'cleanup.json',{'removed_binary':identity,'binary_removed':not binary.exists()})
print(leg,'attribute diagnostic PASS:42samples, no timing measurement',flush=True)
