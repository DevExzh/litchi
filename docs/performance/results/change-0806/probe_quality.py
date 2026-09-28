"""Root-only probe format, test and Clippy gates with source custody."""
import os,subprocess,sys,time
import custody as c
leg=sys.argv[1];assert leg in ['before','after']
p=c.P;out=p/('probe-quality-'+leg);assert not out.exists();out.mkdir()
source=c.source();build=c.read(p/('build-'+leg)/'build.json');assert c.read(build['source']['path'])==source
probe={str(f.relative_to(p/'probe-src')):c.sha(f) for f in (p/'probe-src').rglob('*') if f.is_file()}
c.write(out/'inputs.json',{'source':source,'probe':probe,'driver':c.artifact(__file__)})
manifest=str(p/'probe-src/Cargo.toml')
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
commands=[['cargo','fmt','--manifest-path',manifest,'--','--check'],['cargo','test','--offline','--locked','--release','--manifest-path',manifest,'--all-features','--','--test-threads=1'],['cargo','clippy','--offline','--locked','--release','--manifest-path',manifest,'--all-features','--all-targets','--','-D','warnings']]
rows=[]
for i,cmd in enumerate(commands):
 log=out/f'{i}.log';started=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'started':started,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)});c.write(out/'receipts.json',rows)
 assert r.returncode==0,log
 assert c.source()==source
 assert all(c.sha(p/'probe-src'/n)==h for n,h in probe.items())
 print(leg,'probe gate',i+1,'PASS',flush=True)
c.write(out/'complete.json',{'inputs':c.artifact(out/'inputs.json'),'receipts':c.artifact(out/'receipts.json')})
