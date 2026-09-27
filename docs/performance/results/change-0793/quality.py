"""Root-only probe correctness gates; production bytes are unchanged."""
import os,subprocess,time
import custody as c
out=c.P/'quality';assert not out.exists();out.mkdir();source=c.source();assert source==c.read(c.P/'build/source.json')
base=['cargo','test','--release','--offline','--locked','--manifest-path',str(c.P/'probe-src/Cargo.toml')]
commands=[['cargo','fmt','--manifest-path',str(c.P/'probe-src/Cargo.toml'),'--','--check'],base+['--','--test-threads=1'],base+['--all-features','--','--test-threads=1','--skip','allocation_metrics::tests::']]
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'};rows=[]
for i,cmd in enumerate(commands):
 log=out/f'{i:02}.log';start=time.time()
 with log.open('w') as stream:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=stream,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'log':c.artifact(log)});c.write(out/'receipts.json',rows)
 assert r.returncode==0,log
 assert c.source()==source
 print('probe gate',i,'PASS',flush=True)
c.write(out/'complete.json',{'gates':len(rows),'receipts':c.artifact(out/'receipts.json'),'source':c.artifact(c.P/'build/source.json'),'scope':'Fresh evidence-probe tests; no production changes or fresh production-suite claim','environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL']}})
