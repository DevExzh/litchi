"""Format the exact inherited probe; no production test-suite claim."""
import subprocess,time
import custody as c
out=c.P/'quality.json'
assert not out.exists()
source=c.source();assert source==c.read(c.P/'build/source.json')
probe=c.read(c.P/'build/probe.json')
assert all(c.sha(c.P/n)==h for n,h in probe.items())
cmd=['cargo','fmt','--manifest-path',str(c.P/'probe-src/Cargo.toml'),'--','--check']
log=c.P/'format.log';start=time.time()
with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
c.write(out,{'format':{'command':cmd,'started':start,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)},'source':c.artifact(c.P/'build/source.json'),'probe':probe,'scope':'Exact inherited Rust probe; fresh formatting and release builds. No production source change or fresh production test-suite claim.'})
assert r.returncode==0
assert c.source()==source and all(c.sha(c.P/n)==h for n,h in probe.items())
print('0807 probe formatting PASS',flush=True)
