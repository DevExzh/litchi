"""Root-only serial OPC/XLSX quality gates; each invocation retains its attempt."""
import os
import subprocess
import time
import custody as c

attempt=0
while (c.P/f'quality-{attempt}').exists():attempt+=1
out=c.P/f'quality-{attempt}';out.mkdir();frozen=c.source();c.write(out/'source.json',frozen)
packages=['-p','litchi-opc','-p','litchi-xlsx']
commands=[
 ['cargo','fmt',*packages,'--','--check'],
 ['cargo','check','--offline','--locked',*packages,'--all-features','--all-targets'],
 ['cargo','test','--offline','--locked',*packages,'--all-features','--','--test-threads=2'],
 ['cargo','clippy','--offline','--locked',*packages,'--all-features','--lib','--','-D','warnings'],
 ['cargo','doc','--offline','--locked',*packages,'--all-features','--no-deps'],
 ['python3','-B','tools/check_crate_boundaries.py'],
]
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET/'quality'),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','CARGO_PROFILE_DEV_DEBUG':'0','RUSTDOCFLAGS':'-D warnings','PYTHONDONTWRITEBYTECODE':'1'}
rows=[]
for i,cmd in enumerate(commands):
 log=out/f'{i:02}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'exit_code':r.returncode,'log':c.artifact(log),'started':start,'ended':time.time()});c.write(out/'checks.json',rows)
 print('gate',i+1,'exit',r.returncode,flush=True)
 assert r.returncode==0,log
 assert c.source()==frozen
c.write(c.P/'quality.json',{'source':c.artifact(out/'source.json'),'rows':rows,'environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_DEV_DEBUG','RUSTDOCFLAGS']}})
