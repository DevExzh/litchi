"""Root-only serial probe checks before baseline qualification."""
import os,subprocess,time
import custody as c
out=c.P/'probe-checks';assert not out.exists();out.mkdir()
commands=[['cargo','fmt','--manifest-path',str(c.P/'probe-src/Cargo.toml'),'--','--check'],['cargo','test','--offline','--locked','--release','--all-features','--manifest-path',str(c.P/'probe-src/Cargo.toml')]]
frozen=c.source();rows=[]
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
for index,command in enumerate(commands):
    log=out/f'{index}.log';start=time.time()
    with log.open('w') as stream:r=subprocess.run(command,cwd=c.ROOT,env=env,stdout=stream,stderr=subprocess.STDOUT)
    rows.append({'command':command,'exit_code':r.returncode,'log':c.artifact(log),'started':start,'ended':time.time()});c.write(out/'checks.json',rows)
    assert r.returncode==0,log
    assert c.source()==frozen
c.write(out/'complete.json',{'checks':c.artifact(out/'checks.json'),'source':frozen})
print('Probe checks passed',flush=True)
