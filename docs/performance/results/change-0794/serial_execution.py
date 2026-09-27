"""Root-only, fail-closed continuation after the final candidate quality run."""
import subprocess,sys,time
import custody as c
quality=c.read(c.P/'quality.json')
assert len(quality['rows'])==6 and all(r['exit_code']==0 for r in quality['rows'])
assert c.read(quality['source']['path'])==c.source()==c.read(c.P/'application.json')['source']
steps=[('build-after',['build.py','after']),('attribute-after',['attribute_diagnostic.py','after']),('cross-build-after',['cross_build.py','after']),('native',['capture.py','native']),('cross-native',['cross_capture.py','native']),('allocation',['capture.py','allocation']),('profile',['profile.py'])]
rows=[]
assert not (c.P/'serial-execution.json').exists()
for name,args in steps:
    log=c.P/f'{name}-driver-final.log';assert not log.exists()
    command=[sys.executable,'-B',str(c.P/args[0]),*args[1:]]
    started=time.time()
    print('START',name,flush=True)
    with log.open('w') as out:r=subprocess.run(command,stdout=out,stderr=subprocess.STDOUT,cwd=c.ROOT)
    rows.append({'name':name,'command':command,'started':started,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)})
    c.write(c.P/'serial-execution.json',{'driver':c.artifact(c.P/'serial_execution.py'),'rows':rows})
    assert r.returncode==0,log
    print('PASS',name,flush=True)
