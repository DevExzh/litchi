"""Fresh launcher exit and custody checks; no RSS inference from these controls."""
import driver as d
import subprocess
launcher=d.TARGET/'launcher';cases=[('success',['/bin/true'],0),('failure',['/bin/false'],1),('missing',['/nonexistent-0839'],127)]
rows=[]
for label,args,expected in cases:
 p=d.P/'launcher-checks'/f'{label}.json';p.parent.mkdir(exist_ok=True)
 result=subprocess.run([str(launcher),'--usage',str(p),*args],capture_output=True)
 assert result.returncode==expected
 r=d.read(p);assert r['exit_code']==expected and r['child_pid']>0
 rows.append(dict(label=label,exit_code=result.returncode,usage=d.desc(p)))
for args in [[],['--usage',str(d.SCRATCH/'invalid'),'relative']]:
 result=subprocess.run([str(launcher),*args],capture_output=True);assert result.returncode==2
rows.append(dict(label='invalid_invocations',cases=2,exit_code=2))
d.write(d.P/'launcher-quality.json',dict(status='pass',binary=d.desc(launcher),cases=rows))
print('launcher quality PASS')
