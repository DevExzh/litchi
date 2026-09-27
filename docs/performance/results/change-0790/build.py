"""Root-only strict builds; qualification excludes captured matrix."""
import os,subprocess
import common as c
assert not c.TARGET.exists();c.TARGET.mkdir()
inputs={n:c.sha(c.P/n) for n in ['probe.c','launcher.c','capture.py','build.py','common.py','plan.json','host.json','origin.json','architecture-inputs.json','inherited.json']}
c.write(c.P/'frozen-inputs.json',inputs);build={};commands=[]
for key,source,flags in [('binary','probe.c',['-pthread']),('launcher','launcher.c',[]),('sanitized_launcher','launcher.c',['-fsanitize=address,undefined','-fno-omit-frame-pointer'])]:
 dest=c.TARGET/key;cmd=['/usr/bin/cc','-std=c11','-O2','-g','-Wall','-Wextra','-Werror',*flags,str(c.P/source),'-o',str(dest)];log=c.P/f'build-{key}.log'
 with log.open('wb') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
 commands.append({'command':cmd,'exit_code':r.returncode,'log':c.artifact(log)});c.write(c.P/'build-commands.json',commands);assert r.returncode==0
 build[key]=c.artifact(dest)
build.update(commands=commands,inputs=c.artifact(c.P/'frozen-inputs.json'));c.write(c.P/'build.json',build)
out=c.P/'qualification';out.mkdir();rows=[]
for name,args,stdin,expected in [('invalid',[],b'',2),('relative',['relative'],b'',2),('missing',['/does-not-exist-0790'],b'',127),('false',['/usr/bin/false'],b'',1),('signal',['/bin/sh','-c','kill -TERM $$'],b'',143),('zero',[build['binary']['path'],'--mib','0','--workers','0','--touch','main'],b'+\n'*6,0),('payload',[build['binary']['path'],'--mib','64','--workers','4','--touch','workers'],b'+\n'*6,0)]:
 usage=out/f'{name}.usage.json';cmd=[build['sanitized_launcher']['path']]
 if name!='invalid':cmd+=['--usage',str(usage),*args]
 r=subprocess.run(cmd,input=stdin,stdout=subprocess.PIPE,stderr=subprocess.PIPE,env=os.environ|{'ASAN_OPTIONS':'detect_leaks=1:abort_on_error=1','UBSAN_OPTIONS':'halt_on_error=1'})
 stdout=out/f'{name}.stdout.txt';stderr=out/f'{name}.stderr.log';stdout.write_bytes(r.stdout);stderr.write_bytes(r.stderr)
 row={'name':name,'command':cmd,'exit_code':r.returncode,'expected':expected,'stdout':c.artifact(stdout),'stderr':c.artifact(stderr)}
 if usage.exists():row['usage']=c.artifact(usage)
 rows.append(row);c.write(out/'checks.json',rows);assert r.returncode==expected
 if expected!=2:assert not r.stderr and c.read(usage)['exit_code']==expected
 if expected==0:assert len(r.stdout.splitlines())==12
assert all(c.sha(c.P/n)==h for n,h in inputs.items())
print('Three strict builds and seven launcher qualifications PASS')
