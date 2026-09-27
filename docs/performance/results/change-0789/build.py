"""Root-only strict normal/sanitized calibration builds and bounded qualification."""
import os,subprocess
import common as c
assert not c.TARGET.exists();c.TARGET.mkdir()
assert not (c.P/'build.json').exists()
inputs={n:c.sha(c.P/n) for n in ['probe.c','plan.json','capture.py','build.py','common.py','architecture-inputs.json','host.json','origin.json','production-source.json']}
c.write(c.P/'frozen-inputs.json',inputs)
rows=[];binaries={}
for name,flags in [('probe',[]),('probe-sanitized',['-fsanitize=address,undefined','-fno-omit-frame-pointer'])]:
 dest=c.TARGET/name;cmd=['/usr/bin/cc','-std=c11','-O2','-g','-Wall','-Wextra','-Werror','-pthread',*flags,str(c.P/'probe.c'),'-o',str(dest)];log=c.P/f'build-{name}.log'
 with log.open('wb') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'exit_code':r.returncode,'log':c.artifact(log)})
 c.write(c.P/'build-commands.json',rows);assert r.returncode==0,log
 binaries[name]=c.artifact(dest)
 assert all(c.sha(c.P/n)==h for n,h in inputs.items())
c.write(c.P/'build.json',{'commands':rows,'binary':binaries['probe'],'sanitized_binary':binaries['probe-sanitized'],'inputs':c.artifact(c.P/'frozen-inputs.json')})
out=c.P/'qualification';out.mkdir();checks=[]
base=['--mib','0','--workers','0','--touch','main']
cases=[('mib_over_limit',['--mib','65','--workers','0','--touch','main'],b'',False),('workers_over_limit',['--mib','0','--workers','33','--touch','main'],b'',False),('unknown_touch',['--mib','0','--workers','0','--touch','invalid'],b'',False),('workers_touch_without_workers',['--mib','0','--workers','0','--touch','workers'],b'',False),('eof_ack',base,b'',False),('wrong_ack',base,b'-\n',False)]
for mib,workers,touch in [(0,0,'main'),(64,4,'workers'),(4,32,'main')]:cases.append((f'sanitized-{mib}-{workers}-{touch}',['--mib',str(mib),'--workers',str(workers),'--touch',touch],b'+\n'*6,True))
for name,args,stdin,positive in cases:
 exe=binaries['probe-sanitized' if positive else 'probe'];cmd=[exe['path'],*args];r=subprocess.run(cmd,input=stdin,stdout=subprocess.PIPE,stderr=subprocess.PIPE,env=os.environ|{'ASAN_OPTIONS':'detect_leaks=1:abort_on_error=1','UBSAN_OPTIONS':'halt_on_error=1'})
 stdout=out/f'{name}.stdout.txt';stderr=out/f'{name}.stderr.log';stdout.write_bytes(r.stdout);stderr.write_bytes(r.stderr)
 checks.append({'name':name,'command':cmd,'binary':exe,'stdin_hex':stdin.hex(),'exit_code':r.returncode,'expected_success':positive,'stdout':c.artifact(stdout),'stderr':c.artifact(stderr)})
 c.write(out/'checks.json',checks)
 assert (r.returncode==0)==positive,(name,r.returncode,r.stderr)
 if positive:assert len(r.stdout.splitlines())==12 and not r.stderr
print('Strict builds and nine qualification checks PASS')
