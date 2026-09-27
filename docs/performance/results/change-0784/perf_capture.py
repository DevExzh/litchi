"""Native sampled follow-up; exact capture-owner qualification is offline."""
import subprocess,time
import custody as c
plan=c.read(c.P/'perf-plan.json');build=c.read(c.P/'build/build.json');binary=build['binaries']['profile']
assert c.source()==c.read(c.P/'build/source.json') and c.sha(binary['path'])==binary['sha256']
out=c.P/'perf';assert not out.exists();out.mkdir();rows=[]
for repeat in range(plan['repeats']):
 report=out/f'{repeat}.json';raw=out/f'{repeat}.data';log=out/f'{repeat}.log'
 args=[binary['path'],'--mode','capture','--shape',plan['shape'],'--samples',str(plan['samples']),'--warmup',str(plan['warmup']),'--output',str(report)]
 cmd=['taskset','-c',str(plan['cpu']),'perf','record','-e',plan['event'],'-F',str(plan['frequency']),'--call-graph',plan['call_graph'],'-o',str(raw),'--',*args]
 start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
 row={'repeat':repeat,'command':cmd,'started':start,'ended':time.time(),'exit_code':r.returncode,'plan_sha256':c.sha(c.P/'perf-plan.json'),'binary':binary,'log':c.artifact(log)}
 if report.exists():row['report']=c.artifact(report)
 if raw.exists():row['raw']=c.artifact(raw)
 rows.append(row);c.write(out/'receipts.json',rows)
 assert r.returncode==0,log
 assert c.source()==c.read(c.P/'build/source.json')
 print('perf',repeat,'complete',flush=True)
c.write(out/'complete.json',{'processes':len(rows),'plan_sha256':c.sha(c.P/'perf-plan.json')})
