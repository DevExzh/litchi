"""Root-only syscall diagnostics, never native timing evidence."""
import subprocess,time
import custody as c
plan=c.read(c.P/'trace-plan.json');build=c.read(c.P/'build-after/build.json');binary=build['binaries']['observer'];assert c.artifact(binary['path'])==binary
out=c.P/'traces-after';assert not out.exists();out.mkdir();frozen=c.source();rows=[]
for repeat in range(plan['repeats']):
 for width in plan['widths'] if repeat==0 else reversed(plan['widths']):
  for state in plan['states']:
   stem=f'{repeat}-{width}-{state}';trace=out/f'{stem}.strace';report=out/f'{stem}.json';log=out/f'{stem}.log'
   cmd=['taskset','-c',','.join(map(str,plan['affinity'])),'strace','-f','-c','-e','trace=clone,clone3,futex','-o',str(trace),binary['path'],'--route','parts','--shape','large','--state',state,'--workers',str(width),'--task-floor','0','--samples','1','--warmup','0','--output',str(report)]
   start=time.time()
   with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
   row={'repeat':repeat,'workers':width,'state':state,'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'binary':binary,'plan_sha256':c.sha(c.P/'trace-plan.json'),'artifacts':{f.name:c.artifact(f) for f in [trace,report,log] if f.exists()}}
   rows.append(row);c.write(out/'receipts.json',rows);assert r.returncode==0,row;assert c.source()==frozen
   print(stem,'traced',flush=True)
c.write(out/'complete.json',{'processes':len(rows),'receipts':c.artifact(out/'receipts.json')})
