"""Root-only serial capture; frozen schedule and no overwrite or automatic retries."""
import subprocess,sys,time
import custody as c
lane=sys.argv[1]
assert lane in ('qualification','native','allocation')
plan=c.read(c.P/'plan.json');out=c.P/lane;assert not out.exists();out.mkdir()
legs=['before'] if lane=='qualification' else ['before','after']
builds={leg:c.read(c.P/f'build-{leg}/build.json') for leg in legs}
kind='allocation' if lane in ('qualification','allocation') else 'native'
for build in builds.values():
 for binary in build['binaries'].values():assert c.artifact(binary['path'])==binary
frozen=c.source();c.write(out/'source.json',frozen)
rows=[]
blocks=1 if lane=='qualification' else plan[lane]['blocks']
for block in range(blocks):
 for case in plan['cases']:
  order=['before'] if lane=='qualification' else plan['native']['orders'][block]
  for leg in order:
   stem=f"{block}-{case['shape']}-{case['mode']}-{leg}";report=out/f'{stem}.json';rss=out/f'{stem}.rss';log=out/f'{stem}.log'
   samples=1 if lane=='qualification' else plan[lane]['samples'];warmup=0 if lane=='qualification' else plan[lane]['warmup']
   binary=builds[leg]['binaries'][kind]
   command=['/usr/bin/time','-f','%M','-o',str(rss),'taskset','-c',str(plan['cpu']),binary['path'],'--mode',case['mode'],'--shape',case['shape'],'--samples',str(samples),'--warmup',str(warmup),'--output',str(report)]
   start=time.time()
   with log.open('w') as f:r=subprocess.run(command,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
   row={'lane':lane,'block':block,**case,'leg':leg,'command':command,'exit_code':r.returncode,'started':start,'ended':time.time(),'binary':binary,'log':c.artifact(log),'rss':c.artifact(rss)}
   if report.exists():row['report']=c.artifact(report)
   rows.append(row);c.write(out/'receipts.json',rows)
   assert r.returncode==0,row
   data=c.read(report);assert len(data['samples'])==samples
   assert c.source()==frozen
   print(stem,'passed',flush=True)
c.write(out/'complete.json',{'children':len(rows),'receipts':c.artifact(out/'receipts.json'),'source':c.artifact(out/'source.json')})
