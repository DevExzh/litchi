"""Frozen alternating serial native runs. Never retries or replaces reports."""
import subprocess,time
import custody as c
plan=c.read(c.P/'plan.json');build=c.read(c.P/'build/build.json')
assert c.source()==c.read(c.P/'build/source.json')
assert not (c.P/'native').exists()
out=c.P/'native';out.mkdir();rows=[]
for block,order in enumerate(plan['native']['orders']):
 for shape in plan['native']['shapes']:
  for leg in order:
   binary=build['binaries'][leg];assert c.sha(binary['path'])==binary['sha256']
   stem=f'{block}-{shape}-{leg}';report=out/f'{stem}.json';rss=out/f'{stem}.rss';log=out/f'{stem}.log'
   cmd=['/usr/bin/time','-f','%M','-o',str(rss),'taskset','-c',str(plan['cpu']),binary['path'],'--mode','capture','--shape',shape,'--samples',str(plan['native']['samples']),'--warmup',str(plan['native']['warmup']),'--output',str(report)]
   start=time.time()
   with log.open('w') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
   row={'block':block,'shape':shape,'leg':leg,'command':cmd,'started':start,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log),'rss':c.artifact(rss)}
   if report.exists():row['report']=c.artifact(report)
   rows.append(row);c.write(out/'receipts.json',rows)
   assert r.returncode==0,log
   assert c.source()==c.read(c.P/'build/source.json')
   print(stem,'PASS',flush=True)
c.write(out/'complete.json',{'processes':len(rows),'plan_sha256':c.sha(c.P/'plan.json'),'build_sha256':c.sha(c.P/'build/build.json')})
