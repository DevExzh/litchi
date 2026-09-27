"""Root-only diagnostic process schedule; no retries or timing claims."""
import sys,subprocess,time
import custody as c
p=c.P;lane=sys.argv[1];assert lane in ['control','census'];plan=c.read(p/'plan.json');settings=plan[lane];leg=settings['leg'];build=c.read(p/('build-'+leg)/'build.json');binary=build['binary']
assert c.artifact(binary['path'])==binary
source=c.source();assert source==c.read(p/'source.json')
out=p/lane;assert not out.exists();out.mkdir();c.write(out/'source.json',source);rows=[]
for block in range(settings['blocks']):
 cases=plan['cases'] if block==0 else list(reversed(plan['cases']))
 for case in cases:
  stem=f"{block}-{case['shape']}-{case['mode']}";report=out/(stem+'.json');log=out/(stem+'.log')
  cmd=['taskset','-c',str(plan['cpu']),binary['path'],'--shape',case['shape'],'--mode',case['mode'],'--samples','1','--warmup','0','--output',str(report)]
  start=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
  row={'lane':lane,'block':block,**case,'leg':leg,'command':cmd,'binary':binary,'started':start,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log)}
  if report.exists():row['report']=c.artifact(report)
  rows.append(row);c.write(out/'receipts.json',rows);assert r.returncode==0 and report.exists(),log
  print(lane,stem,'PASS',flush=True)
 assert c.source()==source
assert c.artifact(binary['path'])==binary
c.write(out/'complete.json',{'children':len(rows),'source':c.artifact(out/'source.json'),'receipts':c.artifact(out/'receipts.json')})
