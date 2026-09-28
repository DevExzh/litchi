"""Root-only serial 0806 cross-format capture; no retries or overwrite."""
import subprocess,sys,time
import custody as c
lane=sys.argv[1];assert lane in ['qualification','native']
plan=c.read(c.P/'cross-plan.json');out=c.P/f'cross-{lane}';assert not out.exists();out.mkdir()
source=c.source();c.write(out/'source.json',source)
legs=['before'] if lane=='qualification' else ['before','after']
builds={leg:c.read(c.P/f'cross-build-{leg}/receipt.json') for leg in legs}
for build in builds.values():assert c.artifact(build['binary']['path'])==build['binary']
rows=[]
orders=[['before']] if lane=='qualification' else plan['orders']
for block,order in enumerate(orders):
 for leg in order:
  stem=f'{block}-{leg}';report=out/f'{stem}.json';log=out/f'{stem}.log';rss=out/f'{stem}.rss';binary=builds[leg]['binary']
  samples=1 if lane=='qualification' else plan['samples'];warmup=0 if lane=='qualification' else plan['warmup']
  cmd=['/usr/bin/time','-f','%M','-o',str(rss),'taskset','-c',str(plan['cpu']),binary['path'],'--warmup',str(warmup),'--samples',str(samples),'--semantic-shape',','.join(plan['semantic_shapes']),'--xlsx-shape',','.join(plan['xlsx_shapes']),'--case',','.join(plan['cases']),'--json',str(report)]
  started=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
  row={'block':block,'leg':leg,'command':cmd,'exit_code':r.returncode,'started':started,'ended':time.time(),'binary':binary,'log':c.artifact(log),'rss':c.artifact(rss)}
  if report.exists():row['report']=c.artifact(report)
  rows.append(row);c.write(out/'receipts.json',rows)
  assert r.returncode==0 and report.exists(),log
  assert c.source()==source
  print(lane,stem,'PASS',flush=True)
c.write(out/'complete.json',{'children':len(rows),'receipts':c.artifact(out/'receipts.json'),'source':c.artifact(out/'source.json')})
