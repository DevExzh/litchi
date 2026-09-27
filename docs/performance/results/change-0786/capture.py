"""Root-only serial child matrix, frozen schedule, no automatic retries."""
import subprocess, sys, time
import custody as c
lane=sys.argv[1];assert lane in ('qualification','native','observer')
plan=c.read(c.P/'plan.json');out=c.P/lane;assert not out.exists();out.mkdir()
build=c.read(c.P/'build.json');kind='native' if lane=='native' else 'observer'
binary=build['binaries'][kind];assert c.artifact(binary['path'])==binary
frozen=c.read(build['source']['path']);c.unchanged(frozen)
c.write(out/'source.json',frozen)
rows=[];spec=plan[lane]
for block in range(spec['blocks']):
 cases=plan['cases'] if spec['orders'][block]=='forward' else list(reversed(plan['cases']))
 for case in cases:
  stem=f"{block}-{case['route']}-{case['shape']}-{case['state']}-{case['task_floor']}-{case['workers']}"
  report=out/f'{stem}.json';rss=out/f'{stem}.rss';log=out/f'{stem}.log'
  cmd=['/usr/bin/time','-f','%M','-o',str(rss),'taskset','-c',','.join(map(str,plan['affinity'])),binary['path']]
  for key,value in case.items():cmd.extend(['--'+key.replace('_','-'),str(value)])
  cmd.extend(['--samples',str(spec['samples']),'--warmup',str(spec['warmup']),'--output',str(report)])
  started=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
  row={'lane':lane,'block':block,**case,'command':cmd,'exit_code':r.returncode,'started':started,'ended':time.time(),'binary':binary,'log':c.artifact(log),'rss':c.artifact(rss)}
  if report.exists():row['report']=c.artifact(report)
  rows.append(row);c.write(out/'receipts.json',rows)
  assert r.returncode==0,row
  assert len(c.read(report)['samples'])==spec['samples']
  assert c.tool_source()==frozen['tool']
  print(stem,'passed',flush=True)
 c.unchanged(frozen)
c.write(out/'complete.json',{'children':len(rows),'receipts':c.artifact(out/'receipts.json'),'source':c.artifact(out/'source.json')})
