"""Root-only serial paired capture, no overwrites or retries."""
import subprocess,sys,time
import custody as c
lane=sys.argv[1];assert lane in ('qualification-before','qualification-after','native','observer')
qualification=lane.startswith('qualification-');legs=[lane.split('-')[1]] if qualification else ['before','after'];specname='qualification' if qualification else lane
plan=c.read(c.P/'plan.json');spec=plan[specname];out=c.P/lane;assert not out.exists();out.mkdir()
builds={leg:c.read(c.P/f'build-{leg}/build.json') for leg in legs};kind='native' if lane=='native' else 'observer'
for build in builds.values():
 for binary in build['binaries'].values():assert c.artifact(binary['path'])==binary
frozen={'production':c.source(),'tool':c.tool_source()};c.write(out/'source.json',frozen)
rows=[]
for block in range(spec['blocks']):
 for case in plan['cases']:
  order=legs if qualification else spec['orders'][block]
  for leg in order:
   stem=f"{block}-{case['shape']}-{case['state']}-{case['task_floor']}-{case['workers']}-{leg}";report=out/f'{stem}.json';rss=out/f'{stem}.rss';log=out/f'{stem}.log';binary=builds[leg]['binaries'][kind]
   cmd=['/usr/bin/time','-f','%M','-o',str(rss),'taskset','-c',','.join(map(str,plan['affinity'])),binary['path']]
   for key,value in case.items():cmd.extend(['--'+key.replace('_','-'),str(value)])
   cmd.extend(['--samples',str(spec['samples']),'--warmup',str(spec['warmup']),'--output',str(report)])
   start=time.time()
   with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
   row={'lane':lane,'block':block,'leg':leg,**case,'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'binary':binary,'log':c.artifact(log),'rss':c.artifact(rss)}
   if report.exists():row['report']=c.artifact(report)
   rows.append(row);c.write(out/'receipts.json',rows);assert r.returncode==0,row
   assert len(c.read(report)['samples'])==spec['samples'];assert c.tool_source()==frozen['tool']
   print(stem,'passed',flush=True)
 c.unchanged(frozen)
c.write(out/'complete.json',{'children':len(rows),'receipts':c.artifact(out/'receipts.json'),'source':c.artifact(out/'source.json')})
