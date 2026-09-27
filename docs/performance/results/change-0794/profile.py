"""Root-only serial large-capture heap traces with predeclared counter semantics."""
import subprocess,time
import custody as c
p=c.P;out=p/'profile';assert not out.exists();out.mkdir()
plan=c.read(p/'plan.json');source=c.source();c.write(out/'source.json',source)
builds={leg:c.read(p/f'build-{leg}/build.json') for leg in ['before','after']}
rows=[];decodes=[]
for repeat,order in enumerate(plan['profile']['orders']):
 for leg in order:
  binary=builds[leg]['binaries']['profile'];assert c.artifact(binary['path'])==binary
  stem=f'{repeat}-{leg}';report=out/f'{stem}.json';log=out/f'{stem}.log'
  command=['taskset','-c',str(plan['cpu']),'heaptrack','--record-only','-o',str(out/f'{stem}.heaptrack'),binary['path'],'--mode','capture','--shape','large','--samples','1','--warmup','0','--output',str(report)]
  started=time.time()
  with log.open('w') as stream:r=subprocess.run(command,cwd=c.ROOT,stdout=stream,stderr=subprocess.STDOUT)
  row={'repeat':repeat,'leg':leg,'binary':binary,'command':command,'exit_code':r.returncode,'started':started,'ended':time.time(),'log':c.artifact(log),'traces':[c.artifact(x) for x in out.glob(stem+'.heaptrack*') if x.is_file()]}
  if report.exists():row['report']=c.artifact(report)
  rows.append(row);c.write(out/'receipts.json',rows)
  assert r.returncode==0 and len(row['traces'])==1
  assert c.source()==source
  print(stem,'captured',flush=True)
for row in rows:
 for scope in ['whole','owner']:
  stem=f"{row['repeat']}-{row['leg']}-{scope}";trace=row['traces'][0]
  stacks=out/f'{stem}.stacks';hist=out/f'{stem}.histogram';log=out/f'{stem}.decoded.log';err=out/f'{stem}.stderr'
  command=['heaptrack_print','-f',trace['path'],'-m','0','-t','0','-p','1','-a','1','-T','0','-n','20','--flamegraph-cost-type','allocations','-F',str(stacks),'-H',str(hist)]
  if scope=='owner':command+=['--filter-bt-function','capture_region_0793']
  started=time.time()
  with log.open('w') as stdout,err.open('w') as stderr:r=subprocess.run(command,cwd=c.ROOT,stdout=stdout,stderr=stderr)
  decoded={'repeat':row['repeat'],'leg':row['leg'],'scope':scope,'trace':trace,'command':command,'exit_code':r.returncode,'started':started,'ended':time.time(),'log':c.artifact(log),'stderr':c.artifact(err)}
  for key,path in [('stacks',stacks),('histogram',hist)]:
   if path.exists():decoded[key]=c.artifact(path)
  decodes.append(decoded);c.write(out/'decodes.json',decodes)
  assert r.returncode==0 and stacks.is_file() and hist.is_file()
  print(stem,'decoded',flush=True)
c.write(out/'complete.json',{'children':len(rows),'decodes':len(decodes),'receipts':c.artifact(out/'receipts.json'),'decode_receipts':c.artifact(out/'decodes.json'),'source':c.artifact(out/'source.json')})
