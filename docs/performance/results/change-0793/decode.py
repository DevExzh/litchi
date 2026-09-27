"""Root-only serial heaptrack decoding. Retains full and owner-filtered outputs."""
import subprocess,time
import custody as c
out=c.P/'decoded';assert not out.exists();out.mkdir();rows=[]
plan=c.read(c.P/'plan.json')
for row in c.read(c.P/'heaptrack/receipts.json'):
 trace=row['traces'][0];assert c.artifact(trace['path'])==trace
 for scope in ['whole','owner']:
  stem=f"{row['repeat']}-{row['shape']}-{scope}";stacks=out/f'{stem}.stacks';hist=out/f'{stem}.histogram';log=out/f'{stem}.log';err=out/f'{stem}.stderr'
  cmd=['heaptrack_print','-f',trace['path'],'-m','0','-t','0','-p','1','-a','1','-T','0','-n','20','--flamegraph-cost-type','allocations','-F',str(stacks),'-H',str(hist)]
  if scope=='owner':cmd+=['--filter-bt-function',plan['heaptrack']['filter']]
  start=time.time()
  with log.open('w') as stdout,err.open('w') as stderr:r=subprocess.run(cmd,cwd=c.ROOT,stdout=stdout,stderr=stderr)
  result={'repeat':row['repeat'],'shape':row['shape'],'scope':scope,'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'trace':trace,'log':c.artifact(log),'stderr':c.artifact(err)}
  if stacks.exists():result['stacks']=c.artifact(stacks)
  if hist.exists():result['histogram']=c.artifact(hist)
  rows.append(result);c.write(out/'receipts.json',rows)
  assert r.returncode==0,result
  assert stacks.is_file() and hist.is_file(),result
  print(stem,'decoded',flush=True)
c.write(out/'complete.json',{'decodes':len(rows),'receipts':c.artifact(out/'receipts.json')})
