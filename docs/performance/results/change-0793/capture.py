"""Root-only serial diagnostic capture, with no overwrite or automatic retry."""
import subprocess,sys,time
import custody as c
lane=sys.argv[1];assert lane in ['controls','heaptrack']
plan=c.read(c.P/'plan.json');build=c.read(c.P/'build/receipt.json');source=c.read(c.P/'build/source.json')
assert c.source()==source
for v in build['binaries'].values():assert c.artifact(v['path'])==v
out=c.P/lane;assert not out.exists();out.mkdir();rows=[]
if lane=='controls':jobs=[(r,s,v) for r,order in enumerate(plan[lane]['orders']) for s in plan['shapes'] for v in order]
else:jobs=[(r,s,'profile') for r,shapes in enumerate(plan[lane]['orders']) for s in shapes]
for repeat,shape,variant in jobs:
 stem=f'{repeat}-{shape}-{variant}';report=out/f'{stem}.json';log=out/f'{stem}.log';binary=build['binaries'][variant]
 cmd=['taskset','-c',str(plan['cpu'])]
 if lane=='heaptrack':cmd+=['heaptrack','--record-only','-o',str(out/f'{stem}.heaptrack')]
 cmd += [binary['path'],'--mode','capture','--shape',shape,'--samples',str(plan[lane]['samples']),'--warmup','0','--output',str(report)]
 started=time.time()
 with log.open('w') as stream:r=subprocess.run(cmd,cwd=c.ROOT,stdout=stream,stderr=subprocess.STDOUT)
 row={'repeat':repeat,'shape':shape,'variant':variant,'command':cmd,'started':started,'ended':time.time(),'exit_code':r.returncode,'binary':binary,'log':c.artifact(log)}
 if report.exists():row['report']=c.artifact(report)
 if lane=='heaptrack':row['traces']=[c.artifact(f) for f in out.glob(stem+'.heaptrack*') if f.is_file()]
 rows.append(row);c.write(out/'receipts.json',rows)
 assert r.returncode==0,log
 assert len(c.read(report)['samples'])==plan[lane]['samples']
 if lane=='heaptrack':assert len(row['traces'])==1,row
 assert c.source()==source
 print(stem,'PASS',flush=True)
c.write(out/'complete.json',{'children':len(rows),'receipts':c.artifact(out/'receipts.json'),'source':c.artifact(c.P/'build/source.json')})
