"""Root-only observer capture. Observer elapsed time is never native timing."""
import subprocess,sys,time
import custody as c
lane=sys.argv[1];assert lane in ('perf','heaptrack-before','heaptrack-after')
plan=c.read(c.P/'observer-plan.json');out=c.P/lane;assert not out.exists();out.mkdir()
frozen=c.source();c.write(out/'source.json',frozen);rows=[]
if lane=='perf':
 jobs=[(block,case,leg) for block in range(plan['perf_stat']['blocks']) for case in plan['perf_stat']['cases'] for leg in c.read(c.P/'plan.json')['native']['orders'][block]]
else:
 jobs=[(0,{'shape':plan['heaptrack']['shape'],'mode':plan['heaptrack']['mode']},lane.split('-')[-1])]
for block,case,leg in jobs:
 kind='perf_stat' if lane=='perf' else 'heaptrack';policy=plan[kind]
 binary=c.read(c.P/f'build-{leg}/build.json')['binaries']['native'];assert c.artifact(binary['path'])==binary
 stem=f"{block}-{case['shape']}-{case['mode']}-{leg}";report=out/f'{stem}.json';log=out/f'{stem}.log';stats=out/f'{stem}.stats'
 args=[binary['path'],'--mode',case['mode'],'--shape',case['shape'],'--samples',str(policy['samples']),'--warmup',str(policy['warmup']),'--output',str(report)]
 prefix=['perf','stat','-x',';','-e',policy['events'],'-o',str(stats),'--'] if lane=='perf' else ['heaptrack','-o',str(out/'trace')]
 command=['taskset','-c',str(plan['cpu']),*prefix,*args];start=time.time()
 with log.open('w') as f:r=subprocess.run(command,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
 row={'leg':leg,'block':block,**case,'command':command,'binary':binary,'exit_code':r.returncode,'started':start,'ended':time.time(),'log':c.artifact(log)}
 if report.exists():row['report']=c.artifact(report)
 if stats.exists():row['stats']=c.artifact(stats)
 if lane!='perf':row['traces']=[c.artifact(v) for v in sorted(out.glob('trace*')) if v.is_file()]
 rows.append(row);c.write(out/'receipts.json',rows)
 assert r.returncode==0,row
 assert c.source()==frozen
 print(stem,'observer complete',flush=True)
c.write(out/'complete.json',{'children':len(rows),'receipts':c.artifact(out/'receipts.json')})
