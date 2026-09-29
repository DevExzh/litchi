"""Separate tiny Heaptrack compatibility capture for each immutable binary."""
import driver as d
import sys,time
leg=sys.argv[1];assert leg in ['before','after'];lane='heap';plan=d.read(d.P/'plan.json')
row=next(x for x in plan['heap'] if x['leg']==leg and x['block']==0 and x['shape']=='shared' and x['parts']==4)
build=d.read(d.P/f'build-{leg}.json');assert d.desc(build['binary']['path'])==build['binary']
out=d.P/'observer-preflight'/leg;out.mkdir(parents=True,exist_ok=False)
argv=['taskset','-c',str(plan['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(out/'time.txt'),'heaptrack','-o',str(out/'trace'),build['binary']['path'],'--shape',row['shape'],'--parts',str(row['parts']),'--mode',row['mode'],'--samples','1','--warmup','0','--output',str(out/'report.json')]
d.write(out/'started.json',dict(lane=lane,row=row,binary=build['binary'],build=d.desc(d.P/f'build-{leg}.json'),argv=argv,started_unix=time.time()))
code=d.run('observer-preflight-'+leg,argv)
d.write(out/'receipt.json',dict(lane=lane,row=row,exit_code=code,command=d.desc(d.P/'commands'/('observer-preflight-'+leg)/'receipt.json'),artifacts=[d.desc(p) for p in sorted(out.rglob('*')) if p.is_file()],finished_unix=time.time()))
assert code==0
