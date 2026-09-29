"""Two explicit tiny observer compatibility checks, outside formal populations."""
import driver as d
import time
plan=d.read(d.P/'plan.json');build=d.read(d.P/'build.json');assert d.desc(build['binary']['path'])==build['binary']
rows=[]
for lane in ['profiles','heap']:
 row=dict(parts=4,mode='publish',block=0,samples=1,warmup=0);label='observer-preflight-'+lane
 out=d.P/'observer-preflight'/lane;out.mkdir(parents=True,exist_ok=False)
 argv=['taskset','-c',str(plan['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(out/'time.txt')]
 if lane=='profiles':argv+=['valgrind',*plan['profile_options'],'--callgrind-out-file='+str(out/'profile.callgrind')]
 else:argv+=['heaptrack','-o',str(out/'trace')]
 argv+=[build['binary']['path'],'--parts','4','--mode','publish','--samples','1','--warmup','0','--output',str(out/'report.json')]
 d.write(out/'started.json',dict(lane=lane,row=row,binary=build['binary'],build=d.desc(d.P/'build.json'),argv=argv,started_unix=time.time()))
 code=d.run(label,argv)
 d.write(out/'receipt.json',dict(lane=lane,row=row,exit_code=code,command=d.desc(d.P/'commands'/label/'receipt.json'),artifacts=[d.desc(p) for p in sorted(out.rglob('*')) if p.is_file()],finished_unix=time.time()))
 rows.append(d.desc(out/'receipt.json'));assert code==0 and (out/'report.json').is_file()
d.write(d.P/'observer-preflight-capture.json',dict(status='pass',receipts=rows,scope='Compatibility only; excluded from the72formal report/1464sample population.',script=d.desc(d.P/'observer_preflight.py')))
