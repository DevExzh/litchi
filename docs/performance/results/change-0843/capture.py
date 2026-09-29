"""Serial immutable diagnostic schedule, with separate observer populations."""
import driver as d
import sys,time
lane=sys.argv[1];assert lane in ['qualification','native','profiles','heap']
plan=d.read(d.P/'plan.json');build=d.read(d.P/'build.json');assert d.desc(build['binary']['path'])==build['binary']
freeze=d.read(build['freeze']['path']);assert d.source()==freeze['source']
assert all(d.sha(d.P/n)==h for n,h in freeze['probe'].items())
if lane!='qualification':assert all(d.desc(x['path'])==x for x in d.read(d.P/'admission.json')['files'])
receipts=[]
for row in plan[lane]:
 label=f"{lane}-{row['block']:02}-{row['parts']:04}-{row['mode']}"
 out=d.P/'runs'/label;out.mkdir(parents=True,exist_ok=False)
 argv=['taskset','-c',str(plan['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(out/'time.txt')]
 if lane=='profiles':argv+=['valgrind',*plan['profile_options'],'--callgrind-out-file='+str(out/'profile.callgrind')]
 if lane=='heap':argv+=['heaptrack','-o',str(out/'trace')]
 argv+=[build['binary']['path'],'--parts',str(row['parts']),'--mode',row['mode'],'--samples',str(row['samples']),'--warmup',str(row['warmup']),'--output',str(out/'report.json')]
 if lane=='qualification':argv+=['--artifact',str(out/'artifact.zip')]
 d.write(out/'started.json',dict(lane=lane,row=row,binary=build['binary'],build=d.desc(d.P/'build.json'),argv=argv,started_unix=time.time()))
 code=d.run(label,argv)
 d.write(out/'receipt.json',dict(lane=lane,row=row,exit_code=code,command=d.desc(d.P/'commands'/label/'receipt.json'),artifacts=[d.desc(p) for p in sorted(out.rglob('*')) if p.is_file()],finished_unix=time.time()))
 receipts.append(d.desc(out/'receipt.json'));assert code==0,(label,code)
 assert (out/'report.json').is_file(),label
 print(label,'captured',flush=True)
d.write(d.P/f'capture-{lane}.json',dict(status='pass',receipts=receipts,plan=d.desc(d.P/'plan.json')))
