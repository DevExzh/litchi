"""Fixed serial matched experiment; all builds finish before native capture."""
import driver as d
import sys,time
lane=sys.argv[1];assert lane in ['qualification','native','heap'];leg=sys.argv[2] if len(sys.argv)>2 else None
assert (lane=='qualification')==(leg in ['before','after'])
plan=d.read(d.P/'plan.json')
if lane!='qualification':assert all(d.desc(x['path'])==x for x in d.read(d.P/'admission.json')['files'])
receipts=[]
for row in plan[lane]:
 if leg is not None and row['leg']!=leg:continue
 build=d.read(d.P/f"build-{row['leg']}.json");assert d.desc(build['binary']['path'])==build['binary']
 label=f"{lane}-{row['block']:02}-{row['shape']}-{row['parts']:04}-{row['mode']}-{row['leg']}"
 out=d.P/'runs'/label;out.mkdir(parents=True,exist_ok=False)
 argv=['taskset','-c',str(plan['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(out/'time.txt')]
 if lane=='heap':argv+=['heaptrack','-o',str(out/'trace')]
 argv+=[build['binary']['path'],'--shape',row['shape'],'--parts',str(row['parts']),'--mode',row['mode'],'--samples',str(row['samples']),'--warmup',str(row['warmup']),'--output',str(out/'report.json')]
 if lane=='qualification':argv+=['--artifact',str(out/'artifact.zip')]
 d.write(out/'started.json',dict(lane=lane,row=row,binary=build['binary'],build=d.desc(d.P/f"build-{row['leg']}.json"),argv=argv,started_unix=time.time()))
 code=d.run(label,argv)
 d.write(out/'receipt.json',dict(lane=lane,row=row,exit_code=code,command=d.desc(d.P/'commands'/label/'receipt.json'),artifacts=[d.desc(p) for p in sorted(out.rglob('*')) if p.is_file()],finished_unix=time.time()))
 receipts.append(d.desc(out/'receipt.json'));assert code==0,(label,code)
 assert (out/'report.json').is_file();print(label,'captured',flush=True)
d.write(d.P/f"capture-{lane}{'-'+leg if leg else ''}.json",dict(status='pass',receipts=receipts,plan=d.desc(d.P/'plan.json')))
