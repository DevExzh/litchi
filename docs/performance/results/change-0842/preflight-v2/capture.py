"""Serial fixed-schedule current-lock OPC compression captures."""
import driver as d
import sys,time
lane=sys.argv[1];leg=sys.argv[2] if len(sys.argv)>2 else None
assert lane in ['qualification','native'];assert (leg in ['before','after']) if lane=='qualification' else leg is None
plan=d.read(d.P/'plan.json');rows=plan[lane]
if lane=='qualification':rows=[r for r in rows if r['leg']==leg]
else:
 admission=d.read(d.P/'admission.json')
 assert all(d.desc(x['path'])==x for x in admission['files'])
 assert d.source()==d.read(d.P/'source-after.json')['source']
receipts=[]
for i,row in enumerate(rows):
 stage=row['leg'];build=d.read(d.P/f'build-{stage}.json');binary=build['binary'];assert d.desc(binary['path'])==binary
 frozen=d.read(build['freeze']['path']);assert all(d.sha(d.P/n)==h for n,h in frozen['probe'].items())
 label=f"{lane}-{row.get('block',0):02}-{row['shape']}-{row['mode']}-{'leg-a' if stage=='before' else 'leg-b'}"
 out=d.P/'runs'/label;out.mkdir(parents=True,exist_ok=False)
 report=out/'report.json'
 child=[binary['path'],'--mode',row['mode'],'--shape',row['shape'],'--samples',str(row['samples']),'--warmup',str(row['warmup']),'--output',str(report)]
 if lane=='qualification':child+=['--artifact',str(out/'artifact.zip')]
 argv=['taskset','-c',str(plan['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(out/'time.txt'),*child]
 d.write(out/'started.json',dict(lane=lane,row=row,binary=binary,build=d.desc(d.P/f'build-{stage}.json'),argv=argv,started_unix=time.time()))
 code=d.run(label,argv)
 d.write(out/'receipt.json',dict(lane=lane,row=row,exit_code=code,command=d.desc(d.P/'commands'/label/'receipt.json'),artifacts=[d.desc(p) for p in sorted(out.rglob('*')) if p.is_file()],finished_unix=time.time()))
 receipts.append(d.desc(out/'receipt.json'))
 assert code==0,(label,code)
 assert report.is_file()
 print(label,'captured',flush=True)
d.write(d.P/f"capture-{lane}{('-'+leg) if leg else ''}.json",dict(status='pass',plan=d.desc(d.P/'plan.json'),receipts=receipts))
