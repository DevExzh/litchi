"""Additive fixed tiny-tail confirmation; original native evidence is untouched."""
import driver as d
import time
plan=d.read(d.P/'tail-plan.json');admission=d.read(d.P/'tail-admission.json')
assert all(d.desc(x['path'])==x for x in admission['files'])
assert all(d.desc(x['path'])==x for x in d.read(d.P/'admission.json')['files'])
assert d.source()==d.read(d.P/'source-final.json')['source']
receipts=[]
for row in plan['rows']:
 leg=row['leg'];build=d.read(d.P/f'build-{leg}.json');assert d.desc(build['binary']['path'])==build['binary']
 label=f"tail-{row['block']:02}-tiny-{row['mode']}-{'leg-a' if leg=='before' else 'leg-b'}"
 out=d.P/'tails'/label;out.mkdir(parents=True,exist_ok=False)
 argv=['taskset','-c',str(plan['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(out/'time.txt'),build['binary']['path'],'--mode',row['mode'],'--shape','tiny','--samples',str(row['samples']),'--warmup',str(row['warmup']),'--output',str(out/'report.json')]
 d.write(out/'started.json',dict(lane='tail',row=row,binary=build['binary'],build=d.desc(d.P/f'build-{leg}.json'),argv=argv,started_unix=time.time()))
 code=d.run(label,argv)
 d.write(out/'receipt.json',dict(lane='tail',row=row,exit_code=code,command=d.desc(d.P/'commands'/label/'receipt.json'),artifacts=[d.desc(p) for p in sorted(out.rglob('*')) if p.is_file()],finished_unix=time.time()))
 receipts.append(d.desc(out/'receipt.json'));assert code==0
 d.check()
d.write(d.P/'capture-tail.json',dict(status='pass',receipts=receipts,plan=d.desc(d.P/'tail-plan.json')))
