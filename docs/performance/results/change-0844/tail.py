"""Separate predeclared confirmation of every original p99 review point."""
import driver as d
import qualify as q
import reader as r
import analyze as original
import sys,time,statistics,random,math

def folder(row):return d.P/'tail-runs'/f"{row['block']:02}-{row['shape']}-{row['parts']:04}-{row['mode']}-{row['leg']}"
def argv(row):
 build=d.read(d.P/f"build-{row['leg']}.json");out=folder(row)
 return ['taskset','-c','12','/usr/bin/time','-f','%M %R %F','-o',str(out/'time.txt'),build['binary']['path'],'--shape',row['shape'],'--parts',str(row['parts']),'--mode',row['mode'],'--samples',str(row['samples']),'--warmup',str(row['warmup']),'--output',str(out/'report.json')]

def validate(row):
 out=folder(row);end=d.read(out/'receipt.json');start=d.read(out/'started.json')
 assert start['row']==end['row']==row and end['exit_code']==0
 command=d.read(q.bind(end['command']));cstart=d.read(d.Path(end['command']['path']).parent/'started.json');q.bind(command['log'])
 assert command['argv']==start['argv']==cstart['argv']==argv(row) and command['exit_code']==0
 assert start['started_unix']<=command['started_unix']==cstart['started_unix']<=command['finished_unix']<=end['finished_unix']
 for x in end['artifacts']:q.bind(x)
 assert {d.Path(x['path']).name for x in end['artifacts']}=={p.name for p in out.iterdir() if p.is_file() and p.name!='receipt.json'}
 build=d.read(q.bind(start['build']));assert build==d.read(d.P/f"build-{row['leg']}.json") and build['binary']==start['binary']
 q.bind(start['plan']);assert start['plan']==d.desc(d.P/'tail-plan.json')
 expected=d.read(q.folder_for('qualification',dict(row,block=0))/'report.json')
 result=r.validate_report(out/'report.json',row['parts'],row['mode'],row['samples'],row['warmup'],shape=row['shape'],expected=expected)
 counts=(out/'time.txt').read_text().split();assert len(counts)==3 and all(x.isdigit() for x in counts) and int(counts[0])>0
 return dict(row=row,folder=str(out),started=command['started_unix'],finished=command['finished_unix'],rss_kib=int(counts[0]),minor_faults=int(counts[1]),major_faults=int(counts[2]),**result)

def analyze():
 plan=d.read(d.P/'tail-plan.json');q.bind(plan['original_analysis']);q.bind(plan['runner']);q.bind(plan['admission'])
 for x in d.read(d.P/'admission.json')['files']:q.bind(x)
 rows=[validate(row) for row in plan['rows']]
 assert {x['folder'] for x in rows}=={str(p) for p in (d.P/'tail-runs').iterdir()}
 ordered=sorted(rows,key=lambda x:x['started']);assert [x['row'] for x in ordered]==plan['rows']
 assert all(a['finished']<=b['started'] for a,b in zip(ordered,ordered[1:]))
 main=d.read(d.P/'analysis.json');assert min(x['started'] for x in rows)>max(x['finished'] for x in main['native']+main['heap'])
 for p in (d.P/'commands').glob('*/receipt.json'):
  c=d.read(p)
  if c['argv'][0]=='cargo':assert all(x['finished']<=c['started_unix'] or x['started']>=c['finished_unix'] for x in rows)
 summary=[];flags=[];points=[];spreads=[]
 for case in plan['cases']:
  legs={leg:[x for x in rows if x['row']['leg']==leg and all(x['row'][k]==v for k,v in case.items())] for leg in ['before','after']}
  for metric in ['p50_ns','p95_ns','p99_ns','mean_ns','rss_kib','minor_faults','major_faults']:
   b=[x[metric] for x in legs['before']];a=[x[metric] for x in legs['after']];stats=original.paired(b,a)
   item=dict(**case,metric=metric,before_median=statistics.median(b),after_median=statistics.median(a),paired=stats);summary.append(item)
   if metric in ['p50_ns','p95_ns','p99_ns','rss_kib'] and stats['estimate'] is not None and stats['estimate']>1.05 and stats['ci95_low']>1:flags.append(item)
   if metric in ['p95_ns','p99_ns'] and stats['estimate'] is not None and stats['estimate']>1.05:points.append(item)
  for leg,items in legs.items():
   values=[x['p50_ns'] for x in items]
   if max(values)/min(values)>1.2:spreads.append(dict(**case,leg=leg,values=values))
 return dict(status='pass',reports=len(rows),samples=sum(x['row']['samples'] for x in rows),rows=rows,summary=summary,regression_flags=flags,tail_points=points,spread_flags=spreads,plan=d.desc(d.P/'tail-plan.json'),runner=d.desc(d.P/'tail.py'))

if __name__=='__main__':
 mode=sys.argv[1]
 if mode=='plan':
  d.check();a=d.read(d.P/'analysis.json');cases=[]
  for x in a['tail_review']:
   c={k:x[k] for k in ['shape','parts','mode']}
   if c not in cases:cases.append(c)
  assert len(cases)==6 and all(x['metric']=='p99_ns' for x in a['tail_review'])
  rows=[]
  for block in range(6):
   for case in cases if block%2==0 else cases[::-1]:
    for leg in ['before','after'] if (block+cases.index(case))%2==0 else ['after','before']:
     rows.append(dict(case,block=block,leg=leg,samples=1000,warmup=3))
  d.write(d.P/'tail-plan.json',dict(cases=cases,rows=rows,original_analysis=d.desc(d.P/'analysis.json'),runner=d.desc(d.P/'tail.py'),admission=d.desc(d.P/'admission.json'),scope='Separate post-hoc confirmation of all6originalp99pointflags; no originalrun replaced or pooled.',review_threshold='p50/p95/p99/RSS paired ratio>1.05 andCI low>1; retain alltailpoint ratios>1.05; blockp50max/min>1.20.',samples=72000,reports=72))
 elif mode=='capture':
  plan=d.read(d.P/'tail-plan.json');q.bind(plan['runner']);q.bind(plan['original_analysis']);q.bind(plan['admission'])
  for x in d.read(d.P/'admission.json')['files']:q.bind(x)
  for row in plan['rows']:
   out=folder(row);out.mkdir(parents=True,exist_ok=False);build=d.read(d.P/f"build-{row['leg']}.json");q.bind(build['binary']);command=argv(row)
   d.write(out/'started.json',dict(row=row,argv=command,build=d.desc(d.P/f"build-{row['leg']}.json"),binary=build['binary'],plan=d.desc(d.P/'tail-plan.json'),started_unix=time.time()))
   label='tail-'+out.name;code=d.run(label,command)
   d.write(out/'receipt.json',dict(row=row,exit_code=code,command=d.desc(d.P/'commands'/label/'receipt.json'),artifacts=[d.desc(p) for p in sorted(out.rglob('*')) if p.is_file()],finished_unix=time.time()))
   assert code==0
  d.write(d.P/'capture-tail.json',dict(status='pass',plan=d.desc(d.P/'tail-plan.json'),receipts=[d.desc(folder(row)/'receipt.json') for row in plan['rows']]))
 elif mode=='analyze':
  value=analyze();assert value['reports']==72 and value['samples']==72000
  p=d.P/'tail-analysis.json'
  if p.exists():assert d.read(p)==value
  else:d.write(p,value)
  print('Tail confirmation PASS',value['reports'],value['samples'],'flags',len(value['regression_flags']),'points',len(value['tail_points']),'spreads',len(value['spread_flags']))
 else:raise AssertionError(mode)
