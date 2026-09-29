"""Strict separate tiny-tail reduction; never merges confirmation into native data."""
import driver as d
import qualify as q
import reader as r
import analyze as a
import statistics,math

def analyze():
 plan=d.read(d.P/'tail-plan.json');rows=[]
 for row in plan['rows']:
  leg=row['leg'];folder=d.P/'tails'/f"tail-{row['block']:02}-tiny-{row['mode']}-{'leg-a' if leg=='before' else 'leg-b'}"
  item=q.receipt(folder);assert item['lane']=='tail' and item['row']==row
  qual=d.P/'runs'/f"qualification-v2-00-tiny-{row['mode']}-{'leg-a' if leg=='before' else 'leg-b'}"/'report.json'
  validated=r.validate_report(folder/'report.json','tiny',row['mode'],row['samples'],row['warmup'],expected=d.read(qual))
  raw=d.read(folder/'report.json');ns=sorted(x['elapsed_ns'] for x in raw['samples']);assert len(ns)==1000
  for percentile in [50,95,99]:assert validated[f'p{percentile}_ns']==ns[math.ceil(len(ns)*percentile/100)-1]
  rows.append(dict(**{k:v for k,v in item.items() if k!='folder'},folder=str(folder),**validated))
 assert {x['folder'] for x in rows}=={str(p) for p in (d.P/'tails').iterdir()}
 assert all(x['finished']<=y['started'] for x,y in zip(rows,rows[1:]))
 assert max(x['finished'] for x in d.read(d.P/'analysis.json')['native'])<=rows[0]['started']
 for p in (d.P/'commands').glob('*/receipt.json'):
  c=d.read(p)
  if c['argv'][0]=='cargo':assert all(x['finished']<=c['started_unix'] or x['started']>=c['finished_unix'] for x in rows)
 summary=[];flags=[]
 for mode in ['write','read']:
  for metric in ['p50_ns','p95_ns','p99_ns','rss_kib']:
   values={leg:[x[metric] for x in rows if x['row']['leg']==leg and x['row']['mode']==mode] for leg in ['before','after']}
   paired=a.bootstrap(values['before'],values['after'])
   entry=dict(shape='tiny',mode=mode,metric=metric,before_median=statistics.median(values['before']),after_median=statistics.median(values['after']),paired=paired);summary.append(entry)
   if paired['estimate']>1.05 and paired['ci95_low']>1:flags.append(entry)
 return dict(status='pass',reports=len(rows),samples=sum(x['row']['samples'] for x in rows),rows=rows,summary=summary,regression_flags=flags,plan=d.desc(d.P/'tail-plan.json'),analyzer=d.desc(d.P/'tail_analyze.py'))
if __name__=='__main__':
 value=analyze();assert value['reports']==24 and value['samples']==24000
 p=d.P/'tail-analysis.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Tail confirmation PASS',len(value['regression_flags']),'review flags')
