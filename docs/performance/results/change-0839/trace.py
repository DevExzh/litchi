"""Untimed current-binary private-thread census; no trace timings admitted."""
import driver as d
import capture as c
import readers as r
import re,shutil
assert shutil.which('strace')
cases=[dict(route='parts',shape='small',state=state,task_floor=0,workers=width) for state,width in [('fresh',4),('primed',4),('primed',32)]]
rows=[]
for i,case in enumerate(cases):
 for leg in ['before','after']:
  dest=d.P/'traces'/f'{i}-{leg}';dest.mkdir(parents=True,exist_ok=False)
  binary=c.binary(leg,'native');report=dest/'report.json';log=dest/'clone.log'
  args=['strace','-f','-e','trace=clone,clone3','-o',str(log),binary,*c.args(case,1,0,report)]
  assert d.run(f'trace-{i}-{leg}',['taskset','-c',c.AFF,*args])==0
  checked=r.validate_report(report,case,1,0)
  text=log.read_text();created=[line for line in text.splitlines() if ('clone(' in line or 'clone3(' in line or '<... clone resumed>' in line or '<... clone3 resumed>' in line) and re.search(r'= [1-9][0-9]*\s*$',line)]
  expected=case['workers']*(2 if case['state']=='primed' and leg=='before' else 1)
  assert len(created)==expected,(i,leg,len(created),expected)
  rows.append(dict(case=case,leg=leg,successful_thread_creations=len(created),expected=expected,log=d.desc(log),report=d.desc(report),binary=d.desc(binary),verification_ok=checked['verification_ok']))
d.write(d.P/'trace-analysis.json',dict(status='pass',rows=rows,scope='whole-child successful clone/clone3 calls; setup preload and measured read distinguished by fresh/primed control, trace timing excluded'))
print('trace census PASS')
