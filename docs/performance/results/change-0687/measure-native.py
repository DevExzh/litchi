#!/usr/bin/env python3
"""Native fresh-owner A/A then A/B/B/A; no allocation or ReadAt instrumentation."""
import hashlib,json,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1]
assert phase in ['baseline','candidate']
suffix='before' if phase=='baseline' else 'after'
source=ROOT
before=Path('/home/zhuhe/code/litchi-target-0687-before/release/xls-index-retry-probe-0686')
after=Path('/home/zhuhe/code/litchi-target-0687-after/release/xls-index-retry-probe-0686')
out=P/'native'/phase;out.mkdir(parents=True,exist_ok=True)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
commands=[]
m=dict(phase=phase,samples=100,warmups=3,queries=8,cpu=12,source_sha256={str(f.relative_to(source)):sha(f) for owner in ['litchi-cfb','litchi-xls'] for f in sorted((source/'crates'/owner).rglob('*.rs'))},probe_sha256={str(f.relative_to(ROOT)):sha(f) for f in sorted((P.parent/'change-0686/probe').rglob('*')) if f.is_file()},cases_sha256=sha(P/'cases.json'),corpus={c['path']:sha(ROOT/c['path']) for c in json.loads((P/'cases.json').read_text())},before_binary_sha256=sha(before))
if phase=='candidate':m['after_binary_sha256']=sha(after)
(out/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
for c in json.loads((P/'cases.json').read_text()):
 for mode in ['owned','file']:
  legs=[('aa1',before),('aa2',before)] if phase=='baseline' else [('a1',before),('b1',after),('b2',after),('a2',before)]
  for leg,binary in legs:
   cmd=['taskset','-c','12',str(binary),'--input',c['path'],'--budget',str(c['budget']),'--mode',mode,'--worksheet',str(c['sheet']),'--row',str(c['row']),'--column',str(c['column']),'--queries','8','--warmups','3','--samples','100']
   start=time.monotonic();r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
   name=f"{leg}-{c['case']}-{mode}.json"
   # Compact JSON preserves raw numeric values and avoids millions of indented lines.
   j=json.loads(r.stdout);(out/name).write_text(json.dumps(j,separators=(',',':'))+'\n')
   assert len(j['records'])==100 and all(x['all_queries_agree'] for x in j['records'])
   commands.append(dict(command=cmd,output=name,exit_code=r.returncode,seconds=time.monotonic()-start))
 print(phase,c['case'],flush=True)
(out/'commands.json').write_text(json.dumps(commands,indent=2)+'\n')
m['raw_sha256']={f.name:sha(f) for f in sorted(out.iterdir()) if f.is_file() and f.name!='manifest.json'}
(out/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
