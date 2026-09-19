#!/usr/bin/env python3
"""Whole-process perf counters/RSS at two query lengths; separate from latency."""
import hashlib,json,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1]
assert phase in ['baseline','candidate']
suffix='before' if phase=='baseline' else 'after';source=Path('/home/zhuhe/code/litchi-0686-before') if phase=='baseline' else ROOT
binary=Path('/home/zhuhe/code/litchi-target-0686-'+suffix)/'release/xls-index-retry-probe-0686'
out=P/'diagnostics'/phase;out.mkdir(parents=True,exist_ok=True)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
m=dict(phase=phase,binary_sha256=sha(binary),source_sha256={str(f.relative_to(source)):sha(f) for owner in ['litchi-xls','litchi-cfb'] for f in sorted((source/'crates'/owner).rglob('*.rs'))},probe_sha256={str(f.relative_to(ROOT)):sha(f) for f in sorted((P/'probe').rglob('*')) if f.is_file()},commands=[])
for case in ['54016-stored-1048576','54016-stored-524288','Simple-stored-2097152','synthetic-70000-default','synthetic-100000-default']:
 c=next(c for c in json.loads((P/'cases.json').read_text()) if c['case']==case)
 for repeat in range(3):
  for n in [10,1010]:
   stem=f'{case}-{repeat}-{n}'
   cmd=['/usr/bin/time','-v','-o',str(out/(stem+'.rss.txt')),'perf','stat','-x,','-o',str(out/(stem+'.csv')),'-e','cycles,instructions,branches,branch-misses,cache-misses,page-faults','--','taskset','-c','12',str(binary),'--input',c['path'],'--budget',str(c['budget']),'--mode','owned','--worksheet',str(c['sheet']),'--row',str(c['row']),'--column',str(c['column']),'--queries',str(n),'--samples','1','--warmups','1']
   r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
   j=json.loads(r.stdout);assert j['records'][0]['all_queries_agree'];(out/(stem+'.json')).write_text(json.dumps(j,separators=(',',':'))+'\n')
   m['commands'].append(dict(command=cmd,exit_code=r.returncode))
 print(phase,case,flush=True)
m['cases_sha256']=sha(P/'cases.json');m['corpus']={c['path']:sha(ROOT/c['path']) for c in json.loads((P/'cases.json').read_text())}
m['raw_sha256']={f.name:sha(f) for f in sorted(out.iterdir()) if f.is_file() and f.name!='manifest.json'}
(out/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
