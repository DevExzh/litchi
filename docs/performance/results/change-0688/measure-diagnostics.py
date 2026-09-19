#!/usr/bin/env python3
"""Native-child RSS and extra-query hardware counters; separate from API latency."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1];assert phase in ['baseline','candidate']
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
binary=Path('/home/zhuhe/code/litchi-target-0688-'+('before' if phase=='baseline' else 'after'))/'release/xls0684-repeat'
out=P/'diagnostics'/phase;out.mkdir(parents=True,exist_ok=True)
m=dict(source_sha256=json.loads((P/(phase+'-builds.json')).read_text())[0]['source_sha256'],binary_sha256=sha(binary),commands=[])
for case in ['54016-stored-2097152','54016-late','Simple-stored-2097152','Plan1-stored-2097152']:
 c=next(c for c in json.loads((P/'cases.json').read_text()) if c['case']==case)
 for mode in ['owned','file']:
  for repeat in range(3):
   for n in [10,100010]:
    stem=f'{case}-{mode}-{repeat}-{n}';cmd=['perf','stat','-x,','-o',str(out/(stem+'.csv')),'-e','cycles,instructions,branches,branch-misses,cache-misses,page-faults','--','/usr/bin/time','-v','-o',str(out/(stem+'.rss.txt')),'taskset','-c','12',str(binary),mode,c['path'],str(c['sheet']),str(c['row']),str(c['column']),str(n)]
    r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr;(out/(stem+'.tsv')).write_text(r.stdout);m['commands'].append(dict(command=cmd,exit_code=r.returncode))
  print(phase,case,mode,flush=True)
m['cases_sha256']=sha(P/'cases.json');m['corpus']={c['path']:sha(ROOT/c['path']) for c in json.loads((P/'cases.json').read_text())};m['raw_sha256']={f.name:sha(f) for f in out.iterdir() if f.is_file() and f.name!='manifest.json'};(out/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
