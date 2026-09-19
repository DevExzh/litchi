#!/usr/bin/env python3
"""Long same-owner controls; repetitions are not fresh-owner timing samples."""
import hashlib,json,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1];assert phase in ['baseline','candidate']
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def binary(p):return Path('/home/zhuhe/code/litchi-target-0690-'+('before' if p=='baseline' else 'after'))/'release/xls0684-repeat'
out=P/'repeat'/phase;out.mkdir(parents=True,exist_ok=True)
builds=json.loads((P/(phase+'-builds.json')).read_text());m=dict(source_sha256=builds[0]['source_sha256'],binary_sha256={p:sha(binary(p)) for p in (['baseline'] if phase=='baseline' else ['baseline','candidate'])},probe_sha256={str(f.relative_to(ROOT)):sha(f) for f in (P.parent/'change-0684/repeat-probe').rglob('*') if f.is_file()},cases_sha256=sha(P/'cases.json'),commands=[])
for case in ['54016-stored-2097152','54016-late','54016-missing-1048576','Plan1-stored-2097152','Simple-stored-2097152','45365-first']:
 c=next(c for c in json.loads((P/'cases.json').read_text()) if c['case']==case)
 for mode in ['owned','file']:
  legs=[('aa1','baseline'),('aa2','baseline')] if phase=='baseline' else [('a1','baseline'),('b1','candidate'),('b2','candidate'),('a2','baseline')]
  for leg,p in legs:
   for sample in range(9):
    cmd=['taskset','-c','12',str(binary(p)),mode,c['path'],str(c['sheet']),str(c['row']),str(c['column']),'50000'];r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
    name=f'{case}-{mode}-{leg}-{sample}.tsv';(out/name).write_text(r.stdout);m['commands'].append(dict(command=cmd,output=name,exit_code=r.returncode))
  print(phase,case,mode,flush=True)
m['corpus']={c['path']:sha(ROOT/c['path']) for c in json.loads((P/'cases.json').read_text())};m['raw_sha256']={f.name:sha(f) for f in out.iterdir() if f.is_file() and f.name!='manifest.json'};(out/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
