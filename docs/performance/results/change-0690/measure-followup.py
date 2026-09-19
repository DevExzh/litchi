#!/usr/bin/env python3
"""Supplementary controls for the observed first-query drift and file tails."""
import gzip,hashlib,json,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];out=P/'followup';out.mkdir(exist_ok=True)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
builds={phase:json.loads((P/(phase+'-builds.json')).read_text()) for phase in ['baseline','candidate']}
binaries={phase:Path('/home/zhuhe/code/litchi-target-0690-'+suffix+'/release/xls-index-retry-probe-0686') for phase,suffix in [('baseline','before'),('candidate','after')]}
for phase,binary in binaries.items():assert sha(binary)==next(b['binary_sha256'] for b in builds[phase] if b['binary']=='xls-index-retry-probe-0686')
sources=builds['candidate'][0]['source_sha256'];assert all(sha(ROOT/n)==h for n,h in sources.items())
selected=[('Simple-missing-2097152','owned'),('Simple-stored-2097152','file'),('45365-first','file'),('45365-late','file')]
cases={c['case']:c for c in json.loads((P/'cases.json').read_text())};commands=[]
m=dict(reason='Primary Simple missing q3 has +10 ns paired cost, Simple/45365 file query tails flag, and 45365-late file B1 workflow has large one-leg drift; preserve originals and add controls.',samples=1000,warmups=10,queries=8,cpu=12,current_source_sha256=sources,binary_sha256={phase:sha(b) for phase,b in binaries.items()},cases_sha256=sha(P/'cases.json'),corpus={cases[name]['path']:sha(ROOT/cases[name]['path']) for name,mode in selected},probe_sha256={str(f.relative_to(ROOT)):sha(f) for f in (P.parent/'change-0686/probe').rglob('*') if f.is_file()},compression='Lossless compact JSON in deterministic gzip, outside probe timers.')
for name,mode in selected:
 c=cases[name]
 for leg,phase in [('aa1','baseline'),('aa2','baseline'),('a1','baseline'),('b1','candidate'),('b2','candidate'),('a2','baseline')]:
  cmd=['taskset','-c','12',str(binaries[phase]),'--input',c['path'],'--budget',str(c['budget']),'--mode',mode,'--worksheet',str(c['sheet']),'--row',str(c['row']),'--column',str(c['column']),'--queries','8','--warmups','10','--samples','1000'];start=time.monotonic();r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
  j=json.loads(r.stdout);assert len(j['records'])==1000 and all(x['all_queries_agree'] for x in j['records'])
  file=f'{name}-{mode}-{leg}.json.gz';(out/file).write_bytes(gzip.compress(json.dumps(j,separators=(',',':')).encode(),mtime=0));commands.append(dict(command=cmd,output=file,exit_code=r.returncode,seconds=time.monotonic()-start))
 print(name,mode,flush=True)
m['commands']=commands;m['raw_sha256']={f.name:sha(f) for f in out.iterdir() if f.is_file() and f.name!='manifest.json'};assert all(sha(ROOT/n)==h for n,h in sources.items());(out/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
