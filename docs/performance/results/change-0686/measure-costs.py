#!/usr/bin/env python3
"""Separate allocation and counted-I/O captures; never native latency."""
import hashlib,json,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
phase=sys.argv[1];assert phase in ['baseline','candidate']
suffix='before' if phase=='baseline' else 'after'
source=Path('/home/zhuhe/code/litchi-0686-before') if phase=='baseline' else ROOT
binary=Path('/home/zhuhe/code/litchi-target-0686-'+suffix)/'release'
out=P/'costs'/phase;out.mkdir(parents=True,exist_ok=True)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
commands=[]
manifest={'phase':phase,'head':subprocess.check_output(['git','rev-parse','HEAD'],cwd=source,text=True).strip(),'source_sha256':{str(f.relative_to(source)):sha(f) for owner in ['litchi-xls','litchi-cfb'] for f in sorted((source/'crates'/owner).rglob('*.rs'))},'binary_sha256':{n:sha(binary/n) for n in ['xls0686-alloc','xls-index-budget-probe-0684']},'probe_sha256':{str(f.relative_to(ROOT)):sha(f) for d in [P/'allocation-probe',P.parent/'change-0684/budget-probe'] for f in sorted(d.rglob('*')) if f.is_file()},'cases_sha256':sha(P/'cases.json'),'corpus':{c['path']:sha(ROOT/c['path']) for c in json.loads((P/'cases.json').read_text())}}
(out/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
def run(args,name):
 cmd=['taskset','-c','12',*map(str,args)];start=time.monotonic();r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True)
 (out/name).write_text(r.stdout)
 if r.stderr:(out/(name+'.stderr')).write_text(r.stderr)
 commands.append(dict(command=cmd,output=name,exit_code=r.returncode,seconds=time.monotonic()-start));assert r.returncode==0,(cmd,r.stderr)
for c in json.loads((P/'cases.json').read_text()):
 run([binary/'xls-index-budget-probe-0684','--input',c['path'],'--budget',c['budget'],'--worksheet',c['sheet'],'--row',c['row'],'--column',c['column'],'--queries',8],f"counts-{c['case']}.json")
 for mode in ['owned','file']:
  for op in ['q1','q2','q3','q8']:
   for repeat in range(3):run([binary/'xls0686-alloc',mode,op,c['path'],c['sheet'],c['row'],c['column'],c['budget']],f"alloc-{c['case']}-{mode}-{op}-{repeat}.json")
 print(phase,c['case'],flush=True)
(out/'commands.json').write_text(json.dumps(commands,indent=2)+'\n')
manifest['raw_sha256']={f.name:sha(f) for f in sorted(out.iterdir()) if f.is_file() and f.name!='manifest.json'}
(out/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
