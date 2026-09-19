#!/usr/bin/env python3
"""Supplemental legacy CFB lookup-only A/A and ABBA controls."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1];assert phase in ['baseline','candidate']
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def binary(p):return Path('/home/zhuhe/code/litchi-target-0690-'+('before' if p=='baseline' else 'after'))/'release/cfb-lookup-probe-0690'
out=P/'lookup'/phase;out.mkdir(parents=True,exist_ok=True)
phases=['baseline'] if phase=='baseline' else ['baseline','candidate'];builds={p:json.loads((P/(p+'-lookup-build.json')).read_text()) for p in phases}
m=dict(phase=phase,repetitions=250000,warmups=1000,samples_per_leg=9,builds=builds,commands=[])
for p in phases:assert sha(binary(p))==builds[p]['binary_sha256']
legs=[('aa1','baseline'),('aa2','baseline')] if phase=='baseline' else [('a1','baseline'),('b1','candidate'),('b2','candidate'),('a2','baseline')]
for leg,p in legs:
 for sample in range(9):
  cmd=['taskset','-c','12',str(binary(p)),'--case','all','--format','json','--repetitions','250000','--warmups','1000'];r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
  j=json.loads(r.stdout);assert len(j)==12 and all(v['repetitions']==250000 for v in j)
  name=f'{leg}-{sample}.json';(out/name).write_text(json.dumps(j,indent=2)+'\n');m['commands'].append(dict(command=cmd,output=name,exit_code=r.returncode))
 print(phase,leg,flush=True)
m['raw_sha256']={f.name:sha(f) for f in out.iterdir() if f.name!='manifest.json'};(out/'manifest.json').write_text(json.dumps(m,indent=2)+'\n')
