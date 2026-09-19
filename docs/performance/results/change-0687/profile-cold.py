#!/usr/bin/env python3
"""Attribute the repeated fresh-owner open regression without changing probes."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];rows=[]
for phase,suffix in [('baseline','before'),('candidate','after')]:
 binary=Path('/home/zhuhe/code/litchi-target-0687-'+suffix)/'release/xls-index-retry-probe-0686';data='/home/zhuhe/code/litchi-0687-profile/'+phase+'-cold.data';cmd=['perf','record','-F','997','--call-graph','dwarf','-o',data,'--','taskset','-c','12',str(binary),'--input','test-data/poi/test-data/spreadsheet/54016.xls','--budget','0','--mode','owned','--worksheet','0','--row','0','--column','0','--queries','1','--samples','5000','--warmups','3']
 with (P/(phase+'-cold-profile.stderr')).open('w') as e:r=subprocess.run(cmd,cwd=ROOT,capture_output=False,stdout=subprocess.PIPE,stderr=e,text=True)
 assert r.returncode==0;j=json.loads(r.stdout);(P/(phase+'-cold-profile.json')).write_text(json.dumps(j,separators=(',',':'))+'\n')
 with (P/(phase+'-cold-profile-symbols.txt')).open('w') as o,(P/(phase+'-cold-report.stderr')).open('w') as e:s=subprocess.run(['perf','report','--stdio','--no-children','-i',data],stdout=o,stderr=e)
 assert s.returncode==0;rows.append(dict(phase=phase,command=cmd,exit_code=r.returncode,binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest()))
 print(phase,flush=True)
(P/'cold-profiles-manifest.json').write_text(json.dumps(rows,indent=2)+'\n')
