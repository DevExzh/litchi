#!/usr/bin/env python3
"""Matched selected-query profiles; source and binary identities bound per phase."""
import json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1];assert phase in ['baseline','candidate'];build=json.loads((P/(phase+'-builds.json')).read_text())[0]
suffix='before' if phase=='baseline' else 'after';data='/home/zhuhe/code/litchi-0689-profile/'+phase+'.data';Path(data).parent.mkdir(exist_ok=True)
cmd=['perf','record','-F','997','--call-graph','dwarf','-o',data,'--','taskset','-c','12','/home/zhuhe/code/litchi-target-0689-'+suffix+'/release/xls0684-repeat','owned','test-data/poi/test-data/spreadsheet/54016.xls','0','0','0','2000000']
with (P/(phase+'-profile.tsv')).open('w') as o,(P/(phase+'-profile.stderr')).open('w') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=o,stderr=e)
with (P/(phase+'-profile-symbols.txt')).open('w') as o,(P/(phase+'-profile-report.stderr')).open('w') as e:s=subprocess.run(['perf','report','--stdio','--no-children','-i',data],cwd=ROOT,stdout=o,stderr=e)
assert r.returncode==s.returncode==0
(P/(phase+'-profile-command.json')).write_text(json.dumps(dict(command=cmd,exit_code=r.returncode,report_exit_code=s.returncode,binary_sha256=build['binary_sha256'],source_sha256=build['source_sha256']),indent=2)+'\n')
