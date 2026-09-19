#!/usr/bin/env python3
import json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase='candidate';build=json.loads((P/(phase+'-builds.json')).read_text())[0]
cmd=['perf','record','-F','997','--call-graph','dwarf','-o','/home/zhuhe/code/litchi-0687-profile/candidate.data','--','taskset','-c','12','/home/zhuhe/code/litchi-target-0687-after/release/xls0684-repeat','owned','test-data/poi/test-data/spreadsheet/54016.xls','0','0','0','2000000']
with (P/(phase+'-profile.tsv')).open('w') as o,(P/(phase+'-profile.stderr')).open('w') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=o,stderr=e)
with (P/(phase+'-profile-symbols.txt')).open('w') as o,(P/(phase+'-profile-report.stderr')).open('w') as e:s=subprocess.run(['perf','report','--stdio','--no-children','-i','/home/zhuhe/code/litchi-0687-profile/candidate.data'],cwd=ROOT,stdout=o,stderr=e)
assert r.returncode==s.returncode==0
(P/(phase+'-profile-command.json')).write_text(json.dumps(dict(command=cmd,exit_code=r.returncode,report_exit_code=s.returncode,binary_sha256=build['binary_sha256'],source_sha256=build['source_sha256']),indent=2)+'\n')
