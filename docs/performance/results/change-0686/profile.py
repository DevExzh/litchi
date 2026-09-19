#!/usr/bin/env python3
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1];assert phase in ['baseline','candidate']
binary=Path('/home/zhuhe/code/litchi-target-0686-'+('before' if phase=='baseline' else 'after'))/'release/xls-index-retry-probe-0686'
d=Path('/home/zhuhe/code/litchi-0686-profile');d.mkdir(exist_ok=True)
cmd=['perf','record','-F','997','--call-graph','dwarf','-o',str(d/(phase+'.data')),'--','taskset','-c','12',str(binary),'--input','test-data/poi/test-data/spreadsheet/54016.xls','--budget','1048576','--mode','owned','--worksheet','0','--row','0','--column','0','--queries','10000','--samples','1','--warmups','1']
r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
(P/(phase+'-profile.stderr')).write_text(r.stderr)
j=json.loads(r.stdout);assert j['records'][0]['all_queries_agree']
(P/(phase+'-profile.json')).write_text(json.dumps(j,separators=(',',':'))+'\n')
report=['perf','report','--stdio','--no-children','--percent-limit','1','-i',str(d/(phase+'.data'))]
s=subprocess.run(report,capture_output=True,text=True);assert s.returncode==0,s.stderr
(P/(phase+'-profile-symbols.txt')).write_text(s.stdout)
(P/(phase+'-profile-report.stderr')).write_text(s.stderr)
(P/(phase+'-profile-command.json')).write_text(json.dumps(dict(command=cmd,report_command=report,exit_code=r.returncode,report_exit_code=s.returncode,binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest()),indent=2)+'\n')
