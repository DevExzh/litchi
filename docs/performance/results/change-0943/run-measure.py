#!/usr/bin/env python3
import argparse,datetime,hashlib,json,os,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--name',required=True);p.add_argument('--action',choices=['qualify','capture'],required=True);p.add_argument('--manifest',type=Path,required=True);p.add_argument('--qualification-root',type=Path);a=p.parse_args()
r=Path('/home/zhuhe/litchi-performance-0943');out=r/'runs'/a.name;out.mkdir(parents=True,exist_ok=False)
def desc(p):return dict(path=str(p.resolve()),bytes=p.stat().st_size,sha256=hashlib.sha256(p.read_bytes()).hexdigest())
def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()
runner=r/'tooling/comparison-runner.py';cmd=['/usr/bin/python3','-B',str(runner),'--'+a.action,'--manifest',str(a.manifest)]
if a.action=='qualify':
 assert a.qualification_root is not None
 cmd+=['--qualification-root',str(a.qualification_root)]
else:cmd+=['--lane','all']
env=dict(PATH='/usr/bin:/bin',LC_ALL='C',LANG='C',TZ='UTC')
env['LITCHI_PERF_0943_'+('QUALIFY' if a.action=='qualify' else 'CAPTURE')+'_APPROVED']='1'
record=dict(schema='litchi.performance.0943.outer-process.v1',argv=cmd,cwd=str(r),environment=env,manifest=desc(a.manifest),runner=desc(runner),wrapper=desc(Path(__file__)),started_utc=now())
(out/'start.json').write_text(json.dumps(record,indent=2)+'\n')
with (out/'stdout.log').open('xb') as stdout,(out/'stderr.log').open('xb') as stderr:
 process=subprocess.Popen(cmd,cwd=r,env=env,stdout=stdout,stderr=stderr,start_new_session=True)
 code=process.wait()
record.update(exit_code=code,ended_utc=now(),stdout=desc(out/'stdout.log'),stderr=desc(out/'stderr.log'),terminal_written_before_formal_reduction=True)
(out/'terminal.json').write_text(json.dumps(record,indent=2)+'\n')
print(json.dumps(dict(name=a.name,exit_code=code,terminal=str(out/'terminal.json'))))
raise SystemExit(code)
