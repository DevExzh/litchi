#!/usr/bin/env python3
import argparse,datetime,hashlib,json,os,subprocess
from pathlib import Path
p=argparse.ArgumentParser()
p.add_argument('--name',required=True)
p.add_argument('--source',required=True,type=Path)
p.add_argument('--source-receipt',required=True,type=Path)
p.add_argument('--rustdoc',action='store_true')
p.add_argument('command',nargs=argparse.REMAINDER)
a=p.parse_args()
cmd=a.command[1:] if a.command[:1]==['--'] else a.command
assert cmd
root=Path('/home/zhuhe/litchi-performance-0943/checks')/a.name
root.mkdir(parents=True,exist_ok=False)
def desc(path):
    return dict(path=str(path.resolve()),bytes=path.stat().st_size,sha256=hashlib.sha256(path.read_bytes()).hexdigest())
def now(): return datetime.datetime.now(datetime.timezone.utc).isoformat()
env={k:os.environ[k] for k in ('HOME','CARGO_HOME','RUSTUP_HOME','PATH','LC_ALL','LANG','TZ') if k in os.environ}
env.update(CARGO_TERM_COLOR='never',CARGO_INCREMENTAL='0',CARGO_TARGET_DIR='/home/zhuhe/litchi-performance-0943/target-quality-v1')
if a.rustdoc: env['RUSTDOCFLAGS']='-D warnings'
record=dict(schema='litchi.performance.0943.quality.v1',argv=cmd,cwd=str(a.source.resolve()),source_receipt=desc(a.source_receipt),environment=env,started_utc=now())
(root/'start.json').write_text(json.dumps(record,indent=2)+'\n')
with (root/'stdout.log').open('xb') as out,(root/'stderr.log').open('xb') as err:
    result=subprocess.run(cmd,cwd=a.source,env=env,stdout=out,stderr=err)
record.update(exit_code=result.returncode,ended_utc=now(),stdout=desc(root/'stdout.log'),stderr=desc(root/'stderr.log'))
(root/'terminal.json').write_text(json.dumps(record,indent=2)+'\n')
print(json.dumps(dict(name=a.name,exit_code=result.returncode,receipt=str(root/'terminal.json'))))
raise SystemExit(result.returncode)
