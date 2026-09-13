"""Serial candidate quality, matched ABBA captures and retained baseline repeat."""
import json
import subprocess
import sys
import run
import guard_run
run.freeze('candidate')
quality = json.loads((run.HERE/'quality-plan.json').read_text())['commands']
for name in ['quality-tests','quality-clippy']:
    run.run('candidate',name,quality[name])
for kind in ['normal','alloc']:
    run.build('candidate',kind)
for kind in ['normal','alloc']:
    guard_run.build('candidate',kind)
for repeat in [1,2]:
    run.capture('candidate','native',repeat)
    run.capture('candidate','alloc',repeat)
    guard_run.capture('candidate','native',repeat)
    guard_run.capture('candidate','alloc',repeat)
run.capture('baseline','native',2)
run.capture('baseline','alloc',2)
guard_run.capture('baseline','native',2)
guard_run.capture('baseline','alloc',2)
cap=run.HERE/'cap-boundary/cap_run.py'
for action in ['freeze','build']:
    subprocess.run([sys.executable,'-B',str(cap),'candidate',action],check=True)
for stage,repeat in [('candidate',1),('candidate',2),('baseline',2)]:
    subprocess.run([sys.executable,'-B',str(cap),stage,'capture',str(repeat)],check=True)
