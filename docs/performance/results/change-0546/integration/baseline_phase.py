"""Serial baseline source freeze, quality prerequisites and first measurements."""
import json
import subprocess
import sys
import run
import guard_run
run.freeze('baseline')
quality = json.loads((run.HERE/'quality-plan.json').read_text())['commands']
for name in ['quality-cap-clippy','quality-tests']:
    run.run('baseline',name,quality[name])
for kind in ['normal','alloc']:
    run.build('baseline',kind)
for kind in ['normal','alloc']:
    guard_run.build('baseline',kind)
run.capture('baseline','native',1)
run.capture('baseline','alloc',1)
guard_run.capture('baseline','native',1)
guard_run.capture('baseline','alloc',1)
cap = run.HERE/'cap-boundary/cap_run.py'
for action in ['freeze','build','capture']:
    subprocess.run([sys.executable,'-B',str(cap),'baseline',action]+(['1'] if action=='capture' else []),check=True)
