import json
import run
import guard_run
run.freeze('baseline')
quality = json.loads((run.HERE/'quality-plan.json').read_text())['commands']
run.run('baseline','quality-guard-clippy',quality['quality-guard-clippy'])
run.run('baseline','quality-tests',quality['quality-tests'])
run.build('baseline','normal')
run.build('baseline','alloc')
guard_run.build('baseline','normal')
guard_run.build('baseline','alloc')
run.capture('baseline','native',1)
run.capture('baseline','alloc',1)
guard_run.capture('baseline','native',1)
guard_run.capture('baseline','alloc',1)
