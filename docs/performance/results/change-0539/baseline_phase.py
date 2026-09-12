import json
import run
run.freeze('baseline')
quality = json.loads((run.HERE/'quality-plan.json').read_text())['commands']
run.run('baseline','quality-tests',quality['quality-tests'])
run.build('baseline','normal')
run.build('baseline','alloc')
run.capture('baseline','native',1)
run.capture('baseline','alloc',1)
