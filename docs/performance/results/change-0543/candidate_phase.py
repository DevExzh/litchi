import json
import run
import guard_run
run.freeze('candidate')
quality = json.loads((run.HERE/'quality-plan.json').read_text())['commands']
run.run('candidate','quality-tests',quality['quality-tests'])
run.build('candidate','normal')
run.build('candidate','alloc')
guard_run.build('candidate','normal')
guard_run.build('candidate','alloc')
for repeat in [1,2]:
    run.capture('candidate','native',repeat)
    run.capture('candidate','alloc',repeat)
    guard_run.capture('candidate','native',repeat)
    guard_run.capture('candidate','alloc',repeat)
run.capture('baseline','native',2)
run.capture('baseline','alloc',2)
guard_run.capture('baseline','native',2)
guard_run.capture('baseline','alloc',2)
