"""Run only after the normal release build receipt completes successfully."""
import json
import run

assert json.loads((run.HERE/'baseline/build-normal.receipt.json').read_text())['exit_code'] == 0
binary = run.SCRATCH/'baseline-normal'
run.run('baseline','symbols',['nm','-C',str(binary)],binary)
lines = (run.HERE/'baseline/symbols.stdout').read_text().splitlines()
plan = run.plan_data()
candidates = {name:[line for line in lines if line.split(' ',2)[-1] == name] for name in plan['profile']['owner_candidates']}
owner = next(name for name,matches in candidates.items() if matches)
run.write(run.HERE/'symbol-observation.json',dict(owner=owner,candidates=candidates,binary_sha256=run.sha(binary),plan_sha256=run.sha(run.HERE/'plan.json'),command=['nm','-C',str(binary)],receipt_sha256=run.sha(run.HERE/'baseline/symbols.receipt.json'),stdout_sha256=run.sha(run.HERE/'baseline/symbols.stdout'),selection='First exact symbol present in the predeclared preference order.',selected_utc=run.now()))
run.capture('baseline','profile')
