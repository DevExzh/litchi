"""Freeze and run conditional diagnostics after all pilot gates pass."""
import json
import subprocess
import run
import eager_run

run.check_source('candidate')
native = json.loads((run.HERE / 'comparison.json').read_text())
alloc = json.loads((run.HERE / 'allocation-analysis.json').read_text())
guard = json.loads((run.HERE / 'guard-analysis.json').read_text())
assert native['native_admission']['passed']
assert alloc['planning_gate_passed']
assert guard['admission_status'] == 'pass'
cap = json.loads((run.HERE/'cap-boundary/cap-analysis.json').read_text())
assert cap['comparison']['admission_passed']
eager_run.create_plan()
owner = run.plan_data()['profile']['owner']
symbols = {}
for stage in ['baseline', 'candidate']:
    binary = run.SCRATCH / (stage + '-normal')
    command = ['nm', '-C', '--defined-only', str(binary)]
    output = subprocess.check_output(command, text=True)
    lines = [line for line in output.splitlines() if owner in line]
    assert lines, (stage, owner)
    symbols[stage] = {
        'command': command, 'binary_sha256': run.sha(binary),
        'source_manifest_sha256': run.sha(run.HERE / stage / 'source-manifest.json'),
        'matching_lines': lines,
    }
run.write(run.HERE / 'symbol-observation.json', {
    'created_utc': run.now(), 'owner': owner,
    'plan_sha256': run.sha(run.HERE / 'plan.json'), 'stages': symbols,
})
files = [
    'plan.json', 'run.py', 'eager-plan.json', 'eager_run.py', 'analyze_eager.py',
    'analyze_planning.py', 'analyze_hardware.py', 'conditional_phase.py',
    'comparison.json', 'allocation-analysis.json', 'guard-analysis.json',
    'symbol-observation.json', 'cap-boundary/cap-analysis.json',
]
run.write(run.HERE / 'conditional-frozen-inputs.json', {
    'created_utc': run.now(), 'stage': 'before first conditional capture',
    'files': {name: run.sha(run.HERE / name) for name in files},
})
order = [('baseline', 1), ('candidate', 1), ('candidate', 2), ('baseline', 2)]
for stage, repeat in order:
    run.capture(stage, 'profile', repeat)
for stage, repeat in order:
    run.capture(stage, 'hardware', repeat)
for stage, repeat in order:
    eager_run.capture(stage, repeat)
