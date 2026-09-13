import json
import run
import sys
stage = sys.argv[1]
assert stage in ('baseline', 'candidate')
run.check_source(stage)
for name,command in json.loads((run.HERE/'quality-plan.json').read_text())['commands'].items():
    run.run(stage,'final-'+name,command)
