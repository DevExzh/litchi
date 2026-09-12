import json
import run
run.check_source('baseline')
for name,command in json.loads((run.HERE/'quality-plan.json').read_text())['commands'].items():
    run.run('baseline','final-'+name,command)
