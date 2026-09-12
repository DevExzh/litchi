"""Run the original eight quality gates on the restored production source."""
import json
import re
from run import HERE, run, sha, write
plan=json.loads((HERE/'quality-plan.json').read_text())
checks=[]
for name, command in plan['commands'].items():
    run('restored', name, command)
    folder=HERE/'restored'
    receipt=folder/(name+'.receipt.json')
    text=(folder/(name+'.stdout')).read_text()+(folder/(name+'.stderr')).read_text()
    checks.append(dict(name=name,receipt_sha256=sha(receipt),executed_tests=sum(int(n) for n in re.findall(r'test result: ok\. (\d+) passed;',text))))
write(HERE/'restored-quality-summary.json',dict(status='pass',stage='restored',quality_plan_sha256=sha(HERE/'quality-plan.json'),checks=checks,successful_test_executions=sum(c['executed_tests'] for c in checks)))
