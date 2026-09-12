"""Run the frozen applicable quality commands serially with source-bound receipts."""
import argparse
import json
import re
from run import HERE, run, sha, write

parser = argparse.ArgumentParser()
parser.add_argument('stage', choices=['candidate', 'final'])
parser.add_argument('--only')
args = parser.parse_args()
plan = json.loads((HERE/'quality-plan.json').read_text())
checks = []
for name, command in plan['commands'].items():
    if args.only and name != args.only:
        continue
    run(args.stage, name, command)
    folder = HERE/args.stage
    receipt = folder/(name+'.receipt.json')
    text = (folder/(name+'.stdout')).read_text() + (folder/(name+'.stderr')).read_text()
    count = sum(int(n) for n in re.findall(r'test result: ok\. (\d+) passed;', text))
    checks.append(dict(name=name, receipt_sha256=sha(receipt), executed_tests=count))
write(HERE/(args.stage+'-quality-summary'+('-'+args.only if args.only else '')+'.json'), dict(
    status='pass', stage=args.stage, quality_plan_sha256=sha(HERE/'quality-plan.json'),
    checks=checks, successful_test_executions=sum(c['executed_tests'] for c in checks)))
