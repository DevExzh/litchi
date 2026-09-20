#!/usr/bin/env python3
"""Derive validation totals from the retained successful command logs."""
import hashlib
import json
from pathlib import Path
import re
import sys

P = Path(__file__).resolve().parent
inputs = {}
lanes = {}
for lane in ['docx', 'harness', 'evidence']:
    receipt = P / f'quality-{lane}.json'
    inputs[receipt.name] = hashlib.sha256(receipt.read_bytes()).hexdigest()
    rows = json.loads(receipt.read_text())
    assert len(rows) == {'docx': 5, 'harness': 1, 'evidence': 6}[lane]
    summaries = []
    for row in rows:
        assert row['exit_code'] == 0
        log = P / row['log']
        digest = hashlib.sha256(log.read_bytes()).hexdigest()
        assert digest == row['log_sha256']
        inputs[log.name] = digest
        matches = re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;', log.read_text())
        summary = {'name': row['name'], 'exit_code': 0, 'command': row['command']}
        if matches:
            summary['test_suites'] = len(matches)
            summary.update({key: sum(int(match[index]) for match in matches)
                            for index, key in enumerate(['passed', 'failed', 'ignored'])})
        summaries.append(summary)
    lanes[lane] = summaries
result = {'status': 'pass', 'inputs': inputs, 'lanes': lanes,
          'script_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest()}
encoded = json.dumps(result, indent=2) + '\n'
output = P / 'quality-summary.json'
if sys.argv[1:] == ['--check']:
    assert output.read_text() == encoded
else:
    assert not sys.argv[1:]
    output.write_text(encoded)
print('PASS: quality receipts and log totals agree')
