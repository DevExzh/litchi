"""Summarize exactly the successful source-bound quality receipts."""
import json
import re
from checks import COMMANDS
from run import HERE, sha, write

candidate_manifest = HERE / 'candidate/source-manifest.json'
rows = []
for name, _ in COMMANDS:
    stage = 'preflight-3' if name in ('xlsx-tests', 'clippy') else 'candidate'
    folder = HERE / stage
    assert (folder / 'source-manifest.json').read_bytes() == candidate_manifest.read_bytes()
    filename = 'check-' + name + '.receipt.json'
    receipt = json.loads((folder / filename).read_text())
    assert receipt['exit_code'] == 0
    assert receipt['source_manifest_sha256'] == sha(candidate_manifest)
    log = (folder / ('check-' + name + '.stdout')).read_text()
    count = sum(map(int, re.findall(r'test result: ok\. (\d+) passed;', log)))
    rows.append(dict(name=filename, stage=stage, receipt_sha256=sha(folder / filename),
                     exit_code=0, executed_tests=count))
write(HERE / 'quality-summary.json', dict(
    status='pass', stage='candidate', checks=rows,
    executed_tests=sum(row['executed_tests'] for row in rows),
    source_manifest_sha256=sha(candidate_manifest),
    alias_scope='XLSX tests and Clippy reuse preflight-3 only on exact candidate manifest equality.',
))
print('Quality gates:', len(rows), 'passing test executions:', sum(row['executed_tests'] for row in rows))
