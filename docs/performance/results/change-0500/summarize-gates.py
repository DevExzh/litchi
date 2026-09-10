#!/usr/bin/env python3
"""Summarize recorded current gate receipts without adding overlapping suites."""
from pathlib import Path
import hashlib,json,re
HERE=Path(__file__).resolve().parent
result={}
for name in json.loads((HERE/'gate-commands.json').read_text()):
    receipt=HERE/'checks'/(name+'.json'); log=receipt.with_suffix('.log')
    r=json.loads(receipt.read_text()); data=log.read_bytes()
    assert r['exit_code']==0 and hashlib.sha256(data).hexdigest()==r['log_sha256'],name
    counts=[tuple(map(int,m)) for m in re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;',data.decode())]
    result[name]={'exit_code':r['exit_code'],'receipt_sha256':hashlib.sha256(receipt.read_bytes()).hexdigest(),'log_sha256':r['log_sha256'],'suite_summaries':len(counts),'passed':sum(v[0] for v in counts),'failed':sum(v[1] for v in counts),'ignored':sum(v[2] for v in counts)}
(HERE/'gate-summary.json').write_text(json.dumps({'scope':'Per-gate counts; focused and broad suites overlap and are not additive. Compile-only checks have no executed test count.','gates':result},indent=2)+'\n')
print(json.dumps({name:result[name] for name in ['docx-default','docx-features','doctests']}))
