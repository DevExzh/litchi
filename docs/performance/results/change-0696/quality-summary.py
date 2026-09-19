#!/usr/bin/env python3
"""Retain gate totals and final log/source bindings."""
import hashlib,json,re
from pathlib import Path
P=Path(__file__).resolve().parent
rows=json.loads((P/'integration/results.json').read_text())
for row in rows:
 path=P/'integration'/(row['name']+'.log')
 row['log_sha256']=hashlib.sha256(path.read_bytes()).hexdigest()
 matches=re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;',path.read_text())
 row['test_totals']={name:sum(int(x[i]) for x in matches) for i,name in enumerate(['passed','failed','ignored'])}
(P/'quality-summary.json').write_text(json.dumps(rows,indent=2)+'\n')
for row in rows:print(row['name'],row['exit_code'],row['test_totals'])
