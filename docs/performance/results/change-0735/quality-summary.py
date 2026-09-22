#!/usr/bin/env python3
"""Summarize passed/failed/ignored test counts from exact retained gate logs."""
import json,re
from pathlib import Path
P=Path(__file__).resolve().parent
q=json.loads((P/'quality.json').read_text());result={}
def counts(path):
 rows=re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;',path.read_text());assert rows
 return dict(zip(['passed','failed','ignored'],[sum(int(r[i]) for r in rows) for i in range(3)]))
for label,index in [('all_targets',2),('doctests',4)]:result[label]=counts(P/q['runs'][index]['output'])
for v in ['baseline','candidate']:
 b=json.loads((P/f'{v}-build.json').read_text());quality=json.loads((P/b['quality']).read_text());result[v+'_probe']=counts(P/quality['runs'][1]['output'])
(P/'quality-summary.json').write_text(json.dumps(result,indent=2)+'\n');print(result)
