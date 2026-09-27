"""Summarize retained Rust test results, with no test execution."""
import json,re,sys
from pathlib import Path
p=Path(__file__).resolve().parent
q=json.loads((p/'quality.json').read_text())
rows=[]
for r in q['rows']:
 if r['command'][1]!='test':continue
 text=Path(r['log']['path']).read_text()
 for match in re.finditer(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored; (\d+) measured; (\d+) filtered out',text):
  status,*counts=match.groups();rows.append(dict(zip(['status','passed','failed','ignored','measured','filtered'],[status,*map(int,counts)])))
assert rows and all(r['status']=='ok' and r['failed']==0 for r in rows)
result={'suites':len(rows),**{k:sum(r[k] for r in rows) for k in ['passed','failed','ignored']},'results':rows}
out=p/'quality-summary.json'
if '--check' in sys.argv:assert json.loads(out.read_text())==result
else:assert not out.exists();out.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
print(json.dumps({k:v for k,v in result.items() if k!='results'}))
