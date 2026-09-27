"""Validate archived gate receipts, measured samples and final packet seal."""
from pathlib import Path
import hashlib,json,re
import analyze
P=Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def validate():
 q=json.loads((P/'quality.json').read_text());assert len(q['rows'])==8
 assert sha(P/q['source_file'])==q['source_sha256']
 results=[]
 for row in q['rows']:
  assert row['exit']==0 and sha(P/row['log'])==row['sha256']
  matches=re.findall(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored',(P/row['log']).read_text())
  assert all(m[0]=='ok' and m[2]=='0' for m in matches)
  results.append({'command':row['command'],'passed':sum(int(m[1]) for m in matches),'ignored':sum(int(m[3]) for m in matches)})
 analysis=analyze.analyze();assert analysis==json.loads((P/'analysis.json').read_text())
 allocation=json.loads((P/'allocation-analysis.json').read_text())
 raw=json.loads((P/'allocations-0/runs.json').read_text());assert len(raw)==4
 for row,receipt in zip(allocation['rows'],raw):
  assert row['case']==receipt['case'] and row['leg']==receipt['leg']
  assert receipt['exit']==receipt['decode_exit']==row['decode_exit']==0
  for key in ['stdout','stderr','capture','summary']:
   assert sha(P/'allocations-0'/receipt[key])==receipt[key+'_sha256']
  assert sha(P/'allocations-0'/row['histogram'])==row['histogram_sha256']
  assert sha(P/'allocations-0'/row['log'])==row['log_sha256']
  pairs=[list(map(int,l.split())) for l in (P/'allocations-0'/row['histogram']).read_text().splitlines()]
  assert sum(n for _,n in pairs)==row['allocation_calls']
  assert sum(size*n for size,n in pairs)==row['allocated_bytes_from_histogram']
  summary=(P/'allocations-0'/receipt['summary']).read_text()
  assert int(re.search(r'calls to allocation functions: (\d+)',summary)[1])==row['allocation_calls']
  data=[json.loads(l) for l in (P/'allocations-0'/receipt['stdout']).read_text().splitlines() if l.startswith('{')]
  assert len(data)==1 and data[0]['output_sha256']==analysis['cases'][row['case']]['output_sha256']
 property=json.loads((P/'property-final-1024.json').read_text())
 assert property['exit']==0 and property['source_unchanged'] and sha(P/property['log'])==property['sha256']
 seal=P/'seal.json'
 if seal.exists():
  files={str(f.relative_to(P)):sha(f) for f in sorted(P.rglob('*')) if f.is_file() and f!=seal}
  assert files==json.loads(seal.read_text())['files'],'packet seal mismatch'
 return {'gates':results,'cases':analysis['cases']}
if __name__=='__main__':print(json.dumps(validate(),indent=2))
