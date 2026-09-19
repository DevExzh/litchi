#!/usr/bin/env python3
"""Verify supplemental graph/error parity and retain all individual comparisons."""
import hashlib,json,math,random,statistics
from pathlib import Path
P=Path(__file__).resolve().parent
rng=random.Random(693)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def q(v,f):return sorted(v)[max(0,math.ceil(len(v)*f)-1)]
def stats(v):
 medians=[statistics.median(rng.choices(v,k=len(v))) for _ in range(2000)]
 return dict(p50_ns=statistics.median(v),mean_ns=statistics.mean(v),p95_ns=q(v,.95),p99_ns=q(v,.99),median_bootstrap_95_ns=[q(medians,.025),q(medians,.975)])
records=[]
for name in ['baseline','compare']:
 path=P/f'refusal-runs-{name}.json'
 if path.exists():records+=json.loads(path.read_text())
assert records
summary=[];bindings={};values={}
for record in records:
 assert record['exit_code']==0
 path=P/record['output']
 assert sha(path)==record['output_sha256']
 assert sha(path.with_suffix('.stderr'))==record['stderr_sha256']
 build=json.loads((P/f"build-refusal-{record['phase']}.json").read_text())
 assert record['binary_sha256']==build['binary_sha256']
 cases={};current=None;header={}
 for line in path.read_text().splitlines():
  fields=line.split('\t')
  if fields[0]=='case':
   current=fields[1];assert current not in cases
   cases[current]={'metadata':{},'samples':[]}
  elif fields[0].isdigit():
   assert current is not None and int(fields[0])==len(cases[current]['samples'])
   cases[current]['samples'].append(int(fields[1]))
  elif fields[0] in ['sample_ns']:pass
  elif fields[0]=='all_iterations_passed':assert fields[1]=='true';header[fields[0]]=fields[1]
  elif current is None:header[fields[0]]=fields[1]
  else:cases[current]['metadata'][fields[0]]=fields[1:]
 assert header['samples']=='100' and header['warmups']=='5' and header['all_iterations_passed']=='true'
 assert len(cases)==int(header['cases'])==10
 for case,data in cases.items():
  assert len(data['samples'])==100
  if case in bindings:assert bindings[case]==data['metadata'],case
  else:bindings[case]=data['metadata']
  row=dict(case=case,leg=record['leg'],phase=record['phase'],samples=100,**stats(data['samples']))
  summary.append(row);values[case,record['leg']]=row
pairs=[('a1','a0','aa'),('a3','a2','aa'),('b0','a2','ab'),('b1','a3','ab')]
comparisons=[]
for case in bindings:
 for numerator,denominator,kind in pairs:
  if (case,numerator) not in values:continue
  a=values[case,denominator];b=values[case,numerator]
  delta={key:(b[key]/a[key]-1)*100 for key in ['p50_ns','mean_ns','p95_ns','p99_ns']}
  comparisons.append(dict(case=case,pair=f'{numerator}/{denominator}',kind=kind,delta_pct=delta,delta_ns={key:b[key]-a[key] for key in delta}))
for name,data in [('refusal-summary.json',summary),('refusal-bindings.json',bindings),('refusal-comparisons.json',comparisons),('refusal-review-triggers.json',[dict(case=r['case'],pair=r['pair'],metric=k,delta_pct=v,delta_ns=r['delta_ns'][k]) for r in comparisons if r['kind']=='ab' for k,v in r['delta_pct'].items() if v>5])]:
 (P/name).write_text(json.dumps(data,indent=2)+'\n')
print('verified',len(records),'native refusal matrix receipts;',len(summary),'case legs')
if any(row['leg']=='b1' for row in summary):
 lines=['| Case | Baseline medians (µs) | Candidate medians (µs) | Paired change |','| --- | ---: | ---: | ---: |']
 for case in bindings:
  v={r['leg']:r['p50_ns']/1000 for r in summary if r['case']==case}
  d=[r['delta_pct']['p50_ns'] for r in comparisons if r['case']==case and r['kind']=='ab']
  lines.append(f"| {case} | {v['a2']:.2f} / {v['a3']:.2f} | {v['b0']:.2f} / {v['b1']:.2f} | {d[0]:+.2f}% / {d[1]:+.2f}% |")
 (P/'refusal-tables.md').write_text('\n'.join(lines)+'\n')
