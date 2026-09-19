#!/usr/bin/env python3
"""Check supplemental lookup identities, outcomes, fixtures and loop means."""
import hashlib,json,statistics,subprocess,sys
from pathlib import Path
PACKET=Path(__file__).resolve().parent;ROOT=PACKET.parents[3]
assert len(sys.argv)==1 or sys.argv[1:]==['--initial']
P=PACKET/'initial-lookup' if len(sys.argv)>1 else PACKET
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
base=read(PACKET/'baseline.json');data={};outcomes={}
for phase in ['baseline','candidate']:
 d=P/'lookup'/phase;m=read(d/'manifest.json');assert m['repetitions']==250000 and m['samples_per_leg']==9 and m['warmups']==1000
 legs=['aa1','aa2'] if phase=='baseline' else ['a1','b1','b2','a2']
 expected={f'{leg}-{sample}.json' for leg in legs for sample in range(9)}
 assert len(m['commands'])==len(expected) and {c['output'] for c in m['commands']}==expected
 for n,h in m['raw_sha256'].items():assert sha(d/n)==h,n
 for p,b in m['builds'].items():
  assert b==read(P/(p+'-lookup-build.json')) and b['exit_code']==0
  for n,h in b['source_sha256'].items():
   actual=sha(ROOT/n) if p=='candidate' else hashlib.sha256(subprocess.check_output(['git','show',base['baseline_head']+':'+n],cwd=ROOT)).hexdigest()
   assert actual==h,n
  for n,h in b['probe_sha256'].items():
   source=ROOT/n
   if P!=PACKET:source=P/source.relative_to(PACKET)
   assert sha(source)==h,n
 for c in m['commands']:
  assert c['exit_code']==0
  leg=c['output'].split('-')[0];suffix='after' if leg in ['b1','b2'] else 'before'
  assert c['command']==['taskset','-c','12','/home/zhuhe/code/litchi-target-0690-'+suffix+'/release/cfb-lookup-probe-0690','--case','all','--format','json','--repetitions','250000','--warmups','1000']
  rows=read(d/c['output']);assert len(rows)==12
  for r in rows:
   assert r['repetitions']==250000 and r['elapsed_ns']>0
   assert r['successes']==(250000 if r['expected'] else 0) and r['errors']==(0 if r['expected'] else 250000)
   outcome={k:v for k,v in r.items() if k!='elapsed_ns'};name=r['case']
   if name in outcomes:assert outcomes[name]==outcome,name
   else:outcomes[name]=outcome
   data.setdefault(name,{}).setdefault(leg,[]).append(r['elapsed_ns']/250000)
rows=[]
for name,legs in data.items():
 assert set(legs)=={'aa1','aa2','a1','b1','b2','a2'} and all(len(v)==9 for v in legs.values())
 medians={k:statistics.median(v) for k,v in legs.items()}
 pct={k:(medians[a]/medians[b]-1)*100 for k,a,b in [('b1_a1','b1','a1'),('b2_a2','b2','a2'),('aa','aa2','aa1'),('abba_a','a2','a1')]}
 rows.append(dict(case=name,outcome=outcomes[name],ns_per_lookup=legs,median_ns=medians,percent=pct))
(P/'lookup-comparison.json').write_text(json.dumps(rows,indent=2)+'\n')
lines=['# Supplemental legacy CFB lookup loop means','','Nine processes per leg; 250,000 calls/process, 1,000 warmups. Fixture creation/parsing excluded. No individual-query tail claim.','','| Case | A1 ns | B1 ns | B1/A1 | B2/A2 | A/A | ABBA A |','|---|---:|---:|---:|---:|---:|---:|']
for r in rows:
 m=r['median_ns'];v=r['percent'];lines.append(f"| {r['case']} | {m['a1']:.2f} | {m['b1']:.2f} | {v['b1_a1']:+.2f}% | {v['b2_a2']:+.2f}% | {v['aa']:+.2f}% | {v['abba_a']:+.2f}% |")
(P/'lookup-summary.md').write_text('\n'.join(lines)+'\n');print('PASS 12 lookup groups x 6 legs x 9 process samples, fixture/source/probe bindings and all outcomes exact.')
