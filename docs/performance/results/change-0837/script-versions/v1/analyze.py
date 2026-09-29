"""Strict report custody, paired block statistics, and opened-byte controls."""
import math,random,statistics,sys
from pathlib import Path
import driver as d

def descriptor(value):
 path=Path(value['path'])
 if path.exists():assert d.desc(path)==value,path
 else:
  cleanup=d.read(d.P/'cleanup.json')
  assert cleanup['status']=='pass' and value in cleanup['artifacts'],path
 return path

def quantile(values,q):return sorted(values)[math.ceil(len(values)*q)-1]
def midpoint(values):
 ordered=sorted(values);n=len(ordered)
 return ordered[n//2] if n%2 else (ordered[n//2-1]+ordered[n//2])//2
def report(row,pids):
 path=descriptor(row['report']);r=d.read(path)
 assert r['schema']=='fresh-publication-v1' and r['command']=='run'
 assert r['case']==row['case'] and r['samples']==row['samples'] and r['warmup']==row['warmup']
 values=r['elapsed_ns'];warm=r['warmup_elapsed_ns']
 assert len(values)==row['samples'] and len(warm)==row['warmup']
 assert values==r['sample_elapsed_ns'] and warm+values==r['all_elapsed_ns']
 assert all(type(v) is int and v>0 for v in values+warm)
 assert type(r['pid']) is int and r['pid']>0 and r['pid'] not in pids;pids.add(r['pid'])
 assert type(r['vm_hwm_bytes']) is int and r['vm_hwm_bytes']>0
 assert r['all_outputs_validated'] is True and r['output_stable_within_process'] is True
 assert r['output_sha256']==r['expected_output_sha256']
 for key in ['output_sha256','input_semantic_sha256','member_canonical_sha256']:
  assert len(r[key])==64 and all(c in '0123456789abcdef' for c in r[key])
 assert type(r['output_bytes']) is int and r['output_bytes']>0
 names=[m['name'] for m in r['members']];assert names==sorted(set(names))
 assert '[Content_Types].xml' in names and '_rels/.rels' in names
 assert all(type(m['bytes']) is int and m['bytes']>=0 and len(m['sha256'])==64 for m in r['members'])
 if row['case'].endswith('-edit'):
  assert r['fixture_raw_sha256']==d.sha(d.P/'fixtures/fixture.xml.zip') and r['selected_xml_part']=='/custom/data.xml'
 else:assert r['fixture_raw_sha256'] is None and r['selected_xml_part'] is None
 receipt=d.read(descriptor(row['receipt']));descriptor(receipt['log'])
 binary=d.read(d.P/f"build-{row['stage']}.json")['binary'];assert binary==row['binary'];descriptor(binary)
 assert receipt['exit_code']==0 and receipt['argv']==['taskset','-c','12',binary['path'],'run',row['case'],str(row['samples']),str(row['warmup']),str(d.P/'fixtures'),str(path)]
 assert receipt['finished_unix']>=receipt['started_unix']
 return dict(**row,p50_ns=midpoint(values),p95_ns=quantile(values,.95),p99_ns=quantile(values,.99),min_ns=min(values),max_ns=max(values),mean_ns=sum(values)/len(values),rss_bytes=r['vm_hwm_bytes'],output_bytes=r['output_bytes'],output_sha256=r['output_sha256'],semantic_sha256=r['input_semantic_sha256'],member_sha256=r['member_canonical_sha256'],members=r['members'],started=receipt['started_unix'],finished=receipt['finished_unix'])

def validate(kind):
 plan=d.read(d.P/'plan.json');manifest=d.read(d.P/(kind+'.json'));assert manifest['status']=='pass'
 expected=[]
 if kind=='qualification':
  for case in plan['cases']:
   for stage in ['baseline','candidate']:expected.append(dict(case=case,stage=stage,samples=1,warmup=0))
 else:
  for block in range(plan['blocks']):
   cases=plan['cases'] if block%2==0 else list(reversed(plan['cases']))
   stages=['baseline','candidate'] if block%2==0 else ['candidate','baseline']
   for case in cases:
    for stage in stages:expected.append(dict(block=block,case=case,stage=stage,samples=plan['samples'],warmup=plan['warmup']))
 assert manifest['reports']==len(expected)==len(manifest['rows'])
 assert manifest['samples']==sum(r['samples'] for r in expected)
 pids=set();rows=[]
 for row,wanted in zip(manifest['rows'],expected):
  assert all(row.get(k)==v for k,v in wanted.items())
  rows.append(report(row,pids))
 for previous,current in zip(rows,rows[1:]):assert previous['finished']<=current['started']
 oracle=[]
 for case in plan['cases']:
  group=[r for r in rows if r['case']==case]
  first=group[0]
  assert all(r['semantic_sha256']==first['semantic_sha256'] and r['member_sha256']==first['member_sha256'] and r['members']==first['members'] for r in group),case
  for stage in ['baseline','candidate']:
   selected=[r for r in group if r['stage']==stage]
   assert len({r['output_sha256'] for r in selected})==len({r['output_bytes'] for r in selected})==1
  if case.endswith('-edit'):assert len({r['output_sha256'] for r in group})==1,case
  base=next(r for r in group if r['stage']=='baseline');candidate=next(r for r in group if r['stage']=='candidate')
  oracle.append(dict(case=case,semantic_sha256=first['semantic_sha256'],member_sha256=first['member_sha256'],baseline_output_sha256=base['output_sha256'],candidate_output_sha256=candidate['output_sha256'],baseline_bytes=base['output_bytes'],candidate_bytes=candidate['output_bytes'],size_ratio=candidate['output_bytes']/base['output_bytes']))
 return rows,oracle

def analysis(kind):
 rows,oracle=validate(kind)
 out=dict(status='pass',kind=kind,reports=len(rows),samples=sum(r['samples'] for r in rows),oracles=oracle,rows=rows)
 if kind=='qualification':return out
 plan=d.read(d.P/'plan.json');summaries=[];flags=[]
 for case in plan['cases']:
  by_stage={stage:[r for r in rows if r['case']==case and r['stage']==stage] for stage in ['baseline','candidate']}
  entry=dict(case=case,stages={},paired={})
  for stage,group in by_stage.items():
   entry['stages'][stage]={key:statistics.median(r[key] for r in group) for key in ['p50_ns','p95_ns','p99_ns','rss_bytes']}
   for metric in ['p50_ns','rss_bytes']:
    spread=max(r[metric] for r in group)/min(r[metric] for r in group)
    if spread>1.2:flags.append(dict(case=case,stage=stage,metric=metric,kind='spread',ratio=spread))
  for metric in ['p50_ns','rss_bytes']:
   ratios=[]
   for block in range(plan['blocks']):
    base=next(r for r in by_stage['baseline'] if r['block']==block)
    candidate=next(r for r in by_stage['candidate'] if r['block']==block)
    ratios.append(candidate[metric]/base[metric])
   rng=random.Random(plan['statistics']['seed']);boot=sorted(statistics.median(rng.choices(ratios,k=len(ratios))) for _ in range(plan['statistics']['bootstrap']))
   median=statistics.median(ratios);low,high=plan['statistics']['endpoints']
   entry['paired'][metric]=dict(ratios=ratios,median=median,bootstrap_95=[boot[low],boot[high]])
   if median>1.05:flags.append(dict(case=case,metric=metric,kind='regression',ratio=median))
  summaries.append(entry)
 for r in oracle:
  if r['size_ratio']>1.01:flags.append(dict(case=r['case'],kind='size_growth',ratio=r['size_ratio']))
 out.update(summaries=summaries,flags=flags,performance_claim='none pending disposition; matched prepared-graph publication only')
 return out

if __name__=='__main__':
 kind=sys.argv[1];assert kind in ['qualification','native']
 result=analysis(kind);path=d.P/(kind+'-analysis.json')
 if sys.argv[2:]==['--check']:assert d.read(path)==result
 else:assert not sys.argv[2:];d.write(path,result)
 print(kind,'analysis PASS',result['reports'],result['samples'],flush=True)
