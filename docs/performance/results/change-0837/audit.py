"""Offline source/build custody and independent arithmetic audit (run with -B)."""
import copy,json,math,re,statistics,sys
from pathlib import Path
from unittest.mock import patch
import driver as d
import analyze as a

def main():
 builds={s:d.read(d.P/f'build-{s}.json') for s in ['baseline','candidate']}
 freezes={s:d.read(a.descriptor(b['freeze'])) for s,b in builds.items()}
 for s,b in builds.items():
  a.descriptor(b['binary']);r=d.read(a.descriptor(b['receipt']));a.descriptor(r['log']);assert r['exit_code']==0
  for name,h in freezes[s]['probe'].items():assert d.sha(d.P/name)==h
  a.descriptor(freezes[s]['driver'])
 assert freezes['baseline']['probe']==freezes['candidate']['probe']
 base=freezes['baseline']['source'];candidate=freezes['candidate']['source']
 delta=sorted(n for n in base.keys()|candidate.keys() if base.get(n)!=candidate.get(n))
 expected=['crates/litchi-opc/src/package.rs','crates/litchi-opc/src/pkgwriter.rs','crates/soapberry-zip/src/office.rs','crates/soapberry-zip/tests/streaming_interrupted_write.rs']
 assert delta==expected,delta
 for stage,inventory in [('baseline',base),('candidate',candidate)]:
  for name in delta:
   if name in inventory:assert d.sha(d.P/'sources'/stage/name)==inventory[name]
 quality=d.read(d.P/'quality.json');assert quality['status']=='pass' and quality['source']==candidate
 for binding in quality['commands']:
  receipt=d.read(a.descriptor(binding));a.descriptor(receipt['log']);assert receipt['exit_code']==0
 for group in ['normative','unrelated']:
  for name,h in d.read(d.P/'origin.json')[group].items():assert d.sha(d.ROOT/name)==h
 admission=d.read(d.P/'capture-admission.json');assert admission['status']=='pass'
 for name,binding in admission['inputs'].items():assert d.desc(d.P/name)==binding
 assert d.read(d.P/'qualification-analysis.json')==a.analysis('qualification')
 native=d.read(d.P/'native-analysis.json');assert native==a.analysis('native')
 rows=d.read(d.P/'native.json')['rows'];raw=[(row,d.read(row['report']['path'])) for row in rows]
 independent=[]
 for summary in native['summaries']:
  case=summary['case'];pair={}
  for metric in ['p50_ns','rss_bytes']:
   values=[]
   for block in range(6):
    by_stage={row['stage']:report for row,report in raw if row['block']==block and row['case']==case}
    def measure(report):
     if metric=='rss_bytes':return report['vm_hwm_bytes']
     values=sorted(report['elapsed_ns']);assert len(values)==20
     return (values[9]+values[10])//2
    values.append(measure(by_stage['candidate'])/measure(by_stage['baseline']))
   assert values==summary['paired'][metric]['ratios']
   assert statistics.median(values)==summary['paired'][metric]['median']
   pair[metric]=statistics.median(values)
  independent.append(dict(case=case,ratios=pair))
 mutations=[]
 first=rows[0];original=d.read(first['report']['path']);original_read=d.read
 changes=[('sample-count',lambda r:r.update(samples=19)),('case',lambda r:r.update(case='bogus')),('duration-zero',lambda r:r['elapsed_ns'].__setitem__(0,0)),('duration-copy',lambda r:r['sample_elapsed_ns'].__setitem__(0,1)),('warmup',lambda r:r.update(warmup_elapsed_ns=[])),('rss',lambda r:r.update(vm_hwm_bytes=0)),('output-oracle',lambda r:r.update(expected_output_sha256='0'*64)),('validation',lambda r:r.update(all_outputs_validated=False)),('fixture',lambda r:r.update(fixture_raw_sha256='0'*64)),('member-duplicate',lambda r:r['members'].append(r['members'][0]))]
 for name,mutate in changes:
  damaged=copy.deepcopy(original);mutate(damaged)
  def mocked_read(path):return damaged if Path(path)==Path(first['report']['path']) else original_read(path)
  with patch.object(d,'read',side_effect=mocked_read):
   try:a.report(first,set())
   except AssertionError:mutations.append(name)
   else:raise AssertionError('mutation accepted: '+name)
 try:a.report(first,{original['pid']})
 except AssertionError:mutations.append('duplicate-pid')
 else:raise AssertionError('duplicate PID accepted')
 commands=[]
 for path in sorted((d.P/'commands').glob('*/receipt.json')):
  r=d.read(path);a.descriptor(r['log']);assert r['finished_unix']>=r['started_unix']
  commands.append(dict(label=path.parent.name,exit_code=r['exit_code']))
 result=dict(status='pass',source_delta=delta,independent=independent,negative_checks=mutations,commands=commands,reports=len(rows),samples=sum(len(r['elapsed_ns']) for _,r in raw))
 if sys.argv[1:]==['--check']:assert d.read(d.P/'audit.json')==result
 else:assert not sys.argv[1:];d.write(d.P/'audit.json',result)
 print('audit PASS',len(rows),'reports',len(mutations),'negative checks')
if __name__=='__main__':main()
