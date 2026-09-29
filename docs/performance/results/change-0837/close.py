"""Offline closure of the rejected compression candidate and retained retry fix."""
import re,sys
from pathlib import Path
import driver as d
import analyze as a

def main():
 origin=d.read(d.P/'origin.json')
 if not (d.P/'cleanup.json').exists():d.check()
 else:
  cleanup=d.read(d.P/'cleanup.json');assert cleanup['status']=='pass'
  assert not d.TARGET.exists() and not d.SCRATCH.exists()
  head=d.output(['git','rev-parse','HEAD'])
  if head!=origin['base']:
   assert d.output(['git','rev-parse','HEAD^'])==origin['base']
   seal=d.read(d.P/'seal.json')
   for name,h in seal['files'].items():assert d.sha(d.ROOT/name)==h
 for group in ['normative','unrelated']:
  for name,h in origin[group].items():assert d.sha(d.ROOT/name)==h
 build=d.read(d.P/'build-baseline.json');assert build==d.read(d.P/'build-baseline-v2.json')
 for value in build.values():a.descriptor(value)
 freeze=d.read(build['freeze']['path']);a.descriptor(freeze['driver'])
 for name,h in freeze['probe'].items():assert d.sha(d.P/name)==h
 baseline=freeze['source'];rejected=d.read(d.P/'rejected-source.json')['source'];quality=d.read(d.P/'final-v2-quality.json');final=quality['source'];assert final==d.source()
 final_delta=sorted(n for n in baseline.keys()|final.keys() if baseline.get(n)!=final.get(n))
 rejected_delta=sorted(n for n in baseline.keys()|rejected.keys() if baseline.get(n)!=rejected.get(n))
 assert final_delta==['crates/soapberry-zip/src/office.rs','crates/soapberry-zip/tests/streaming_interrupted_write.rs']
 assert rejected_delta==['crates/litchi-opc/src/package.rs','crates/litchi-opc/src/pkgwriter.rs',*final_delta]
 for stage,inventory,names in [('baseline',baseline,rejected_delta),('candidate',rejected,rejected_delta),('final',final,final_delta)]:
  for name in names:
   if name in inventory:assert d.sha(d.P/'sources'/stage/name)==inventory[name]
 for binding in quality['commands']:
  receipt=d.read(a.descriptor(binding));a.descriptor(receipt['log']);assert receipt['exit_code']==0
 rejection=d.read(d.P/'rejection.json');assert rejection['adoption']=='rejected-correctness'
 failed=d.read(a.descriptor(rejection['failed_gate']));assert failed['exit_code']==101
 failed_log=Path(a.descriptor(failed['log'])).read_text()
 assert 'publication_raw_copies_unselected_members_and_inverse_restores_exact_artifact ... FAILED' in failed_log
 assert 'source_backed_paragraph_copy.rs:183:5' in failed_log
 assert not any((d.P/n).exists() for n in ['build-candidate.json','qualification.json','native.json','capture-admission.json'])
 rg=d.read(d.P/'red-green.json');assert rg['status']=='pass'
 for binding,code in zip(rg['receipts'],[101,0]):
  r=d.read(a.descriptor(binding));a.descriptor(r['log']);assert r['exit_code']==code
 preflight=d.read(d.P/'baseline-preflight.json');assert preflight['status']=='pass' and preflight['reports']==6 and preflight['samples']==12
 pids=set();reports=[a.report(row,pids) for row in preflight['rows']]
 for previous,current in zip(reports,reports[1:]):assert previous['finished']<=current['started']
 import admit
 fixture=admit.fixture()
 commands=[];tests={}
 for path in sorted((d.P/'commands').glob('*/receipt.json')):
  r=d.read(path);log=Path(a.descriptor(r['log'])).read_text();assert r['finished_unix']>=r['started_unix']
  commands.append(dict(label=path.parent.name,exit_code=r['exit_code']))
  counts=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',log)
  if counts:tests[path.parent.name]=dict(zip(['passed','failed','ignored'],[sum(int(row[i]) for row in counts) for i in range(3)]))
 failed_commands=[r['label'] for r in commands if r['exit_code']]
 assert failed_commands==['baseline-interruption-regression','build-baseline','candidate-focused-applied','candidate-tests','final-tests'],failed_commands
 assert 'Checking litchi-opc' in (d.P/'commands/final-v2-check/output.log').read_text()
 assert 'Compiling litchi-opc' in (d.P/'commands/final-v2-tests/output.log').read_text()
 assert 'fresh_publication_uses_owned_staging_and_keeps_logical_payload' not in (d.P/'commands/final-v2-tests/output.log').read_text()
 assert 'interrupted_owned_payload_writes_resume_without_poisoning_or_double_counting ... ok' in (d.P/'commands/final-v2-tests/output.log').read_text()
 assert tests['final-v2-tests']['failed']==0 and tests['final-v2-tests']['passed']>1000
 result=dict(status='pass',final_source_delta=final_delta,rejected_source_delta=rejected_delta,baseline_preflight_reports=len(reports),baseline_preflight_samples=sum(r['samples'] for r in reports),comparative_reports=0,comparative_samples=0,fixture=fixture,commands=commands,tests=tests,performance_claim='none; candidate rejected before capture')
 if sys.argv[1:]==['--check']:assert d.read(d.P/'closure.json')==result
 else:assert not sys.argv[1:];d.write(d.P/'closure.json',result)
 print('closure PASS',len(commands),'commands;',tests['final-v2-tests'])
if __name__=='__main__':main()
