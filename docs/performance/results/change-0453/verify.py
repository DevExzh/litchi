#!/usr/bin/env python3
"""Portable verification of source epochs, real commands, output and allocator evidence."""
import argparse,hashlib,importlib.util,json,re
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def module(name):
 s=importlib.util.spec_from_file_location(name,ROOT/(name+'.py'));m=importlib.util.module_from_spec(s);s.loader.exec_module(m);return m
d=module('derive');oracle=module('verify-report')
def sha(raw):return hashlib.sha256(raw).hexdigest()
def member(name):
 p=Path(name);d.require(not p.is_absolute() and '..' not in p.parts,'safe artifact path');p=(ROOT/p).resolve();d.require(p.is_relative_to(ROOT.resolve()),'artifact escape');return p
def artifact(r):
 member(r['path']);raw=d.raw(r['path']);d.require(len(raw)==r['bytes'] and sha(raw)==r['sha256'],'artifact '+r['path']);return raw
def manifest(r):
 member(r['path']);raw=d.raw(r['path']);v=json.loads(raw);d.require(sha(raw)==r['sha256'] and len(v)==r['files'],'source manifest');return v
def check(sealed=False,cleanup=False):
 files=d.load('source-files.json');d.require(files==['crates/litchi-pptx/src/presentation/source_cross_copy.rs', 'crates/litchi-pptx/tests/source_backed_cross_copy.rs', 'tools/perf-baseline/src/pptx_provider_lifecycle.rs', 'tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs', 'tools/perf-baseline/src/filesystem.rs', 'tools/perf-baseline/src/docx_story_hyperlinks.rs', 'tools/perf-baseline/src/parallel_metrics.rs', 'tools/perf-baseline/src/xls_numeric.rs', 'tools/perf-baseline/src/lib.rs', 'tools/perf-baseline/src/operation_metrics.rs', 'tools/perf-baseline/src/bin/xls_source_attribution.rs', 'tools/perf-baseline/src/bin/xlsb_crud.rs', 'tools/perf-baseline/src/bin/docx_source_selected_paragraph.rs', 'tools/perf-baseline/src/bin/xlsx_cell_values_abba.rs'],'source inventory')
 builds={n:d.load(n+'-build.json') for n in ['baseline','candidate']};base=manifest(builds['baseline']['source_manifest']);final=manifest(builds['candidate']['source_manifest'])
 for n in files:
  for label,m in [('before',base),('after',final)]:d.require(sha(d.raw('candidate/'+label+'-'+Path(n).name+'.txt'))==m[n],'exact source '+label+' '+n)
 parent=d.load('parent-sources.json');d.require(set(parent['sources'])==set(files),'parent source inventory')
 for n,h in parent['sources'].items():d.require(sha(d.raw('candidate/parent-'+Path(n).name+'.txt'))==h,'parent source '+n)
 for n in files[:2]:d.require(base[n]==parent['sources'][n],'baseline original production')
 restored=d.load('checks/baseline-restoration.json');d.require(restored['status']=='pass' and restored['build_exit_code']==0 and restored['restored']=={n:final[n] for n in files},'candidate restoration')
 production=files[:2];d.require(production==['crates/litchi-pptx/src/presentation/source_cross_copy.rs','crates/litchi-pptx/tests/source_backed_cross_copy.rs'],'production source set')
 d.require({k:v for k,v in base.items() if k not in production}=={k:v for k,v in final.items() if k not in production},'identical harness and other source for both builds')
 required=['baseline-build-r3','candidate-build','final-strict-r2','final-harness-strict-r2','final-pptx-r2','final-opc-r2','final-harness-r2','final-doc-r2','final-workspace-r2','final-format-r2','final-boundaries-r2','candidate-fallback-proof-r4','baseline-alloc-preflight-r4','baseline-alloc-oracle-r2','fuzz-lock','fuzz-build','fuzz-smoke','final-fuzz-strict','final-fuzz-format']
 protocol=d.load('protocol.json');required+=['confirmation-'+str(i) for i in range(8)];required+=['pilot-'+str(i) for i in protocol['pilot_lanes']]+['formal-'+str(i) for i in range(24)]
 for tag in required:d.require((ROOT/'checks'/f'{tag}.json').is_file(),'required '+tag)
 failures={'baseline-alloc-preflight','baseline-alloc-preflight-r2','baseline-alloc-oracle','final-harness-strict','candidate-fallback-proof-r2','candidate-fallback-proof-r3','harness-lint-repair','harness-lint-repair-r2','harness-lint-repair-r3','harness-lint-repair-r4'}
 for p in (ROOT/'checks').glob('*.json'):
  r=d.load(str(p.relative_to(ROOT)))
  if 'source_before' not in r:continue
  d.require(r['change']==453 and r['driver_sha256']==sha(d.raw('check.py')),'check driver')
  d.require(r['source_unchanged'] and r['source_before']==r['source_after'],'frozen command source')
  manifest(r['source_after']);passed=p.stem not in failures
  d.require(r['status']==('pass' if passed else 'failed') and (r['exit_code']==0)==passed,'check status '+p.stem)
  log=artifact(r['log']).decode()
  if 'test' in r['argv']:
   rows=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',log)
   for i,key in enumerate(['passed_tests','failed_tests','ignored_tests']):d.require(r[key]==sum(int(x[i]) for x in rows),'test log count')
  if p.stem in required and p.stem.startswith(('final-','formal-','pilot-','fuzz-','confirmation-')) and p.stem not in failures:d.require(manifest(r['source_after'])==final,'final source epoch '+p.stem)
 for tag,count in [('final-pptx-r2',850),('final-opc-r2',471),('final-harness-r2',381)]:
  r=d.load('checks/'+tag+'.json');d.require(r['passed_tests']==count and r['failed_tests']==0,'final tests '+tag)
 log=artifact(d.load('checks/final-pptx-r2.json')['log']).decode()
 for name in ['shared_payload_pins_managed_bytes_across_cache_bypass_and_lazy_fallback','source_backed_cross_copy_lazily_copies_shared_payload_on_writer_memory_refusal']:d.require(name+' ... ok' in log,'required ownership/fallback test')
 for name,tag in [('baseline','baseline-build-r3'),('candidate','candidate-build')]:
  r=d.load('checks/'+tag+'.json');d.require(builds[name]['source_manifest']==r['source_after'] and builds[name]['revision']==r['revision'],'build custody')
 pre=d.load('baseline-preflight-30.json');oracle.check_report(pre)
 prebuild=d.load('draft-baseline-r2-build.json');d.require(pre['binary_sha256']==prebuild['binaries']['alloc']['sha256'] and pre['source_revision']==prebuild['revision'],'pre-optimization exploratory identity')
 d.require(protocol['status']=='frozen' and protocol['cpu']==2 and protocol['workers']==1 and protocol['samples']==30 and protocol['warmups']==3 and protocol['reports']==24,'frozen protocol')
 for n,h in protocol['bound_files'].items():d.require(sha(d.raw(n))==h,'frozen dependency '+n)
 normal=[dict(provider=p,corpus=c,build=b,instrumentation='normal') for p in ['bytes','range'] for b in ['baseline','candidate'] for c in ['plain','media-rich']]
 alloc=[dict(provider='bytes',corpus=c,build=b,instrumentation='alloc') for b in ['baseline','candidate'] for c in ['plain','media-rich']]
 expected_order=[dict(x,repeat='R1') for x in normal]+[dict(x,repeat='R2') for x in reversed(normal)]+[dict(x,repeat='R1') for x in alloc]+[dict(x,repeat='R2') for x in reversed(alloc)]
 d.require(protocol['order']==expected_order and protocol['pilot_lanes']==list(range(8))+list(range(16,20)),'balanced independent matrices')
 paths=list((ROOT/'runs').glob('*/receipt.json'))+list((ROOT/'pilots').glob('*/receipt.json'))
 d.require(len(paths)==36 and len(list((ROOT/'runs').glob('*/receipt.json')))==24,'report inventory')
 for p in paths:
  r=d.load(str(p.relative_to(ROOT)));lane=r['lane'];build=builds[lane['build']];binary=build['binaries'][lane['instrumentation']]
  d.require(lane==expected_order[int(p.parent.name)],'capture lane');d.require(r['status']=='pass' and r['exit_code']==r['oracle_exit_code']==0,'capture status')
  d.require(r['source_unchanged'] and r['source_before']==r['source_after']==builds['candidate']['source_manifest'],'capture source')
  d.require(r['binary']==binary and r['build_sha256']==sha(d.raw(lane['build']+'-build.json')),'binary binding')
  d.require(r['protocol_sha256']==sha(d.raw('protocol.json')) and r['capture_sha256']==sha(d.raw('capture.py')) and r['oracle_sha256']==sha(d.raw('verify-report.py')),'capture drivers')
  for a in r['artifacts'].values():artifact(a)
  report=json.loads(artifact(r['artifacts']['report']));oracle.check_report(report)
  d.require(report['binary_sha256']==binary['sha256'] and report['binary_bytes']==binary['bytes'] and report['current_exe']==binary['path'] and report['source_revision']==build['revision'],'report executable')
  d.require(report['instrumentation']==('none' if lane['instrumentation']=='normal' else 'system_allocator_operation_scoped'),'separate instrumentation')
  samples,warmup=(1,0) if r['pilot'] else (30,3)
  d.require(report['corpus']==lane['corpus'] and report['provider']==lane['provider'] and report['samples']==samples and report['warmup']==warmup,'report lane')
  for key,value in protocol['corpora'][lane['corpus']].items():d.require(report[key]==value,'exact corpus/output identity')
  argv=r['argv'];expected=[binary['path'],'provider-lifecycle','--corpus',lane['corpus'],'--provider',lane['provider'],'--samples',str(samples),'--warmup',str(warmup),'--source-revision',build['revision'],'--output',argv[argv.index('--output')+1]]
  if lane['provider']=='range':expected+=['--max-range','65536','--delay-us','200','--transfer-bytes-per-second','26214400','--transfer-delay-policy','separate-sleeps']
  d.require(argv==['taskset','-c','2','/usr/bin/time','-v','-o',argv[6]]+expected,'exact workload argv')
 module('confirmation-derive').check(final,builds)
 measurements=d.derive();d.require(measurements==d.load('measurements.json') and d.render(measurements)==d.raw('measurements.md').decode(),'derived measurements')
 decision=d.load('decision.json');d.require(decision['goal_complete'] is False and decision['native_coverage_promoted'] is False and decision['decision']=='retain shared PPTX decoded payloads','scoped decision')
 review=d.load('regression-review.json');d.require(review['flags']==measurements['review_flags'] and len(review['dispositions'])==len(review['flags']) and all(x['reviewed'] is True and x['reason'] and x['flag']==f for x,f in zip(review['dispositions'],review['flags'])),'all flags reviewed')
 smoke=d.load('checks/fuzz-smoke.json');d.require('Done 1000 runs' in artifact(smoke['log']).decode() and '-runs=1000' in smoke['argv'] and '-seed=453' in smoke['argv'],'fuzz completion')
 fuzz_build=d.load('checks/fuzz-build.json');d.require('RUSTC_BOOTSTRAP=1' in fuzz_build['argv'] and any('-Z sanitizer=address' in a and '-sanitizer-coverage-level=4' in a for a in fuzz_build['argv']),'instrumented fuzz')
 prepared=d.load('checks/fuzz-prepared.json');d.require(prepared['status']=='pass' and prepared['task']=='/tmp/litchi-goal-0453-opc-fuzz' and prepared['driver_sha256']==sha(d.raw('fuzz-control.py')),'fuzz preparation')
 inputs={r['path']:r for r in prepared['inputs']};d.require(len(inputs)==18 and inputs['fuzz_targets/parse_opc.rs']['sha256']==final['crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs'],'fuzz source')
 d.require(inputs['Cargo.toml']['sha256']==sha(d.raw('checks/fuzz-Cargo.toml.txt')),'fuzz manifest')
 seeds=d.load('seed-manifest.json');d.require(len(seeds)==16,'seed count')
 for r in seeds:
  artifact(r);c=inputs['corpus/'+Path(r['path']).name];d.require(c['sha256']==r['sha256'] and c['bytes']==r['bytes'],'seed copy')
 artifacts=d.load('checks/fuzz-artifacts.json');a={r['path']:r for r in artifacts['artifacts']}
 d.require(artifacts['status']=='pass' and artifacts['task']==prepared['task'] and artifacts['files']==len(a) and artifacts['bytes']==sum(r['bytes'] for r in a.values()),'fuzz inventory')
 for n,r in inputs.items():d.require(a[n]==r,'preserved fuzz input')
 d.require(a['Cargo.lock']['sha256']==sha(d.raw('checks/fuzz-Cargo.lock.txt')),'fuzz lock');d.require(a['target/x86_64-unknown-linux-gnu/release/parse_opc']['bytes']>0,'fuzz binary')
 if cleanup:
  p=d.load('checks/fuzz-cleanup.json');d.require(p['status']=='pass' and p['temporary_directory_absent'] and p['task']==prepared['task'],'fuzz cleanup')
  d.require(p['files_removed']==artifacts['files'] and p['bytes_removed']==artifacts['bytes'] and p['artifact_manifest_sha256']==sha(d.raw('checks/fuzz-artifacts.json')),'fuzz cleanup custody')
  d.require(d.load('checks/precleanup.json')['status']=='pass','precleanup verification')
  p=d.load('checks/binary-cleanup.json');d.require(p['status']=='pass' and p['temporary_directory_absent'] and p['task']=='/tmp/litchi-goal-0453-pptx-payload','binary cleanup')
  expected={}
  for name in ['draft-baseline-build.json','draft-baseline-r2-build.json','baseline-build.json','candidate-build.json']:
   for b in d.load(name)['binaries'].values():expected[Path(b['path']).name]=b
  d.require(p['binaries']==expected and p['files_removed']==len(expected) and p['bytes_removed']==sum(b['bytes'] for b in expected.values()),'cleaned binary identities')
 if sealed:
  inventory={}
  for line in (ROOT/'SHA256SUMS').read_text().splitlines():
   h,n=line.split('  ',1);d.require(n not in inventory and sha(member(n).read_bytes())==h,'seal member');inventory[n]=h
  d.require(set(inventory)=={str(p.relative_to(ROOT)) for p in ROOT.rglob('*') if p.is_file() and p.name!='SHA256SUMS'},'seal coverage')
 return {'status':'pass','change':453,'pptx_tests':850,'opc_tests':471,'harness_tests':381,'normal_samples':480,'allocator_samples':240,'separate_confirmation_samples':240,'fuzz_runs':1000,'sealed':sealed,'cleanup':cleanup}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--sealed',action='store_true');p.add_argument('--cleanup',action='store_true');a=p.parse_args();print(json.dumps(check(a.sealed,a.cleanup)))
