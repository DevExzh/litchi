from pathlib import Path
import hashlib,json,sys
repo=Path(__file__).resolve().parents[5]; root=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(repo))
from tools import summarize_crud_baseline as baseline, perf_compare, validate_perf_corpus_binding as binding
capture=json.loads((root/'capture.json').read_text()); build=json.loads((root/'build-identity.json').read_text())
rows=[r for r in capture['runs'] if r['preflight']]
assert len(rows)==32 and len({r['key'] for r in rows})==32
oracles={}; records=[]
for run in rows:
 assert run['exit_code']==0 and run['samples']==1 and run['warmups']==0
 for key,sha in run['artifact_sha256'].items():
  assert hashlib.sha256((root/run[key]).read_bytes()).hexdigest()==sha
 report=json.loads((root/run['report']).read_text()); catalog=json.loads((root/run['catalog']).read_text())
 binding.validate_binding(report,catalog)
 role=build['roles'][run['role']]; binary=role['binaries'][run['phase']]
 assert report['environment']['git_revision']==role['source']['revision']
 assert report['environment']['git_worktree_dirty'] is False
 assert report['environment']['rustc_version'].startswith('rustc 1.98.1 ')
 assert report['environment']['logical_cpus_available']==1
 assert report['binary_identity']['binary_sha256']==binary['sha256']==run['binary_sha256']
 assert report['binary_identity']['binary_bytes']==binary['bytes']
 assert len(report['results'])==1
 result=report['results'][0]; assert result['case']==run['selector']
 elapsed=baseline._validate_elapsed(result,1,run['report'])
 source=result['source']['pptx_cross_copy']; assert all(v is True for v in source['gates'].values())
 assert source['expected_output_sha256']==result['output_sha256']==source['output_sha256'][0]
 assert source['destination_archive_sha256']==result['corpus']['archive_sha256']
 assert source['source_archive_sha256']!=source['destination_archive_sha256']
 assert source['destination_slide_count_after']==source['destination_slide_count_before']+1
 assert source['destination_slide']<source['destination_slide_count_before']
 assert source['insertion_position']<=source['destination_slide_count_before']
 for name in ('plan_ns','commit_ns','publication_ns','reopen_ns','output_sha256'):
  assert len(source[name])==1
 phase_total=sum(source[name][0] for name in ('plan_ns','commit_ns','publication_ns'))
 life=run['selector'].endswith('_lifecycle')
 if life:
  assert source['lifecycle_ns']==elapsed['samples'] and phase_total<=elapsed['samples'][0]
 else:
  assert 'lifecycle_ns' not in source and phase_total==elapsed['samples'][0]
 sink=result['sink']; assert 0<sink['largest_write']<=65536
 assert sum(sink['write_size_buckets'].values())==sink['write_calls']
 assert sink['write_size_buckets']['bytes_over_65536']==0
 operation=result.get('operation_metrics')
 if life:
  perf_compare._validate_operation_metrics(operation,run['report'],elapsed['samples'],1,elapsed_sample_order=elapsed['sample_order'])
  allocation=operation['allocation']; expected='measured' if run['phase']=='allocator' else 'unavailable'
  assert allocation['status']==expected
  if expected=='measured': perf_compare._validate_allocator_operation_evidence(result,run['report'],1)
 else:
  assert operation.get('allocation') is None
  perf_compare._validate_operation_metrics(operation,run['report'],elapsed['samples'],1,elapsed_sample_order=elapsed['sample_order'])
 stable={k:v for k,v in source.items() if k not in {'plan_ns','commit_ns','publication_ns','reopen_ns','output_sha256','lifecycle_ns'}}
 identity=(stable,result['corpus'],result['output_sha256'],sink['accepted_bytes'])
 previous=oracles.setdefault(run['selector'],identity); assert previous==identity
 records.append({'key':run['key'],'report':run['report'],'output_sha256':result['output_sha256'],'accepted_bytes':sink['accepted_bytes'],'write_calls':sink['write_calls'],'allocation_status':operation['allocation']['status'] if life else 'unavailable'})
print(json.dumps({'status':'pass','reports':len(records),'selectors':len(oracles),'records':records},indent=2))
