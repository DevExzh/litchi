#!/usr/bin/env python3
"""Read-only verification and descriptive summaries for the 0411 XLS capture."""
import argparse,hashlib,json,math,statistics,struct,subprocess,sys
from pathlib import Path

if not __debug__:
 raise RuntimeError('Verification requires assertions enabled')

CASES=['xls_semantic_open','xls_source_backed_open','xls_eager_open_list_worksheets','xls_source_backed_open_list_worksheets','xls_eager_open_one_cell','xls_source_backed_open_one_cell']
FLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes'
def load(path):
 if path.exists(): data=path.read_bytes()
 else: data=subprocess.check_output(['zstd','-q','-dc',str(path)+'.zst'])
 return json.loads(data)
def stats(samples):
 ordered=sorted(samples);n=len(samples)
 k=0
 for candidate in range(1,n//2+1):
  if sum(math.comb(n,j) for j in range(candidate))/2**n<=0.025:k=candidate
  else:break
 return dict(samples=n,p50_ns=statistics.median(ordered),mean_ns=statistics.mean(ordered),p95_ns=ordered[math.ceil(n*.95)-1],p99_ns=ordered[math.ceil(n*.99)-1],p50_iid_95pct_interval_ns=[ordered[k-1],ordered[n-k]] if k else None,median_interval_order_statistic=k,operations_per_second_from_mean=1e9/statistics.mean(ordered))
def constant(values,label):
 assert values and len(set(values))==1,label
 return values[0]
def main():
 p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--repo-root',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
 a=p.parse_args();sys.path.insert(0,str(a.repo_root));from tools import perf_compare,validate_perf_corpus_binding
 identity=load(a.root/'build-identity.json');protocol=load(a.root/'protocol.json')
 commands=load(a.root/'commands.json');assert commands and all(c['exit_code']==0 for c in commands)
 names=['Comments','Untouched'];name_bytes=b''.join(len(n.encode()).to_bytes(8,'little')+n.encode() for n in names)
 digests={'open+list':hashlib.sha256(name_bytes).hexdigest(),'open+one-cell':hashlib.sha256(struct.pack('<d',42.0)).hexdigest()}
 out=dict(schema_version=1,performance_claim='none',revision=identity['revision'],protocol=protocol,reports=[],corpus=None)
 for kind,count,samples,warmup in [('normal',4,500,20),('allocator',2,30,3)]:
  for repeat in range(1,count+1):
   label=f'{kind}-{repeat}';r=load(a.root/(label+'.json'));catalog=load(a.root/(label+'.catalog.json'))
   # Validate bindings through the same strict checker, using original files or
   # reconstructed private temporary copies for lossless compressed reports.
   import tempfile
   with tempfile.TemporaryDirectory(prefix='litchi-goal-0411-verify-') as tmp:
    rp=Path(tmp)/'report.json';cp=Path(tmp)/'catalog.json';rp.write_text(json.dumps(r));cp.write_text(json.dumps(catalog));validate_perf_corpus_binding.validate_paths(rp,cp)
   perf_compare.validate_parallel_metrics(r)
   assert r['environment']['git_revision']==identity['revision']
   assert r['environment']['git_worktree_dirty'] is False
   assert r['environment']['cpu_affinity']=='2' and r['environment']['rustflags']==FLAGS
   assert r['environment']['rustc_version'].startswith('rustc 1.98.1 ')
   binary='litchi-perf-baseline-alloc' if kind=='allocator' else 'litchi-perf-baseline'
   assert r['binary_identity']['binary_sha256']==identity['binaries'][binary]['sha256']
   assert r['tool']['instrumentation']==('system_allocator_operation_scoped' if kind=='allocator' else 'none')
   assert r['configuration']['cases']==CASES and len(r['results'])==6
   assert r['configuration']['filesystem_cache_states']==['warm'] and r['configuration']['execution_workers']==[1]
   assert r['configuration']['samples_per_case']==samples and r['configuration']['warmup_iterations_per_case']==warmup
   results=[]
   for case,v in zip(CASES,r['results']):
    assert v['case']==case
    corpus=v['corpus']
    if out['corpus'] is None:out['corpus']=corpus
    assert corpus==out['corpus']
    assert corpus['generator']==protocol['corpus_generator']
    assert corpus['archive_bytes']==16995840 and corpus['target_payload_bytes']==80946
    assert corpus['archive_member_count']==10
    expected=protocol['corpus_expected']
    for field in ['name','shape','entry_count','archive_member_count','archive_sha256','target_entry']:
     assert corpus[field]==expected[field],field
    assert corpus['target_payload_sha256']==expected['workbook_sha256']
    elapsed=v['elapsed_ns'];assert len(elapsed['samples'])==samples
    assert all(type(n) is int and n>0 for n in elapsed['samples'])
    assert sorted(elapsed['sample_order'])==list(range(samples))
    metrics=v['operation_metrics']
    perf_compare._validate_operation_metrics(metrics,label,elapsed['samples'],r['schema_version'],elapsed_sample_order=elapsed['sample_order'])
    allocation=metrics['allocation'];assert allocation['status']==('measured' if kind=='allocator' else 'unavailable')
    op='open+list' if 'list' in case else ('open+one-cell' if 'one_cell' in case else 'open')
    assert v['output_sha256']==(corpus['archive_sha256'] if op=='open' else digests[op])
    summary=dict(case=case,output_sha256=v['output_sha256'])
    if kind=='normal':
     summary['elapsed']=stats(elapsed['samples'])
     assert 'values' not in allocation['allocation_calls']
    else:
     summary['allocation']={}
     for field in ['allocation_calls','reallocation_calls','deallocation_calls','failed_allocation_calls','allocated_bytes','deallocated_bytes']:
      values=allocation[field]['values'];assert len(values)==samples
      summary['allocation'][field]=dict(min=min(values),max=max(values),values=values)
     assert allocation['failed_allocation_calls']['values']==[0]*samples
     summary['allocation']['live_and_peak_snapshots']={field:allocation[field] for field in ['live_bytes_before','live_bytes_after','peak_live_bytes_before','peak_live_bytes_after']}
    if 'source_backed' in case:
     source=v['source'];xls=source['xls'];assert xls['archive_sha256']==corpus['archive_sha256'] and xls['workbook_stream_sha256']==corpus['target_payload_sha256']
     summary['source']={}
     for field in ['read_calls','read_bytes','max_in_flight_reads']:
      assert len(source[field])==samples;summary['source'][field]=constant(source[field],field)
     summary['source']['xls']={}
     for field,values in xls.items():
      if isinstance(values,list):
       assert len(values)==samples;summary['source']['xls'][field]=constant(values,field)
      else:summary['source']['xls'][field]=values
     assert xls['parsed_sheet_counts']==[2]*samples and xls['parsed_cell_counts']==[int(op=='open+one-cell')]*samples
     assert xls['unselected_worksheet_read_bytes']==[0]*samples and xls['opaque_payload_read_bytes']==[0]*samples
     assert xls['source_version_stability_verified']==[True]*samples
     assert xls['complete_archive_materialized_bytes']==[0]*samples
     if op=='open+one-cell':
      assert min(xls['selected_worksheet_read_bytes'])>0 and xls['selected_query_reads_only_selected_worksheet']==[True]*samples
     else:assert xls['selected_worksheet_read_bytes']==[0]*samples and xls['open_reads_zero_worksheet_payload']==[True]*samples
    else:assert v.get('source') is None
    results.append(summary)
   out['reports'].append(dict(label=label,results=results))
 out['normal_repeat_spread']={}
 for case in CASES:
  vals=[next(v for v in r['results'] if v['case']==case)['elapsed']['p50_ns'] for r in out['reports'] if r['label'].startswith('normal')]
  out['normal_repeat_spread'][case]=dict(p50_ns=vals,max_minus_min_percent_of_min=100*(max(vals)-min(vals))/min(vals))
 a.output.write_text(json.dumps(out,indent=2)+'\n');print('Verified six reports, 12000 normal and 360 allocator observations')
if __name__=='__main__':main()
