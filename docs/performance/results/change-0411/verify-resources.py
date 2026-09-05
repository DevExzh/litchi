#!/usr/bin/env python3
"""Validate and summarize the whole-process 0411 resource captures."""
import argparse,csv,json,re,subprocess,tempfile,sys
from pathlib import Path

if not __debug__:
 raise RuntimeError('Verification requires assertions enabled')

def read(path):
 return path.read_text() if path.exists() else subprocess.check_output(['zstd','-q','-dc',str(path)+'.zst'],text=True)
def main():
 p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--repo-root',type=Path,default=Path.cwd())
 a=p.parse_args();sys.path.insert(0,str(a.repo_root.resolve()));from tools import validate_perf_corpus_binding
 identity=json.loads(read(a.root/'build-identity.json'));commands=json.loads(read(a.root/'commands.json'))
 out=dict(schema_version=1,performance_claim='none',scope='whole process: corpus generation, layout/oracle, warmups, operations, cloning, validation and serialization; not operation-local',rss_kib={},profiles={},counters={})
 for kind,count in [('normal',4),('allocator',2)]:
  for i in range(1,count+1):
   label=f'{kind}-{i}';value=read(a.root/(label+'.time.txt'));assert 'Exit status: 0' in value
   out['rss_kib'][label]=int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',value)[1])
 for family in ['eager','source_backed']:
  case=f'xls_{family}_open_one_cell'
  for label,samples,warmup in [(family+'-profile',1000,20),*((f'{family}-stat-{i}',300,10) for i in range(1,4))]:
   r=json.loads(read(a.root/(label+'.json')));assert r['configuration']['cases']==[case] and len(r['results'])==1
   assert r['configuration']['samples_per_case']==samples and r['configuration']['warmup_iterations_per_case']==warmup
   assert len(r['results'][0]['elapsed_ns']['samples'])==samples
   assert sorted(r['results'][0]['elapsed_ns']['sample_order'])==list(range(samples))
   with tempfile.TemporaryDirectory(prefix='litchi-goal-0411-resource-') as tmp:
    rp=Path(tmp)/'report.json';cp=Path(tmp)/'catalog.json';rp.write_text(json.dumps(r));cp.write_text(read(a.root/(label+'.catalog.json')));validate_perf_corpus_binding.validate_paths(rp,cp)
   assert r['environment']['git_revision']==identity['revision'] and r['environment']['git_worktree_dirty'] is False
   assert r['binary_identity']['binary_sha256']==identity['binaries']['litchi-perf-baseline']['sha256']
   assert r['environment']['cpu_affinity']=='2'
   assert r['tool']['instrumentation']=='none'
   assert r['results'][0]['corpus']['archive_sha256']=='6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53'
   assert r['results'][0]['output_sha256']==__import__('hashlib').sha256(__import__('struct').pack('<d',42.0)).hexdigest()
   command=next(c for c in commands if c['label']==label);assert command['exit_code']==0
   if label.endswith('profile'):
    text=read(a.root/(label+'-self.stdout'));assert '# Total Lost Samples: 0' in text
    out['profiles'][family]=dict(lost_samples=0,perf_record_command=command['argv'],whole_process_event_count=int(re.search(r'Event count \(approx.\): (\d+)',text)[1]))
   else:
    values={}
    for row in csv.reader(line for line in read(a.root/(label+'.csv')).splitlines() if line and not line.startswith('#')):
     assert len(row)>=5
     count=float(row[0]);pct=float(row[4]);assert count>=0 and 0<pct<=100
     values[row[2]]=dict(count=count,unit=row[1],running_ns=int(row[3]),scheduled_percent=pct)
    expected={'task-clock','cycles','instructions','branches','branch-misses','page-faults','context-switches','cpu-migrations','l2_cache_req_stat.dc_access_in_l2','l2_cache_req_stat.dc_hit_in_l2'}
    assert set(values)==expected
    assert all(values[event]['count']>0 for event in expected-{'context-switches','cpu-migrations'})
    out['counters'][label]=dict(events=values,ipc_from_scaled_whole_process_counts=values['instructions']['count']/values['cycles']['count'],branch_miss_percent_from_scaled_whole_process_counts=100*values['branch-misses']['count']/values['branches']['count'])
 out['limitations']=['Perf scales multiplexed hardware events; retain scheduled percentages and do not infer operation-local IPC.','Native L2 requests/hits are not exact L1 or LLC metrics.','RSS is a child high-water mark for the entire six-case dispatch, not a per-case peak.','No isolated peak live bytes, physical-I/O or exact request-size distribution is inferred.']
 a.output.write_text(json.dumps(out,indent=2)+'\n');print('Verified 6 RSS captures, 2 profiles and 6 PMU reports')
if __name__=='__main__':main()
