#!/usr/bin/env python3
"""One create-only, bounded diagnostic profile; never a latency comparison."""
import argparse,datetime,hashlib,json,os,signal,subprocess
from pathlib import Path
if not __debug__:raise RuntimeError('optimized Python is not supported by this diagnostic helper')
ROOT=Path('/home/zhuhe/litchi-performance-0962')
def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()
def descriptor(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''):h.update(b)
 return {'path':str(p),'bytes':p.stat().st_size,'sha256':h.hexdigest()}
def write(p,v):
 with p.open('x') as f:json.dump(v,f,indent=2);f.write('\n')
def validate(v,c,n):
 assert v['schema']=='litchi.xlsx-managed-compression-benchmark.v1' and v['status']=='completed'
 for k,e in [('shape',c['shape']),('workload',c['workload']),('workers',1),('warmup_iterations',0),('measured_samples',n)]:assert v[k]==e,(k,v[k],e)
 case=v['case']
 for k in ['shape','workload','source_bytes','source_sha256','output_bytes','output_sha256','semantic_sha256','update_count','expected_cpu_tasks','context_limits']:assert case[k]==c[k],k
 assert case['preflight_actual_commit_inverse_verified'] is True
 assert case['admission_expected_worker_permits']==case['admission_observed_worker_permits']==0
 assert case['admission_cpu_tasks_used']==c['expected_cpu_tasks']
 assert case['sink_capacity_bytes']==c['context_limits']['output_bytes'] and case['sink_max_write_bytes']==65536
 assert case['warmup_iterations']==0 and case['measured_samples']==n and len(case['samples'])==n
 for i,s in enumerate(case['samples']):
  assert s['sample']==i
  phases=[s[k] for k in ['ingress_ns','edit_commit_ns','publication_ns','commit_drop_ns','total_ns']]
  assert all(type(x) is int and x>0 for x in phases) and sum(phases[:4])<=phases[4]
  for k in ['output_bytes','output_sha256','semantic_sha256']:assert s[k]==c[k],(i,k)
  assert s['output_bytes_used']==c['output_bytes'] and s['cpu_tasks_used']==c['expected_cpu_tasks']
  assert s['fixture_inverse_restored_exact_source'] is True
  assert type(s['sink_write_calls']) is int and s['sink_write_calls']>0
  assert type(s['sink_largest_write']) is int and 0<s['sink_largest_write']<=65536
  assert all(s['resources_after_drop'][k]==0 for k in ['memory','workers','io_concurrency','objects'])
 return {'status':'passed','samples':n,'oracles':'source/output/semantic hashes; preflight inverse; per-row fixture inverse, phase conservation, sink bounds and resource drain'}
def main():
 ap=argparse.ArgumentParser();ap.add_argument('--plan',type=Path,required=True);ap.add_argument('--case',required=True);ap.add_argument('--output',type=Path,required=True);a=ap.parse_args()
 plan=json.loads(a.plan.read_text());c=next(x for x in plan['cases'] if x['case_id']==a.case)
 binary=Path(plan['binary']['path']);assert descriptor(binary)==plan['binary'];assert descriptor(Path(__file__))==plan['capture_helper']
 assert plan['status']=='authorized' and plan['purpose']=='diagnostic CPU attribution only'
 a.output.mkdir(exist_ok=False);env={'PATH':'/usr/bin:/bin','LANG':'C','LC_ALL':'C','TZ':'UTC','RUST_BACKTRACE':'0'}
 cmd=['/usr/bin/taskset','--cpu-list','12-15','/usr/bin/perf','record','-e','cycles:u','-F',str(plan['frequency_hz']),'--call-graph','fp','--period','--timestamp','--sample-cpu','--buildid-all','--no-buildid-cache','--strict-freq','--user-callchains','-o',str(a.output/'perf.data'),'--',str(binary),'--shape',c['shape'],'--workload',c['workload'],'--workers','1','--warmup','0','--samples',str(plan['samples']),'--output',str(a.output/'output')]
 write(a.output/'start.json',{'utc':now(),'command':cmd,'env':env,'plan':descriptor(a.plan),'binary':descriptor(binary),'helper':descriptor(Path(__file__)),'cwd':plan['cwd']})
 with (a.output/'stdout.log').open('xb') as out,(a.output/'stderr.log').open('xb') as err:
  proc=subprocess.Popen(cmd,cwd=plan['cwd'],env=env,stdout=out,stderr=err,start_new_session=True)
  write(a.output/'launch.json',{'utc':now(),'pid':proc.pid,'process_group':proc.pid})
  timeout=False;failure=None
  try:rc=proc.wait(timeout=900)
  except subprocess.TimeoutExpired:
   timeout=True;os.killpg(proc.pid,signal.SIGTERM)
   try:rc=proc.wait(timeout=10)
   except subprocess.TimeoutExpired:os.killpg(proc.pid,signal.SIGKILL);rc=proc.wait()
  except BaseException as exc:
   failure=repr(exc)
   os.killpg(proc.pid,signal.SIGTERM)
   try:rc=proc.wait(timeout=10)
   except subprocess.TimeoutExpired:os.killpg(proc.pid,signal.SIGKILL);rc=proc.wait()
 try:os.killpg(proc.pid,0);group_clear=False
 except ProcessLookupError:group_clear=True
 write(a.output/'terminal.json',{'utc':now(),'returncode':rc,'timeout':timeout,'pid':proc.pid,'terminal_before_parse':True,'failure':failure,'process_group_clear':group_clear})
 if rc!=0 or timeout or failure or not group_clear:raise RuntimeError('profile child failed; no retry')
 report=a.output/'output/report.json'
 validation=validate(json.loads(report.read_text()),c,plan['samples']);validation['report']=descriptor(report)
 validation['binary_unchanged']=descriptor(binary)==plan['binary'];assert validation['binary_unchanged']
 validation['perf_data']=descriptor(a.output/'perf.data');validation['stderr']=descriptor(a.output/'stderr.log')
 write(a.output/'validation.json',validation);print(json.dumps({'status':'validated','case':a.case,'output':str(a.output)}))
if __name__=='__main__':main()
