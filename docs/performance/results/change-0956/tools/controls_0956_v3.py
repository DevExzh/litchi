#!/usr/bin/env python3
"""No-workload v3 custody controls for the 0956 capture/reducer pair."""
from __future__ import annotations
import copy, hashlib, importlib.util, json, shutil, tempfile
from pathlib import Path

HERE=Path(__file__).resolve().parent
ROOT=HERE.parent
CAP_SPEC=importlib.util.spec_from_file_location('capture_0956_v3', HERE/'capture_0956_v3.py'); CAP=importlib.util.module_from_spec(CAP_SPEC); CAP_SPEC.loader.exec_module(CAP)
RED_SPEC=importlib.util.spec_from_file_location('reduce_0956_v3', HERE/'reduce_0956_v3.py'); RED=importlib.util.module_from_spec(RED_SPEC); RED_SPEC.loader.exec_module(RED)

def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def desc(p): return {'path':str(p.resolve()),'bytes':p.stat().st_size,'sha256':sha(p)}
def wx(p,v):
 p.parent.mkdir(parents=True,exist_ok=True)
 with p.open('xb') as f:f.write((json.dumps(v,indent=2,sort_keys=True)+'\n').encode());f.flush()
def fail(fn,label):
 try: fn()
 except Exception as exc: return {'name':label,'status':'rejected','error':str(exc)}
 raise AssertionError('control unexpectedly passed: '+label)
def synthetic_report(row,plan):
 cfg=next(x for x in plan['cases'] if x['id']==row['case_id']); lim=cfg['context_limits'][str(row['width'])]
 samples=[]
 for i in range(plan['execution']['samples']):
  samples.append({'sample':i,'ingress_ns':10,'edit_commit_ns':20,'publication_ns':30,'commit_drop_ns':5,'total_ns':65,'output_bytes':cfg['output_bytes'],'output_bytes_used':cfg['output_bytes'],'output_sha256':cfg['output_sha256'],'semantic_sha256':cfg['semantic_sha256'],'cpu_tasks_used':cfg['cpu_tasks'],'sink_write_calls':1,'sink_largest_write':1,'resources_after_drop':{'memory':0,'workers':0,'io_concurrency':0,'objects':0},'fixture_inverse_restored_exact_source':True})
 return {'schema':CAP.REPORT_SCHEMA,'status':'completed','shape':row['shape'],'workload':row['workload'],'workers':row['width'],'warmup_iterations':plan['execution']['warmup'],'measured_samples':plan['execution']['samples'],'timing_scope':'with fresh finite managed context prepared outside the clock, the timer begins at SourceBackedEditor construction and covers edit/commit, source-backed publication, returned snapshot drop, and commit/package worker-pool teardown; precomputed source/cache/context setup and bounded sink allocation are outside','phase_scope':'ingress_ns is the direct SourceBackedEditor constructor; edit_commit_ns is selector staging, cell sets, and MultiCommit::commit; publication_ns is publish_multi_commit_to_stream plus returned MultiSnapshot drop; commit_drop_ns is MultiCommit drop; total_ns starts at the constructor and ends after commit teardown','setup_scope':'deterministic corpus and expected output construction, frozen source/output/semantic oracles, SourceCacheLimits and output ceiling, one untimed admission preflight, inverse patch, fresh OwnedSource clone, finite managed context, and sink capacity reservation are outside measured phases','verification_scope':'exact output bytes, immutable fixture source hash, output/semantic hashes, workbook verification, sink bounds, CpuTasks/OutputBytes counters, and post-drop Memory/Workers/IoConcurrency/Objects checks run outside timers; the untimed admission preflight verifies an actual managed-commit inverse (preflight_actual_commit_inverse_verified=true), while each timed sample applies the precomputed inverse patch only as a fixture to its published bytes (fixture_inverse_restored_exact_source)','performance_claim':'none; raw phase timings describe this explicit in-memory harness and do not claim OS thread counts, speedup, allocation behavior, or physical I/O','case':{'id':f"{row['shape']}-{row['workload']}-workers{row['width']}",'shape':row['shape'],'workload':row['workload'],'workers':row['width'],'update_count':1,'source_bytes':cfg['source_bytes'],'source_sha256':cfg['source_sha256'],'output_bytes':cfg['output_bytes'],'output_sha256':cfg['output_sha256'],'semantic_sha256':cfg['semantic_sha256'],'expected_cpu_tasks':cfg['cpu_tasks'],'admission_expected_worker_permits':cfg['permits'][str(row['width'])],'admission_observed_worker_permits':cfg['permits'][str(row['width'])],'admission_cpu_tasks_used':cfg['cpu_tasks'],'preflight_actual_commit_inverse_verified':True,'context_limits':lim,'sink_capacity_bytes':lim['output_bytes'],'sink_max_write_bytes':65536,'warmup_iterations':plan['execution']['warmup'],'measured_samples':plan['execution']['samples'],'samples':samples}}

def build_packet(plan,packet):
 rows=CAP.validate_plan(plan,for_execution=False); packet.mkdir(); (packet/'children').mkdir()
 start={'schema':CAP.RUNNER_SCHEMA+'.outer-start','plan':desc(ROOT/'plan-v3-draft.json'),'runner':plan['helpers']['capture'],'reducer':plan['helpers']['reducer'],'controls':plan['helpers']['controls'],'source_manifest':desc(Path(plan['source']['source_manifest'])),'started_utc':'synthetic','row_count':len(rows)}
 wx(packet/'capture-start.json',start)
 time_text='''Maximum resident set size (kbytes): 100\nUser time (seconds): 1.0\nSystem time (seconds): 0.1\nMajor (requiring I/O) page faults: 0\nMinor (reclaiming a frame) page faults: 2\nVoluntary context switches: 3\nInvoluntary context switches: 4\n'''
 for idx,row in enumerate(rows):
  child=packet/'children'/row['id']; child.mkdir(); out=child/'workload-output'; out.mkdir()
  expected=RED._capture._expected_child_contract(row,plan,child)
  argv=expected['argv']; startrec={'schema':CAP.RUNNER_SCHEMA+'.start','index':idx,'row':row,'argv':argv,'cwd':expected['cwd'],'environment':expected['environment'],'cpu_mask_requested':expected['cpu_mask'],'started_utc':'synthetic','deadline_seconds':expected['deadline_seconds'],'terminal_path':str((child/'terminal.json').resolve())}
  report=synthetic_report(row,plan); wx(out/'report.json',report); (child/'time-v.txt').write_text(time_text); (child/'stdout.log').write_bytes(b''); (child/'stderr.log').write_bytes(b'')
  term={'schema':CAP.RUNNER_SCHEMA+'.terminal','index':idx,'row':row,'argv':argv,'cwd':expected['cwd'],'environment':expected['environment'],'cpu_mask_requested':expected['cpu_mask'],'cpu_mask_observed':expected['cpu_mask'],'affinity_status':'verified','started_utc':'synthetic','ended_utc':'synthetic','elapsed_seconds':0.0,'deadline_seconds':expected['deadline_seconds'],'returncode':0,'deadline_exceeded':False,'forced_cleanup':False,'process_group_clear':True,'status':'completed','artifacts':{}}
  wx(child/'start.json',startrec); wx(child/'terminal.json',term)
  artifacts={name:desc(path) for name,path in {'stdout':child/'stdout.log','stderr':child/'stderr.log','time':child/'time-v.txt','report':out/'report.json'}.items()}; wx(child/'post-terminal-artifacts.json',artifacts)
  validation={'schema':CAP.RUNNER_SCHEMA+'.validation','status':'passed','row':row,'report':desc(out/'report.json'),'terminal':desc(child/'terminal.json')}; wx(child/'validation.json',validation)
 outer={'schema':CAP.RUNNER_SCHEMA+'.outer-terminal','status':'completed','ended_utc':'synthetic','children_started':len(rows),'children_expected':len(rows),'failed_index':None,'capture_start':desc(packet/'capture-start.json'),'plan':desc(ROOT/'plan-v3-draft.json'),'runner':plan['helpers']['capture']}; wx(packet/'capture-terminal.json',outer)

def main():
 plan=json.loads((ROOT/'plan-v3-draft.json').read_text()); CAP.validate_plan(plan,for_execution=False)
 checks=[{'name':'positive_plan_and_custody','status':'passed','rows':144}]
 with tempfile.TemporaryDirectory(prefix='0956-v3-controls-',dir='/tmp') as td:
  packet=Path(td)/'packet'; build_packet(plan,packet)
  out=Path(td)/'reduction.json'; RED.reduce(ROOT/'plan-v3-draft.json',packet,out); result=json.loads(out.read_text()); assert result['status']=='completed' and result['child_count']==144
  checks.append({'name':'positive_synthetic_144_reduction','status':'passed','children':result['child_count'],'groups':len(result['groups'])})
  child_id=result["groups"][0]["children"][0]["id"]
  child=packet/"children"/child_id
  tp=child/"start.json"; original_start=tp.read_bytes(); val=json.loads(original_start); val["row"]["position"]=99; tp.write_text(json.dumps(val,indent=2,sort_keys=True)+"\n")
  checks.append(fail(lambda:RED.reduce(ROOT/"plan-v3-draft.json",packet,Path(td)/"tampered-row-output.json"),"tampered_schedule_row_rejected"))
  tp.write_bytes(original_start)
  pp=child/"post-terminal-artifacts.json"; original_post=pp.read_bytes(); pv=json.loads(original_post); pv["stdout"]["sha256"]="0"*64; pp.write_text(json.dumps(pv,indent=2,sort_keys=True)+"\n")
  checks.append(fail(lambda:RED.reduce(ROOT/"plan-v3-draft.json",packet,Path(td)/"tampered-artifact-output.json"),"tampered_post_terminal_artifact_rejected"))
  pp.write_bytes(original_post)

 zero=RED.bootstrap_geometric_mean([0.0,1.0,1.0,1.0,1.0,1.0],seed=956,resamples=1000); assert zero['point_geometric_mean']==0.0 and all(isinstance(v,(int,float)) for v in zero['ci95']); checks.append({'name':'zero_counter_bootstrap','status':'passed','point':zero['point_geometric_mean'],'ci95':zero['ci95']})
 out=ROOT/'review/controls-v3-custody-final.json'; wx(out,{'schema':'litchi.performance.0956.native-width-controls.v3','status':'passed','workload_started':False,'plan':desc(ROOT/'plan-v3-draft.json'),'capture':desc(HERE/'capture_0956_v3.py'),'reducer':desc(HERE/'reduce_0956_v3.py'),'checks':checks})
 print(json.dumps({'output':str(out),'sha256':sha(out),'checks':len(checks)},sort_keys=True))
if __name__=='__main__': main()
