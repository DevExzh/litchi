from pathlib import Path
import json,hashlib,collections
r=Path('/home/zhuhe/litchi-performance-0957')
def d(p):return {'path':str(p),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
s=r/'allocation-summary-v3.json';x=json.loads(s.read_text());peak='max_setup_drained_core_peak_above_setup_live_before';out=[]
for c in x['configs']:
 rows=[]
 for p in sorted((r/'capture-allocation-v1/reports').glob(c['case_id']+'-w'+str(c['workers'])+'-*/report.json')):rows+=json.loads(p.read_text())['case']['rows']
 out.append({'case_id':c['case_id'],'workers':c['workers'],'samples':len(rows),'drained_core_allocation_calls':[q['phases']['core_combined_with_drain']['allocation_calls']['mean'] for q in c['repeats']],'drained_core_allocated_bytes':[q['phases']['core_combined_with_drain']['allocated_bytes']['mean'] for q in c['repeats']],'setup_drained_peak_process_means':[q['derived_peak_fields'][peak]['mean'] for q in c['repeats']],'setup_drained_peak_sample_range':[min(q[peak] for q in rows),max(q[peak] for q in rows)],'postcore_first_threads':dict(collections.Counter(q['post_core_threads']['first_thread_count'] for q in rows)),'nonzero_worker_drain_samples':sum(any(q['worker_drain'][m] for m in ['allocation_calls','deallocation_calls','reallocation_calls']) for q in rows),'all_pre_sample_final_threads_one':all(q['pre_sample_threads']['final_thread_count']==1 for q in rows),'all_post_core_final_threads_one':all(q['post_core_threads']['final_thread_count']==1 for q in rows)})
flags=[]
for c in x['width_ratios_to_worker1']:
 flags.append({'case_id':c['case_id'],'workers':c['workers'],'repeat':c['repeat'],'primary':{m:c['ratios']['core_combined_with_drain'][m] for m in ['allocation_calls','allocated_bytes']},'peak':c['derived_peak_ratios'][peak],'all_phase_flags':{phase:{m:v for m,v in metrics.items() if v['regression_over_5_percent']} for phase,metrics in c['ratios'].items() if any(v['regression_over_5_percent'] for v in metrics.values())}})
p=r/'allocator-findings-v1.json'
with p.open('x') as f:json.dump({'schema':'litchi.performance.0957.findings.v1','input':d(s),'generator':d(Path(__file__)),'configs':out,'comparisons':flags,'scope':'Descriptive allocator data; no timing, RSS, unique-heap, causal attribution or adoption claim.'},f,indent=2);f.write('\n')
print(json.dumps(d(p)))
