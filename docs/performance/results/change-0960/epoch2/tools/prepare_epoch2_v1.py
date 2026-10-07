from pathlib import Path
import json,hashlib,datetime,os,subprocess,platform
r=Path('/tmp/litchi-performance-0960-epoch2'); old=Path('/home/zhuhe/litchi-performance-0960')
def desc(p):
 p=Path(p);s=p.stat();return {'path':str(p),'bytes':s.st_size,'mode':s.st_mode&511,'executable':bool(s.st_mode&73),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
def write(p,x):
 with p.open('x') as f:f.write(json.dumps(x,indent=2)+'\n')
now=datetime.datetime.now(datetime.timezone.utc).isoformat()
h={'created_utc':now,'affinity':sorted(os.sched_getaffinity(0)),'capture_cpus':[12,13,14,15],'uname':list(platform.uname()),'lscpu':subprocess.check_output(['lscpu'],text=True),'meminfo':Path('/proc/meminfo').read_text(),'output_filesystem':subprocess.check_output(['df','-T','/tmp'],text=True),'scope':'Fresh independent epoch; outputs on tmpfs. No pooling or comparison to partial epoch timings. Same within-epoch baseline/candidate environment. No governor modification, no cold-cache claim.'}
write(r/'host-v1.json',h)
p=json.loads((old/'plan-final-v1.json').read_text());p['created_utc']=now;p['plan_id']='0960-raw-cell-capacity-screen-epoch2';p['plan_revision']='epoch2-prospective-v1';p['status']='prospective';p['status_note']='Fresh full experiment after interrupted unsealed first epoch; no samples reused and no timing-driven restart.';p['supersedes']=None
p['capture_root']={l:str(r/f'capture-{l}-v1') for l in ['native','allocator']};p['analysis_root']=str(r/'analysis-v1')
hd=desc(r/'host-v1.json');p['host']={**hd,'status':'bound','receipt':hd};p['root_go']={'status':'pending_review'}
for l in ['native','allocator']:p['lanes'][l]['execution']['workload_cwd']=str(r)
p['epoch_history']={'prior_plan':desc(old/'plan-final-v1.json'),'prior_packet':'/home/zhuhe/litchi-performance-0960/capture-native-v1','prior_state':'incomplete;33 validated of36;34 start records; no lane terminal; missing tool handle; no allocator','prior_interruption_inventory':desc(Path('/home/zhuhe/code/litchi/docs/performance/results/change-0960/review/partial-capture-interruption-v1.json')),'prior_timings_inspected_for_selection':False,'pooled_samples':False,'sample_replacement':False,'policy':'One fresh complete screening epoch, fail-stop/no child retries. Retain every outcome. Any failure ends this epoch; no automatic further epoch. Original rule and selector unchanged.'}
p['source_revalidation']=desc(r/'source-revalidation-v1.json')
write(r/'plan-prospective-v1.json',p);print(desc(r/'plan-prospective-v1.json'))
