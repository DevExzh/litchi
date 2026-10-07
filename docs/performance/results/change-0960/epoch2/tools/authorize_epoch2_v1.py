from pathlib import Path
import json,hashlib,datetime,os,importlib.util
r=Path('/tmp/litchi-performance-0960-epoch2');old=Path('/home/zhuhe/litchi-performance-0960')
def desc(p):
 s=p.stat();return {'path':str(p),'bytes':s.st_size,'mode':s.st_mode&511,'executable':bool(s.st_mode&73),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
def write(p,x):
 with p.open('x') as f:f.write(json.dumps(x,indent=2)+'\n')
def module(p,name):
 s=importlib.util.spec_from_file_location(name,p);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);return m
p=json.loads((r/'plan-prospective-v1.json').read_text());review=r/'review/rerun-feasibility-v1.json';assert review.exists()
go={'status':'GO','created_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'scope':'One fresh independent screening epoch; no adoption','prospective_plan':desc(r/'plan-prospective-v1.json'),'independent_review':desc(review),'source_revalidation':desc(r/'source-revalidation-v1.json'),'host':desc(r/'host-v1.json'),'original_rule':p['root_screening_rule'],'original_selector':p['root_screen_selector'],'exclusions':'Prior incomplete epoch excluded completely; no timing analysis, pooling, replacement, or retry. No further automatic epoch.','capture_roots':p['capture_root']}
write(r/'root-go-v1.json',go);p['status']='authorized';p['plan_revision']='epoch2-final-v1';p['root_go']={'status':'GO','receipt':desc(r/'root-go-v1.json')};write(r/'plan-final-v1.json',p)
os.environ['LITCHI_PERF_0960_CAPTURE_APPROVED']='1';p['_plan_path']=str(r/'plan-final-v1.json')
c=module(old/'tools/screen_0960_v7.py','capture');c.validate_plan(p,capture=True)
d=module(old/'tools/reduce_0960_v3.py','reducer');canonical,_=d.canonical_plan(p);d.validate_plan_bindings(p,canonical,r/'plan-final-v1.json',old/'tools/reduce_0960_v3.py')
write(r/'preflight-v1.json',{'status':'passed','runner_actual_capture_preflight':True,'reducer_actual_plan_bindings':True,'plan':desc(r/'plan-final-v1.json'),'capture_cpus_allowed':set([12,13,14,15])<=os.sched_getaffinity(0),'source_revalidation':desc(r/'source-revalidation-v1.json')});print(desc(r/'plan-final-v1.json'))
