"""Sequential write-once publication Callgrind capture; no timing claim."""
from __future__ import annotations
import argparse,datetime,hashlib,importlib.util,json,pathlib,subprocess,sys
P=pathlib.Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(pathlib.Path(p).read_text())
def desc(p):
 p=pathlib.Path(p).resolve();return dict(path=str(p),bytes=p.stat().st_size,sha256=hashlib.sha256(p.read_bytes()).hexdigest())
def write(p,v):
 with pathlib.Path(p).open('x') as f:json.dump(v,f,indent=2,sort_keys=True);f.write('\n')
def require(v,m):
 if not v:raise ValueError(m)
def load(name,path):
 spec=importlib.util.spec_from_file_location(name,path);mod=importlib.util.module_from_spec(spec);sys.modules[name]=mod;spec.loader.exec_module(mod);return mod
def utc():return datetime.datetime.now(datetime.timezone.utc).isoformat()
def check(plan):
 require(desc(plan['source_checker']['path'])==plan['source_checker'],'source helper custody')
 h=load('source0900',plan['source_checker']['path']);h.source_check(plan)
 for d in list(plan['tools'].values())+list(plan['symbols'].values())+[plan['native_context'],plan['prior_semantic_validator']]:require(desc(d['path'])==d,'input custody '+d['path'])
def command(plan,row,folder):
 return ['/usr/bin/taskset','-c',str(plan['cpu']),'/usr/bin/valgrind',*plan['options'],'--callgrind-out-file='+str(folder/'callgrind.out'),plan['binary']['path'],'--warmup','0','--samples','1','--case',plan['case'],'--xlsx-cell-crud-shape',row['shape'],'--json',str(folder/'report.json')]
def validate_semantics(plan,folder,shape):
 m=load('semantic0705',plan['prior_semantic_validator']['path']);r=read(folder/'report.json');require(r['schema_version']==1,'schema');require(r['binary_identity']['binary_sha256']==plan['binary']['sha256'],'binary');require(r['environment']['cpu_affinity']==str(plan['cpu']),'affinity')
 require(all(r['environment'][k] is None for k in ['git_revision','git_worktree_dirty','rustc_version']),'expected unavailable runtime metadata')
 require(r['tool']['profile']=='release' and r['tool']['instrumentation']=='none','native binary under external observer')
 c=r['configuration'];require(c['samples_per_case']==1 and c['warmup_iterations_per_case']==0 and c['cases']==[plan['case']] and c['xlsx_cell_crud_shapes']==[shape],'configuration')
 require(len(r['results'])==1,'result count');v=r['results'][0];require(v['case']==plan['case'],'case')
 corpus=m.validate_corpus(v['corpus'],shape);m.validate_sink(v['sink'],shape);phases,constants=m.validate_source(v,corpus,shape,1,compatibility=False)
 reference=next(v for v in read(plan['native_context']['path'])['rows'] if v['shape']==shape and v['lane']=='native')
 require(corpus==reference['corpus'] and constants==reference['constants'] and v['sink']==reference['sink'],'native context semantic/resource identity')
 return dict(status='pass',shape=shape,corpus=corpus,constants=constants,sink=v['sink'],timing_policy='Instrumented elapsed durations retained but excluded from native conclusions.',native_reference=plan['native_context'])
def main():
 parser=argparse.ArgumentParser();parser.add_argument('--index',type=int,required=True);a=parser.parse_args();plan=read(P/'plan.json');row=plan['schedule'][a.index];require(row['index']==a.index,'index')
 if a.index>0:
  admission=read(P/'first-child-admission.json');require(admission['status']=='pass','first child raw admission')
  prior=plan['schedule'][a.index-1];require(read(P/'capture'/f"{prior['index']:02}-{prior['shape']}"/'semantics.json')['status']=='pass','prior child semantic gate')
 check(plan);folder=P/'capture'/f"{a.index:02}-{row['shape']}";folder.mkdir(parents=True);argv=command(plan,row,folder)
 start=dict(argv=argv,cwd=str(ROOT),environment=plan['environment'],plan=desc(P/'plan.json'),helper=desc(__file__),binary=plan['binary'],row=row,started_utc=utc());write(folder/'start.json',start)
 with (folder/'stdout.txt').open('x') as out,(folder/'stderr.txt').open('x') as err:
  r=subprocess.run(argv,cwd=ROOT,env=plan['environment'],stdout=out,stderr=err)
 end=dict(start,finished_utc=utc(),exit_code=r.returncode,artifacts=[desc(f) for f in sorted(folder.iterdir()) if f.is_file()]);write(folder/'command.json',end);require(r.returncode==0,'profile command failed')
 try:semantics=validate_semantics(plan,folder,row['shape'])
 except Exception as e:write(folder/'semantic-failure.json',dict(status='fail',error=repr(e)));raise
 write(folder/'semantics.json',semantics);check(plan);print('profile',a.index,row['shape'],'semantic pass',flush=True)
if __name__=='__main__':main()
