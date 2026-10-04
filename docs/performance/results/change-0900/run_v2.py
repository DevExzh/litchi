"""0900 current-source native XLSX phase capture and offline replay."""
from __future__ import annotations
import argparse, datetime, hashlib, importlib.util, json, pathlib, statistics, subprocess, sys
P=pathlib.Path(__file__).resolve().parent
ROOT=P.parents[3]
def read(p): return json.loads(pathlib.Path(p).read_text())
def sha(p): return hashlib.sha256(pathlib.Path(p).read_bytes()).hexdigest()
def desc(p):
 p=pathlib.Path(p).resolve();return dict(path=str(p),bytes=p.stat().st_size,sha256=sha(p))
def write(p,v):
 with pathlib.Path(p).open('x') as f:json.dump(v,f,indent=2,sort_keys=True);f.write('\n')
def require(v,message):
 if not v:raise ValueError(message)
def utc():return datetime.datetime.now(datetime.timezone.utc).isoformat()
def load_validator(plan):
 d=plan['prior_semantic_validator'];require(desc(d['path'])==d,'validator custody')
 s=importlib.util.spec_from_file_location('phase0705',d['path']);m=importlib.util.module_from_spec(s);sys.modules[s.name]=m;s.loader.exec_module(m);return m

def source_check(plan):
 d=plan['source_manifest'];require(desc(d['path'])==d,'source manifest custody')
 for name,item in read(d['path'])['files'].items():
  f=ROOT/name;require(f.stat().st_size==item['bytes'] and sha(f)==item['sha256'],'source mismatch '+name)
 require(desc(plan['binary']['path'])==plan['binary'],'binary custody')
 require(desc(plan['build']['path'])==plan['build'],'build receipt custody')

def argv(plan,folder,shape,lane):
 samples=plan['qualification_samples'] if lane=='qualification' else plan['samples']
 warmup=plan['qualification_warmup'] if lane=='qualification' else plan['warmup']
 return ['/usr/bin/taskset','-c',str(plan['cpu']),plan['binary']['path'],'--warmup',str(warmup),'--samples',str(samples),'--case',plan['case'],'--xlsx-cell-crud-shape',shape,'--json',str(folder/'report.json')]

def validate(plan,folder,shape,lane,m):
 samples=plan['qualification_samples'] if lane=='qualification' else plan['samples']
 warmup=plan['qualification_warmup'] if lane=='qualification' else plan['warmup']
 r=read(folder/'report.json');require(r['schema_version']==1,'schema');require(r['binary_identity']['binary_sha256']==plan['binary']['sha256'],'report executable')
 require(r['tool']['profile']=='release' and r['tool']['instrumentation']=='none','native release')
 require(r['environment']['cpu_affinity']==str(plan['cpu']),'affinity')
 # git_output uses removed compile-time CARGO_MANIFEST_DIR, not current cwd.
 require(r['environment']['git_revision'] is None and r['environment']['git_worktree_dirty'] is None,'expected unavailable build-root Git metadata')
 c=r['configuration'];require(c['cases']==[plan['case']] and c['xlsx_cell_crud_shapes']==[shape],'case/shape');require(c['samples_per_case']==samples and c['warmup_iterations_per_case']==warmup,'counts')
 require(len(r['results'])==1,'result count');v=r['results'][0];require(v['case']==plan['case'],'result case')
 corpus=m.validate_corpus(v['corpus'],shape);m.validate_sink(v['sink'],shape);total=m.check_reported_statistics(v['elapsed_ns'],samples)
 phases,constants=m.validate_source(v,corpus,shape,samples,compatibility=False)
 return dict(shape=shape,lane=lane,samples=samples,corpus=corpus,constants=constants,sink=v['sink'],total=total,phases={k:m.stats(vals) for k,vals in phases.items()},phase_shares={k:sum(phases[k])/sum(v['elapsed_ns']['samples']) for k in m.PHASES})

def schedule(plan,lane):
 if lane=='qualification':return [(0,s) for s in plan['shapes']]
 return [(b,s) for b in range(plan['native_repeats']) for s in (plan['shapes'] if b%2==0 else list(reversed(plan['shapes'])))]

def capture(plan,lane,m):
 source_check(plan);target=P/lane
 if lane=='qualification':require(target.is_dir(),'retained first qualification required')
 else:target.mkdir()
 rows=[]
 for block,shape in schedule(plan,lane):
  folder=target/f'b{block:02}-{shape}'
  if lane=='qualification' and block==0 and shape=='medium':
   cmd=read(folder/'command.json');require(cmd['exit_code']==0,'retained qualification command')
   for d in cmd['artifacts']:require(desc(d['path'])==d,'retained qualification custody')
   require(cmd['helper']==desc(P/'run.py'),'original helper custody')
   row=validate(plan,folder,shape,lane,m);write(folder/'validation.json',row);rows.append(row);print('retained qualification replay passed; no rerun',flush=True);continue
  folder.mkdir();command=argv(plan,folder,shape,lane)
  start=dict(argv=command,cwd=str(ROOT),environment=plan['environment'],plan=desc(P/'plan.json'),helper=desc(__file__),binary=plan['binary'],started_utc=utc(),block=block,shape=shape,lane=lane)
  write(folder/'start.json',start)
  with (folder/'stdout.txt').open('x') as out,(folder/'stderr.txt').open('x') as err:
   result=subprocess.run(command,cwd=ROOT,env=plan['environment'],stdout=out,stderr=err)
  end=dict(start,finished_utc=utc(),exit_code=result.returncode,artifacts=[desc(f) for f in sorted(folder.iterdir()) if f.is_file()])
  write(folder/'command.json',end);require(result.returncode==0,'child failure '+str(folder))
  try: row=validate(plan,folder,shape,lane,m)
  except Exception as e:
   write(folder/'validation-failure.json',dict(error=repr(e),status='fail'));raise
  write(folder/'validation.json',row);rows.append(row);print(lane,block,shape,'pass',flush=True)
 source_check(plan);write(target/'capture.json',dict(status='pass',children=len(rows),samples=sum(r['samples'] for r in rows),plan=desc(P/'plan.json'),helper=desc(__file__)))

def replay(plan,m,output):
 source_check(plan);rows=[];identities={}
 for lane in ('qualification','native'):
  manifest=read(P/lane/'capture.json');require(manifest['status']=='pass','capture incomplete')
  for block,shape in schedule(plan,lane):
   folder=P/lane/f'b{block:02}-{shape}';cmd=read(folder/'command.json')
   require(cmd['exit_code']==0 and cmd['argv']==argv(plan,folder,shape,lane),'command')
   require(cmd['environment']==plan['environment'] and cmd['cwd']==str(ROOT),'environment')
   require(cmd['helper']==desc(P/'run.py' if lane=='qualification' and block==0 and shape=='medium' else pathlib.Path(__file__)) and cmd['plan']==desc(P/'plan.json'),'command helper/plan')
   for d in cmd['artifacts']:require(desc(d['path'])==d,'artifact custody '+d['path'])
   row=validate(plan,folder,shape,lane,m);require(row==read(folder/'validation.json'),'validation drift')
   identity={k:row[k] for k in ('corpus','constants','sink')}
   if shape in identities:require(identities[shape]==identity,'deterministic semantic/resource drift '+shape)
   identities[shape]=identity;rows.append(dict(row,block=block,report=desc(folder/'report.json')))
 summaries={}
 for shape in plan['shapes']:
  selected=[r for r in rows if r['shape']==shape and r['lane']=='native']
  summaries[shape]={'children':len(selected),'p50_ns_median':statistics.median(r['total']['p50'] for r in selected),'p50_ns_range':[min(r['total']['p50'] for r in selected),max(r['total']['p50'] for r in selected)],'phase_share_ranges':{k:[min(r['phase_shares'][k] for r in selected),max(r['phase_shares'][k] for r in selected)] for k in m.PHASES}}
  values=[r['total']['p50'] for r in selected];summaries[shape]['p50_max_min_ratio']=max(values)/min(values);summaries[shape]['variation_flag']=max(values)/min(values)>plan['variation_flag']['threshold']
 write(output,dict(status='pass',plan=desc(P/'plan.json'),helper=desc(__file__),rows=rows,summaries=summaries,scope=plan['scope'],performance_claim=None))

def main():
 a=argparse.ArgumentParser();a.add_argument('mode',choices=['qualification','native','replay','check']);a.add_argument('--output',type=pathlib.Path,default=P/'analysis.json');args=a.parse_args();plan=read(P/'plan.json');m=load_validator(plan)
 if args.mode in ('qualification','native'):
  if args.mode=='native':require(read(P/'qualification/capture.json')['status']=='pass','qualification required')
  capture(plan,args.mode,m)
 elif args.mode=='replay':replay(plan,m,args.output)
 else:source_check(plan)
if __name__=='__main__':main()
