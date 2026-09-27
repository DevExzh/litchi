"""Offline custody, semantic identity, and diagnostic qualification replay."""
import argparse,subprocess,sys,hashlib,json
from pathlib import Path
import custody as c

def artifact(value,missing=False):
 p=Path(value['path'])
 if not p.is_file():
  assert missing
  cleanup=c.read(c.P/'cleanup.json');assert value in cleanup['removed_binaries'] and cleanup['target_removed']
 else:assert c.artifact(p)==value,p
 return p

def validate(seal=False):
 p=c.P;old=p.parent/'change-0794';plan=c.read(p/'plan.json');origin=c.read(p/'origin.json')
 assert plan['schema']=='litchi.performance.0795.v1' and plan['cpu']==12
 assert plan['owner']=='namespace_uri_probe::capture_region_0793'
 assert plan['callgrind']=={'shapes':['tiny','medium','large'],'shape_orders':[['tiny','medium','large'],['large','medium','tiny']],'leg_orders':[['before','after'],['after','before']],'events':['Ir','Bc','Bcm','Bi','Bim'],'branch_sim':True,'samples':1,'warmup':0,'positive_dumps':1,'termination_empty':True}
 assert plan['perf']=={'shape':'large','orders':[['before','after'],['after','before']],'samples':100,'warmup':3,'event':'cycles:u','frequency':499,'call_graph':'fp','decode_no_inline':True}
 for n,h in c.read(p/'architecture-inputs.json').items():assert c.sha(c.ROOT/n)==h,n
 for n,h in origin['unrelated'].items():assert c.sha(c.ROOT/n)==h,n
 current_worktrees=subprocess.check_output(['git','worktree','list','--porcelain'],cwd=c.ROOT,text=True)
 def worktrees(text):return {block.splitlines()[0][9:]:block for block in text.strip().split('\n\n') if block}
 old_worktrees,new_worktrees=worktrees(origin['worktrees']),worktrees(current_worktrees)
 assert set(old_worktrees)==set(new_worktrees)
 for path,block in old_worktrees.items():
  if Path(path).resolve()!=c.ROOT.resolve():assert new_worktrees[path]==block,path
 assert c.sha(c.ROOT/'Cargo.lock')==c.sha(p/'workspace-Cargo.lock')
 assert (p/'probe-src/Cargo.toml').read_text()==(p/'probe-src/Cargo.toml.template').read_text().replace('@SRC@',str(c.ROOT))
 for a in c.read(p/'inheritance.json')['previous'].values():artifact(a)
 prior_seal=c.read(old/'seal.json')
 for n,h in prior_seal['files'].items():assert c.sha(old/n)==h,n
 baseline=c.read(p/'baseline-source.json');assert baseline['files']==c.read(old/'build-before/source.json')['files']==c.source()['files']
 assert baseline['revision']==origin['base']
 manifest=c.read(old/'candidate/manifest.json');candidate=baseline['files'].copy()
 assert [r['production'] for r in manifest['files']]==plan['source_allowlist']
 for r in manifest['files']:
  assert c.sha(old/'candidate'/r['before'])==r['before_sha256']==baseline['files'][r['production']]
  assert c.sha(old/'candidate'/r['after'])==r['after_sha256'];candidate[r['production']]=r['after_sha256']
 assert candidate==c.read(old/'build-after/source.json')['files']
 for n in ['quality-before.json','quality.json']:
  q=c.read(old/n);assert c.read(artifact(q['source']))['files']==(baseline['files'] if n=='quality-before.json' else candidate);assert len(q['rows'])==6
  for row in q['rows']:assert row['exit_code']==0;artifact(row['log'])
 builds={}
 for leg,expected in [('before',baseline['files']),('after',candidate)]:
  b=c.read(p/f'build-{leg}/build.json');builds[leg]=b
  assert c.read(artifact(b['source']))['files']==expected
  assert set(b['binaries'])=={'profile','fp'} and len(b['rows'])==2
  for n,h in b['frozen'].items():assert c.sha(p/n)==h,n
  assert b['probe']==c.read(p/'inheritance.json')['probe']
  for n,h in b['probe'].items():assert c.sha(p/'probe-src'/n)==h,n
  for row,kind in zip(b['rows'],['profile','fp']):
   assert row['kind']==kind and row['exit_code']==0;artifact(row['log'])
   assert row['command']==['cargo','build','--offline','--locked','--release','--manifest-path',str(p/'probe-src/Cargo.toml'),'--features','capture-profile']
   assert row['environment']=={'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','RUSTFLAGS':'' if kind=='profile' else '-C force-frame-pointers=yes'}
   artifact(b['binaries'][kind],missing=True)
 assert builds['before']['rows'][-1]['ended']<=builds['after']['rows'][0]['started']
 assert all(builds['before']['binaries'][k]['sha256']!=builds['after']['binaries'][k]['sha256'] for k in ['profile','fp'])
 application=c.read(p/'application.json');artifact(application['manifest']);assert c.read(artifact(application['source']))['files']==candidate
 stale=c.read(p/'stale-build.json');failed=c.read(artifact(stale['failed_build']))
 assert stale['source_bytes_unchanged'] is True
 assert stale['removed_binary_copies']==list(failed['binaries'].values())
 for kind,row in zip(['profile','fp'],failed['rows']):
  assert failed['binaries'][kind]['sha256']==builds['before']['binaries'][kind]['sha256']
  assert row['exit_code']==0
  log=dict(row['log']);log['path']=str(p/'build-after-stale'/Path(log['path']).name);artifact(log)
 for name,item in stale['source_mtimes_before_refresh'].items():
  assert item['sha256']==candidate[name] and item['mtime_ns']/1e9<builds['before']['rows'][0]['started']
 assert builds['before']['rows'][-1]['ended']<=application['completed']<=failed['rows'][0]['started']
 assert failed['rows'][-1]['ended']<=stale['timestamp_refresh_completed']<=builds['after']['rows'][0]['started']
 restoration=c.read(p/'restoration.json');assert c.read(artifact(restoration['source']))['files']==baseline['files']
 assert restoration['completed']>=builds['after']['rows'][-1]['ended']
 expected_profile=[(r,shape,leg) for r,shapes in enumerate(plan['callgrind']['shape_orders']) for shape in shapes for leg in plan['callgrind']['leg_orders'][r]]
 expected_perf=[(r,'large',leg) for r,order in enumerate(plan['perf']['orders']) for leg in order]
 count=0;last=restoration['completed']
 for lane,expected,kind,samples,warmup in [('profiles',expected_profile,'profile',1,0),('perf',expected_perf,'fp',100,3)]:
  complete=c.read(p/lane/'complete.json');rows=c.read(artifact(complete['receipts']));assert len(rows)==complete['children']==len(expected)
  assert c.read(artifact(complete['source']))['files']==baseline['files']
  for row,(repeat,shape,leg) in zip(rows,expected):
   assert row['repeat']==repeat and row['leg']==leg and row.get('shape',shape)==shape
   assert row['exit_code']==0 and last<=row['started']<=row['ended'];last=row['ended']
   assert row['binary']==builds[leg]['binaries'][kind]
   if lane=='profiles':
    stem=f'{repeat}-{shape}-{leg}';art=row['artifacts'];assert set(art)=={stem+ext for ext in ['.json','.log','.callgrind','.callgrind.1']}
    for value in art.values():artifact(value)
    report=artifact(art[stem+'.json']);raw=p/lane/(stem+'.callgrind');owner=plan['owner']
    command=['taskset','-c','12','valgrind','--tool=callgrind','--branch-sim=yes','--collect-atstart=no','--toggle-collect='+owner,'--zero-before='+owner,'--dump-after='+owner,'--callgrind-out-file='+str(raw),row['binary']['path'],'--mode','capture','--shape',shape,'--samples','1','--warmup','0','--output',str(report)]
   else:
    report=artifact(row['report']);artifact(row['log']);artifact(row['compressed']);raw=p/lane/f'{repeat}-{leg}.data'
    command=['taskset','-c','12','perf','record','-e','cycles:u','-F','499','--call-graph','fp','-o',str(raw),'--',row['binary']['path'],'--mode','capture','--shape','large','--samples','100','--warmup','3','--output',str(report)]
   assert row['command']==command
   data=c.read(report);ref=c.read(old/f'qualification/0-{shape}-capture-before.json')
   for n in ['schema','tool','mode','shape','slides','shapes_per_slide','timing_scope','marker','source','fixture']:assert data[n]==ref[n],n
   assert data['allocator']['instrumentation']=='none' and data['allocator']['counter_revision'] is None
   assert data['warmup']==warmup and data['samples_requested']==len(data['samples'])==samples
   for i,sample in enumerate(data['samples']):
    assert sample['index']==i and sample['elapsed_ns']>0 and 'allocation' not in sample
    assert sample['metrics']['elapsed_ns']==sample['elapsed_ns']
    for n in ['captured_shapes_per_slide','captured_slides','shapes_per_slide','slides']:assert sample['metrics'][n]==ref['samples'][0]['metrics'][n]
    for n in ['source_sha256','output','verification']:assert sample[n]==ref['samples'][0][n],n
   count+=samples
 assert count==412
 tests=c.read(p/'parser-tests.json')
 assert tests['command']==['python3','-B',str(p/'test_callgrind_parser.py')] and tests['exit_code']==0 and tests['tests']==3
 text=artifact(tests['log']).read_text();assert 'Ran 3 tests' in text and text.rstrip().endswith('OK')
 artifact(tests['test_source']);artifact(tests['parser'])
 for script in ['cg_analysis.py','native_analysis.py','root_cg_totals.py','root_native_counts.py']:
  subprocess.run([sys.executable,'-B',str(p/script),'--check'],check=True)
 if (p/'cleanup.json').exists():
  clean=c.read(p/'cleanup.json');assert clean['target']==str(c.TARGET) and clean['target_removed'] and not c.TARGET.exists()
  assert clean['removed_binaries']==[b for leg in ['before','after'] for b in builds[leg]['binaries'].values()]
  assert clean['removed_target_bytes']>0
 assert not list(p.rglob('__pycache__'))
 if seal:
  subprocess.run([sys.executable,'-B',str(p/'seal_packet.py')],check=True)
 print(json.dumps({'reports':16,'samples':count,'production_change':False,'native_latency_claim':False,'seal_checked':seal},sort_keys=True))

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--require-final-seal',action='store_true');args=parser.parse_args();validate(args.require_final_seal)
