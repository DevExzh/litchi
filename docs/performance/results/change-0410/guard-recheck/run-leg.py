#!/usr/bin/env python3
"""0410 fixed XLSX normal/guardrail/allocator ABBA capture, one leg per call."""
import argparse,datetime,hashlib,json,os,shutil,subprocess,sys,time
from pathlib import Path
if not __debug__:raise RuntimeError('Assertions must be enabled')
p=argparse.ArgumentParser();p.add_argument('leg',choices=['a1','b1','b2','a2']);args=p.parse_args()
repo=Path('/home/zhuhe/code/litchi');wt=Path('/tmp/litchi-goal-0410-worktree');out=Path('/tmp/litchi-goal-0410-guard-recheck')
role='control' if args.leg.startswith('a') else 'candidate';bins=Path('/tmp/litchi-goal-0410-binaries')/role
identity=json.loads((bins/'identity.json').read_text());rev=identity['revision']
flags='-C force-frame-pointers=yes -C force-unwind-tables=yes'
env=os.environ.copy();env.update(RUSTUP_TOOLCHAIN='1.98.1',RUSTFLAGS=flags,DEBUGINFOD_URLS='')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(n,v):(out/n).write_text(json.dumps(v,indent=2)+'\n')
def git(*a):return subprocess.check_output(['git','-C',str(wt),*a],text=True).strip()
def state():return {'revision':git('rev-parse','HEAD'),'status':git('status','--porcelain=v1')}
commands=json.loads((out/'commands.json').read_text()) if (out/'commands.json').exists() else []
def run(label,argv):
 assert not any(x['label']==label for x in commands),f'{label} already captured'
 started=datetime.datetime.now(datetime.timezone.utc).isoformat();t=time.monotonic()
 with (out/(label+'.stdout')).open('wb') as stdout,(out/(label+'.stderr')).open('wb') as stderr:
  r=subprocess.run([str(x) for x in argv],cwd=wt,env=env,stdout=stdout,stderr=stderr)
 commands.append({'label':label,'argv':[str(x) for x in argv],'cwd':str(wt),'started_utc':started,'wall_seconds':time.monotonic()-t,'exit_code':r.returncode});write('commands.json',commands)
 if r.returncode:raise RuntimeError(f'{label} failed; retained logs')
assert state()['status']==''
run(args.leg+'-checkout',['git','checkout','--detach',rev]);before=state();assert before=={'revision':rev,'status':''}
shutil.copy2(__file__,out/'run-leg.py');shutil.copy2(bins/'identity.json',out/(role+'-build-identity.json'))
write(args.leg+'-environment.json',{'source_before':before,'role':role,'cpu':2,'runtime_environment':{'RUSTFLAGS':flags,'RUSTUP_TOOLCHAIN':'1.98.1','DEBUGINFOD_URLS':''},'build_environment':{'CARGO_TARGET_DIR':'/tmp/litchi-goal-0410-target','CARGO_BUILD_JOBS':'4','CARGO_INCREMENTAL':'0','CARGO_PROFILE_RELEASE_DEBUG':'1','RUSTFLAGS':flags},'protocol':'diagnostic repeat after initial eager guard regression: sequential A1/B1/B2/A2 edit/save guards, same500/20 andcaseorder; initial evidence retained','claim_status':'pending validation and adjudication'})
sys.path.insert(0,str(repo));from tools import perf_compare,perf_abba_summary,validate_perf_corpus_binding
selected='xlsx_file_selected_cell';guards=['xlsx_source_backed_cell_values_one_edit_save','xlsx_eager_cell_values_one_edit_save']
specs=[('guards',guards,500,20,False)]
validation=[]
for kind,cases,samples,warmup,allocator in specs:
 label=args.leg+'-'+kind;name='litchi-perf-baseline-alloc' if allocator else 'litchi-perf-baseline';binary=bins/name
 assert sha(binary)==identity['binaries'][name]['sha256']
 argv=[binary,'--case',','.join(cases),'--filesystem-cache','warm','--xlsx-cell-crud-shape','medium','--samples',samples,'--warmup',warmup,'--json',out/(label+'.json'),'--corpus-manifest',out/(label+'.catalog.json')]
 run(label,['taskset','-c','2','/usr/bin/time','-v','-o',out/(label+'.time.txt'),*argv])
 d=json.loads((out/(label+'.json')).read_text());validate_perf_corpus_binding.validate_paths(out/(label+'.json'),out/(label+'.catalog.json'));perf_compare.validate_parallel_metrics(d)
 assert d['environment']['git_revision']==rev and d['environment']['git_worktree_dirty'] is False
 assert d['environment']['cpu_affinity']=='2' and d['environment']['rustflags']==flags
 assert d['binary_identity']['binary_sha256']==identity['binaries'][name]['sha256']
 assert d['configuration']['cases']==cases and d['configuration']['samples_per_case']==samples and d['configuration']['warmup_iterations_per_case']==warmup
 assert len(d['results'])==len(cases) and {r['case'] for r in d['results']}==set(cases)
 if not allocator:perf_abba_summary._validate_report(d,label,report_role=args.leg)
 for r in d['results']:
  assert len(r['elapsed_ns']['samples'])==samples
  assert r['corpus']['archive_sha256']=='dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036'
  if r['case']==selected:
   perf_compare._validate_operation_metrics(r['operation_metrics'],label,r['elapsed_ns']['samples'],d['schema_version'],elapsed_sample_order=r['elapsed_ns']['sample_order'])
   fs=d['filesystem_evidence'];assert len(fs)==1;fs=fs[0];assert fs['fresh_child_per_sample'] is True and fs['sample_count']==samples
   pids=set()
   for raw in fs['samples']:
    pids.add(raw['child_process_id']);v=raw['xlsx_selected_cell']
    assert v=={'canonical_sheet_name':'Bench01','sheet_position':1,'prepared_selector':'bEnCh01','cell_address':'M29','view':'stored','value_kind':'number','lexical_value':'1028012','digest':'36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1'}
    assert raw['xlsx_source_sha256']==r['corpus']['archive_sha256'];assert raw['xlsx_semantic_sha256']=='020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e'
   assert len(pids)==samples
   if allocator:
    a=r['operation_metrics']['allocation'];assert a['status']=='measured'
    for k in ['allocation_calls','deallocation_calls','reallocation_calls','failed_allocation_calls','allocated_bytes','deallocated_bytes']:
     values=a[k]['values'];assert len(values)==samples and len(set(values))==1,k
     assert values==[raw['allocation_metrics'][k] for raw in fs['samples']],k
    assert a['failed_allocation_calls']['values']==[0]*samples
  else:
   assert r['sink']['accepted_bytes']==4226480 and r['sink']['write_calls']==201 and r['sink']['largest_write']==32768
   if r['case'].startswith('xlsx_source'):
    perf_compare._validate_xlsx_cell_values_range_accounting(d['configuration'],label)
    m=r['source']['xlsx_cell_values']
    for k,value in [('source_read_calls',257),('source_read_bytes',4233005),('workbook_read_bytes',226),('selected_worksheet_read_bytes',6816),('unselected_worksheet_read_bytes',20330),('output_sha256','9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1'),('untouched_member_sha256','7105fcbce160328f666e69fcfd18da9e19fd71dd7b63961e7cddd29d5da1a17d')]:assert m[k]==[value]*samples,k
    for sorted_i,raw_i in enumerate(r['elapsed_ns']['sample_order']):assert sum(m[k][raw_i] for k in ['open_ns','plan_ns','commit_ns','publication_ns'])==r['elapsed_ns']['samples'][sorted_i]
 validation.append({'label':label,'cases':cases,'samples':samples,'warmup':warmup,'status':'passed'})
 assert state()==before and sha(binary)==identity['binaries'][name]['sha256']
write(args.leg+'-validation.json',{'source_after':state(),'reports':validation})
print(args.leg,'validated')
