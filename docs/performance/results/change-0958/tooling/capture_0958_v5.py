#!/usr/bin/env python3
"""Inert-until-GO whole-process heaptrack capture for the 0958 diagnostic.

The driver adapts the retained 0949 custody pattern to the retained 0957
uninstrumented benchmark.  It validates the source, binary, tool, schedule,
and report contract before launch.  Each child writes a start receipt first and
a terminal receipt after process-group reaping and profile custody.  The first
child is admitted by a post-terminal heaptrack_print symbol check; a failed
admission stops the fixed schedule without retrying it.  No timing, RSS, or
adoption decision is made here.
"""
from __future__ import annotations
import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from typing import Any

ROOT=Path(__file__).resolve().parents[1]
PLAN_PATH=ROOT/'plan-v5.json'
SCHEMA='litchi.performance.0958.heaptrack-allocation-attribution.v3'
CHILD_SCHEMA='litchi.performance.0958.heaptrack-child-terminal.v3'
RUNNER_SCHEMA='litchi.performance.0958.heaptrack-runner-terminal.v3'
ADMISSION_SCHEMA='litchi.performance.0958.heaptrack-first-admission-process-terminal.v3'
ADMISSION_RESULT_SCHEMA='litchi.performance.0958.heaptrack-first-admission-result.v3'
REPORT_ORACLE_SCHEMA='litchi.performance.0958.heaptrack-report-oracle.v3'
MAX_JSON_BYTES=64*1024*1024
FIXED_ENV={'PATH':'/usr/bin:/bin','LANG':'C','LC_ALL':'C','RUST_BACKTRACE':'0','TZ':'UTC'}
RELEVANT_SYMBOLS=(
 'litchi_xlsx::cell::Store::merge_omitted_cells',
 'litchi_xlsx::workbook::edit::semantic::transaction::Edit::commit',
 'litchi_perf_baseline::xlsx_managed_compression::run_case',
 'stage_timed_commit',
)

class CaptureError(RuntimeError): pass
def require(c:bool,m:str)->None:
    if not c: raise CaptureError(m)
def utc_now()->str: return dt.datetime.now(dt.timezone.utc).isoformat(timespec='microseconds')
def canonical(v:Any)->bytes: return json.dumps(v,sort_keys=True,separators=(',',':'),ensure_ascii=False,allow_nan=False).encode()
def sha_bytes(v:bytes)->str: return hashlib.sha256(v).hexdigest()
def sha_file(p:Path)->str:
    h=hashlib.sha256()
    with p.open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
    return h.hexdigest()
def read_json(p:Path)->Any:
    require(p.is_file() and not p.is_symlink(),f'not a regular JSON file: {p}')
    require(p.stat().st_size<=MAX_JSON_BYTES,f'JSON exceeds limit: {p}')
    def nodup(pairs:list[tuple[str,Any]])->dict[str,Any]:
        out={}
        for k,v in pairs:
            require(k not in out,f'duplicate JSON key {k!r} in {p}'); out[k]=v
        return out
    try: return json.loads(p.read_text(encoding='utf-8'),object_pairs_hook=nodup)
    except (OSError,UnicodeError,json.JSONDecodeError) as e: raise CaptureError(f'cannot read JSON {p}: {e}') from e
def write_once(p:Path,v:Any)->None:
    require(not p.exists() and not p.is_symlink(),f'refusing retained artifact: {p}')
    p.parent.mkdir(parents=True,exist_ok=True)
    data=(json.dumps(v,indent=2,sort_keys=True,ensure_ascii=False,allow_nan=False)+'\n').encode()
    with p.open('xb') as f: f.write(data); f.flush(); os.fsync(f.fileno())
    fd=os.open(p.parent,os.O_RDONLY|os.O_DIRECTORY)
    try: os.fsync(fd)
    finally: os.close(fd)
def descriptor(p:Path, executable:bool|None=None)->dict[str,Any]:
    p=p.expanduser().resolve(); require(p.is_absolute() and p.is_file() and not p.is_symlink(),f'bad artifact: {p}')
    st=p.stat(); is_exec=bool(st.st_mode&0o111)
    if executable is not None: require(is_exec is executable,f'executable mismatch: {p}')
    return {'path':str(p),'bytes':st.st_size,'mode':st.st_mode&0o777,'executable':is_exec,'sha256':sha_file(p)}
def verify_desc(value:Any,label:str,executable:bool|None=None)->dict[str,Any]:
    require(isinstance(value,dict),f'{label} must be descriptor')
    raw=value.get('path'); require(isinstance(raw,str) and Path(raw).is_absolute(),f'{label}.path not absolute')
    actual=descriptor(Path(raw),executable)
    for key in ('bytes','mode','executable','sha256'):
        if key in value: require(actual[key]==value[key],f'{label}.{key} mismatch')
    return actual
def ancestry()->set[int]:
    out={os.getpid()}; cur=os.getpid()
    while True:
        try: parent=int(Path(f'/proc/{cur}/stat').read_text().split()[3])
        except (OSError,ValueError,IndexError): break
        if parent<=1 or parent in out: break
        out.add(parent); cur=parent
    return out
def process_probe()->list[dict[str,Any]]:
    skip=ancestry(); bad=[]
    for e in Path('/proc').iterdir():
        if not e.name.isdigit() or int(e.name) in skip: continue
        try: comm=(e/'comm').read_text().strip(); args=(e/'cmdline').read_bytes().replace(b'\0',b' ').decode(errors='replace').strip()
        except OSError: continue
        if comm in {'cargo','rustc','rustdoc','perf','heaptrack','heaptrack_gui'} or any(x in args for x in ('xlsx_managed_compression_benchmark','litchi-perf-baseline')):
            bad.append({'pid':int(e.name),'comm':comm,'cmdline':args})
    return bad

def source_tree(root:Path)->dict[str,dict[str,Any]]:
    require(root.is_dir() and not root.is_symlink(),f'source root invalid: {root}')
    out={}
    for dp,dns,fns in os.walk(root,topdown=True,followlinks=False):
        dns.sort(); fns.sort(); dp=Path(dp)
        for name in dns: require(not (dp/name).is_symlink(),f'source dir symlink: {dp/name}')
        for name in fns:
            p=dp/name; require(p.is_file() and not p.is_symlink(),f'source nonregular: {p}')
            st=p.stat(); rel=p.relative_to(root).as_posix(); require(rel not in out,f'duplicate source path {rel}')
            out[rel]={'path':rel,'bytes':st.st_size,'mode':st.st_mode&0o777,'sha256':sha_file(p)}
    return out

def validate_source(binding:dict[str,Any])->dict[str,Any]:
    raw_root=binding.get('root'); require(isinstance(raw_root,str) and Path(raw_root).is_absolute(),'source.root invalid')
    root=Path(raw_root); manifest_desc=verify_desc(binding.get('manifest'),'source manifest')
    manifest=read_json(Path(manifest_desc['path']))
    require(manifest.get('schema')=='litchi.performance.0957.source-manifest.v2','source manifest schema mismatch')
    require(manifest.get('source_root')=='source-v1','source manifest root mismatch')
    wanted={}
    for item in manifest.get('files',[]):
        require(isinstance(item,dict) and isinstance(item.get('path'),str),'malformed source manifest item')
        name=item['path']; q=Path(name)
        require(name and not q.is_absolute() and q.as_posix()==name and '..' not in q.parts,'unnormalized source manifest path')
        require(name not in wanted,'duplicate source manifest path')
        wanted[name]=item
    actual=source_tree(root)
    require(set(actual)==set(wanted),f'source census differs: missing={sorted(set(wanted)-set(actual))[:3]} extra={sorted(set(actual)-set(wanted))[:3]}')
    for name,item in wanted.items():
        for k in ('bytes','mode','sha256'): require(actual[name][k]==item[k],f'source mismatch at {name} field {k}')
    require(len(actual)==manifest.get('file_count') and sum(x['bytes'] for x in actual.values())==manifest.get('bytes'),'source census totals mismatch')
    post=verify_desc(binding.get('postquality_receipt'),'source postquality receipt')
    post_raw=read_json(Path(post['path']))
    require(post_raw.get('status')=='source_manifest_match_after_build_and_quality','source postquality status mismatch')
    require(post_raw.get('integration_manifest',{}).get('sha256')==manifest_desc['sha256'],'postquality manifest binding mismatch')
    freeze=verify_desc(binding.get('integration_freeze'),'integration freeze')
    root_start=verify_desc(binding.get('root_start'),'0958 root start')
    return {'root':str(root),'manifest':manifest_desc,'postquality_receipt':post,'integration_freeze':freeze,'root_start':root_start,'file_count':len(actual),'bytes':sum(x['bytes'] for x in actual.values()),'base_commit':manifest.get('base_commit')}

def validate_binary(binding:dict[str,Any],source:dict[str,Any])->dict[str,Any]:
    binary=verify_desc(binding.get('binary'),'retained benchmark binary',True)
    retention=verify_desc(binding.get('retention_receipt'),'executable retention receipt')
    raw=read_json(Path(retention['path'])); require(raw.get('status')=='retained_byte_verified_executables','retention status mismatch')
    matches=[x for x in raw.get('executables',[]) if isinstance(x,dict) and x.get('role')=='benchmark_example']
    require(len(matches)==1,'retention lacks unique benchmark example')
    wanted=matches[0].get('retained_executable',{})
    raw_path=wanted.get('path')
    require(isinstance(raw_path,str),'retained executable path missing')
    # Retention receipts are rooted at the 0957 external root and use paths
    # relative to that root; the plan binds the independently retained file
    # by its absolute descriptor.
    receipt_root=Path(retention['path']).parent.parent
    expected=(receipt_root/Path(raw_path)).resolve() if not Path(raw_path).is_absolute() else Path(raw_path).resolve()
    require(str(expected)==binary['path'] and wanted.get('bytes')==binary['bytes'] and wanted.get('sha256')==binary['sha256'],'retained binary does not match retention')
    build=verify_desc(binding.get('build_receipt'),'0957 integration build receipt')
    build_raw=read_json(Path(build['path'])); require(build_raw.get('status')=='passed','0957 build receipt not passed')
    require(build_raw.get('source_commit')==source['base_commit'],'build receipt source commit differs from validated source')
    build_source=build_raw.get('source'); require(isinstance(build_source,dict),'build receipt source binding missing')
    require(build_source.get('source_commit')==source['base_commit'],'build receipt nested source commit differs from validated source')
    require(build_source.get('file_count')==source['file_count'] and build_source.get('bytes')==source['bytes'],'build receipt source census differs from validated source')
    build_manifest=build_source.get('manifest'); require(isinstance(build_manifest,dict),'build receipt source manifest binding missing')
    require(build_manifest.get('bytes')==source['manifest']['bytes'] and build_manifest.get('sha256')==source['manifest']['sha256'],'build receipt manifest differs from validated source manifest')
    return {'binary':binary,'retention_receipt':retention,'build_receipt':build}

def validate_tool(name:str,value:dict[str,Any],version:str)->dict[str,Any]:
    tool=verify_desc(value,f'tool {name}',True)
    command=[tool['path'],'--version']
    try: proc=subprocess.run(command,env=dict(FIXED_ENV),stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.PIPE,timeout=15,text=True,check=False)
    except OSError as e: raise CaptureError(f'tool {name} version probe failed: {e}') from e
    output=(proc.stdout+'\n'+proc.stderr).strip(); require(proc.returncode==0,f'tool {name} version rc={proc.returncode}')
    require(version in output,f'tool {name} version does not contain {version!r}: {output[:200]!r}')
    return {**tool,'version':version,'version_argv':command,'version_output_sha256':sha_bytes(output.encode())}

def validate_plan()->tuple[dict[str,Any],dict[str,Any]]:
    plan=read_json(PLAN_PATH)
    require(plan.get('schema')==SCHEMA and plan.get('plan_id')=='0958' and plan.get('plan_revision')=='v5','plan identity mismatch')
    require(plan.get('status') in {'prospective_no_capture','capture_ready'},'plan is not prospective')
    require(plan.get('capture_authorization')=='root_GO_required','plan capture authorization missing')
    require(plan.get('claim_scope',{}).get('whole_process_only') is True,'whole-process claim scope missing')
    require(plan.get('claim_scope',{}).get('native_latency_or_rss_claims') is False,'native/RSS exclusion missing')
    require(plan.get('environment',{}).get('fixed')==FIXED_ENV,'fixed environment changed')
    matrix=plan.get('matrix',{}); require(matrix.get('children')==6 and matrix.get('serial') is True and matrix.get('workers')==1 and matrix.get('warmup')==0 and matrix.get('samples')==3 and matrix.get('cpu')==12 and matrix.get('deadline_seconds')==900 and matrix.get('no_retry') is True,'matrix changed')
    expected=[('h00-medium-percent-r1','medium','one-percent'),('h01-dense-percent-r1','dense-sparse','one-percent'),('h02-medium-one-edit-r1','medium','one-edit'),('h03-medium-one-edit-r2','medium','one-edit'),('h04-dense-percent-r2','dense-sparse','one-percent'),('h05-medium-percent-r2','medium','one-percent')]
    jobs=matrix.get('jobs'); require(isinstance(jobs,list) and len(jobs)==6,'matrix jobs count changed')
    for item,(name,shape,workload) in zip(jobs,expected):
        require(item.get('name')==name and item.get('shape')==shape and item.get('workload')==workload and item.get('workers')==1 and item.get('warmup')==0 and item.get('samples')==3,'matrix job changed')
    source=validate_source(plan['source']); binary=validate_binary(plan['binary'],source)
    tools={name:validate_tool(name,plan['tools'][name],plan['tool_versions'][name]) for name in ('heaptrack','heaptrack_print','taskset','time')}
    runner=verify_desc(plan.get('runner'),'capture runner')
    script_path=Path(__file__).resolve()
    require(not Path(__file__).is_symlink(),'executing runner must not be a symlink')
    require(script_path==Path(runner['path']).resolve(),'executing runner path differs from plan.runner')
    require(sha_file(script_path)==runner['sha256'],'executing runner hash differs from plan.runner')
    committed=verify_desc(plan.get('committed_0949_doc'),'committed 0949 attribution doc')
    fixture=verify_desc(plan.get('fixture_oracle'),'0956 fixture oracle')
    fixture_raw=read_json(Path(fixture['path']))
    require(isinstance(fixture_raw,dict) and all(k in fixture_raw for k in ('medium-one-percent','dense-sparse-one-percent','medium-one-edit')),'fixture oracle cases missing')
    for key in ('medium-one-percent','dense-sparse-one-percent','medium-one-edit'):
        case=fixture_raw[key]; require(isinstance(case,dict) and isinstance(case.get('invariant'),dict) and isinstance(case.get('by_workers'),dict) and isinstance(case['by_workers'].get('1'),dict),'fixture oracle worker-one binding missing')
    admission=plan.get('first_child_admission',{}); require(admission.get('child')=='h00-medium-percent-r1' and admission.get('required_any_symbol'),'first admission binding missing')
    capture_root=plan.get('capture_root'); admission_root=plan.get('admission_root')
    require(isinstance(capture_root,str) and Path(capture_root).is_absolute(),'capture_root invalid')
    require(isinstance(admission_root,str) and Path(admission_root).is_absolute(),'admission_root invalid')
    require(not Path(capture_root).exists() and not Path(admission_root).exists(),'capture/admission root already exists; no retry')
    return plan,{'plan':descriptor(PLAN_PATH),'runner':runner,'source':source,'binary':binary,'tools':tools,'committed_0949_doc':committed,'fixture_oracle':fixture,'fixture_raw':fixture_raw,'jobs':[x['name'] for x in jobs],'capture_root':capture_root,'admission_root':admission_root,'workload_not_started':True}

def group_clear(pgid:int|None)->bool:
    if pgid is None:return True
    try: os.killpg(pgid,0)
    except ProcessLookupError:return True
    except PermissionError:return False
    return False
def wait_clear(pgid:int|None,seconds:float)->bool:
    end=time.monotonic()+seconds
    while time.monotonic()<end:
        if group_clear(pgid):return True
        time.sleep(.05)
    return group_clear(pgid)
def stop_group(pgid:int|None,proc:subprocess.Popen[bytes])->dict[str,Any]:
    actions=[]
    if pgid is not None:
        try: os.killpg(pgid,signal.SIGTERM); actions.append('SIGTERM')
        except ProcessLookupError: pass
    try: proc.wait(timeout=15)
    except subprocess.TimeoutExpired:
        if pgid is not None:
            try: os.killpg(pgid,signal.SIGKILL); actions.append('SIGKILL')
            except ProcessLookupError: pass
        proc.wait()
    clear=wait_clear(pgid,30)
    return {'actions':actions,'process_group_clear':clear,'reaped':True}
def profiles(folder:Path)->list[Path]:
    return [p for p in sorted(folder.iterdir(),key=lambda x:x.name) if p.name in {'heaptrack.data.gz','heaptrack.data.gz.zst'} and p.is_file() and not p.is_symlink()]
def child_command(job:dict[str,Any],plan:dict[str,Any],folder:Path)->list[str]:
    return [plan['tools']['time']['path'],'-v','-o',str(folder/'time.txt'),plan['tools']['taskset']['path'],'--cpu-list','12',plan['tools']['heaptrack']['path'],'--record-only','-o',str(folder/'heaptrack.data.gz'),plan['binary']['binary']['path'],'--shape',job['shape'],'--workload',job['workload'],'--workers','1','--warmup','0','--samples','3','--output',str(folder/'workload')]
def child_artifacts(folder:Path)->dict[str,Any]:
    candidates=profiles(folder)
    workload=folder/'workload'; report=workload/'report.json'
    return {'start':'start.json','terminal':'terminal.json','stdout':'stdout.log','stderr':'stderr.log','time':'time.txt','workload_dir':'workload','report':'workload/report.json','profile_requested':'heaptrack.data.gz','profile_candidates':[p.name for p in candidates],'profile_actual':descriptor(candidates[0]) if len(candidates)==1 else None,'report_descriptor':descriptor(report) if report.is_file() else None}

def run_process(argv:list[str],cwd:Path,env:dict[str,str],stdout_path:Path,stderr_path:Path,deadline:int)->dict[str,Any]:
    proc=None; pid=None; pgid=None; rc=None; timed=False; launch_error=None; cleanup={'actions':[],'process_group_clear':True,'reaped':False}; started=utc_now(); launched=utc_now()
    try:
        with stdout_path.open('xb') as out, stderr_path.open('xb') as err:
            try:
                proc=subprocess.Popen(argv,cwd=cwd,env=env,stdin=subprocess.DEVNULL,stdout=out,stderr=err,start_new_session=True)
                pid=proc.pid; pgid=pid; launched=utc_now()
                try: rc=proc.wait(timeout=deadline)
                except subprocess.TimeoutExpired:
                    timed=True; cleanup=stop_group(pgid,proc); rc=proc.returncode
                if not cleanup['reaped']:
                    rc=proc.wait(); cleanup={'actions':[],'process_group_clear':wait_clear(pgid,30),'reaped':True}
            except BaseException as e:
                launch_error=f'{type(e).__name__}: {e}'
                if proc is not None:
                    cleanup=stop_group(pgid,proc); rc=proc.returncode
    except BaseException as e:
        launch_error=launch_error or f'{type(e).__name__}: {e}'
        if proc is not None and not cleanup['reaped']:
            cleanup=stop_group(pgid,proc); rc=proc.returncode
    return {'started_utc':started,'launched_utc':launched,'terminal_utc':utc_now(),'pid':pid,'process_group_id':pgid,'returncode':rc,'timed_out':timed,'launch_error':launch_error,'cleanup':cleanup}

def report_oracle(job:dict[str,Any],report_path:Path,fixture:dict[str,Any])->dict[str,Any]:
    raw=read_json(report_path)
    require(raw.get('schema')=='litchi.xlsx-managed-compression-benchmark.v1' and raw.get('status')=='completed','benchmark report status/schema invalid')
    for key,wanted in [('shape',job['shape']),('workload',job['workload']),('workers',1),('warmup_iterations',0),('measured_samples',3)]: require(raw.get(key)==wanted,f'report {key} mismatch')
    require(raw.get('performance_claim')=='none; raw phase timings describe this explicit in-memory harness and do not claim OS thread counts, speedup, allocation behavior, or physical I/O','report performance claim changed')
    case=raw.get('case'); require(isinstance(case,dict),'report case missing')
    fixture_key=f"{job['shape']}-{job['workload']}"; expected=fixture.get(fixture_key); require(isinstance(expected,dict),'fixture case missing')
    invariant=expected.get('invariant'); worker=expected.get('by_workers',{}).get('1',{}); oracle=worker.get('oracle')
    require(isinstance(invariant,dict) and isinstance(oracle,dict),'fixture oracle malformed')
    for key in ('shape','workload','workers'): require(case.get(key)==job.get(key) if key != 'workers' else case.get(key)==1,f'case {key} mismatch')
    exact_case=(
        ('update_count',invariant.get('update_count')),('source_bytes',invariant.get('source_bytes')),
        ('source_sha256',invariant.get('source_sha256')),('output_bytes',invariant.get('output_bytes')),
        ('output_sha256',invariant.get('output_sha256')),('semantic_sha256',invariant.get('semantic_sha256')),
        ('expected_cpu_tasks',invariant.get('expected_cpu_tasks')),
        ('sink_capacity_bytes',invariant.get('sink_capacity_bytes')),('sink_max_write_bytes',invariant.get('sink_max_write_bytes')),
        ('admission_expected_worker_permits',oracle.get('admission_expected_worker_permits')),
        ('admission_observed_worker_permits',oracle.get('admission_observed_worker_permits')),
        ('admission_cpu_tasks_used',oracle.get('admission_cpu_tasks_used')),
        ('context_limits',oracle.get('context_limits')),
    )
    for key,wanted in exact_case: require(case.get(key)==wanted,f'case {key} differs from retained fixture oracle')
    require(case.get('warmup_iterations')==0 and case.get('measured_samples')==3,'case sample contract mismatch')
    require(case.get('preflight_actual_commit_inverse_verified') is True,'report inverse preflight oracle failed')
    samples=case.get('samples'); require(isinstance(samples,list) and len(samples)==3,'report sample count mismatch')
    resources={'memory':0,'workers':0,'io_concurrency':0,'objects':0}
    for index,sample in enumerate(samples):
        require(isinstance(sample,dict) and sample.get('sample')==index,f'sample {index} identity mismatch')
        for key,wanted in (
            ('output_bytes',invariant.get('output_bytes')),('output_bytes_used',invariant.get('output_bytes')),
            ('cpu_tasks_used',invariant.get('expected_cpu_tasks')),('output_sha256',invariant.get('output_sha256')),
            ('semantic_sha256',invariant.get('semantic_sha256')),
        ): require(sample.get(key)==wanted,f'sample {index} {key} differs from retained fixture oracle')
        require(sample.get('resources_after_drop')==resources,f'sample {index} resources-after-drop mismatch')
        require(sample.get('sink_largest_write')==invariant.get('sink_max_write_bytes'),'sample sink largest write differs from retained bound')
        require(isinstance(sample.get('sink_write_calls'),int) and sample['sink_write_calls']>0,'sample sink write call count invalid')
        require(sample.get('fixture_inverse_restored_exact_source') is True,f'sample {index} inverse oracle failed')
    return {'schema':REPORT_ORACLE_SCHEMA,'status':'pass','fixture_case':fixture_key,'report':descriptor(report_path),'shape':raw['shape'],'workload':raw['workload'],'workers':raw['workers'],'warmup_iterations':raw['warmup_iterations'],'samples':len(samples),'case':{'update_count':case['update_count'],'source_bytes':case['source_bytes'],'source_sha256':case['source_sha256'],'output_bytes':case['output_bytes'],'output_sha256':case['output_sha256'],'semantic_sha256':case['semantic_sha256'],'expected_cpu_tasks':case['expected_cpu_tasks'],'context_limits':case['context_limits'],'sink_capacity_bytes':case['sink_capacity_bytes'],'sink_max_write_bytes':case['sink_max_write_bytes']}}

def admission_print(plan:dict[str,Any],validation:dict[str,Any],job:dict[str,Any],folder:Path,profile:Path,report_receipt:dict[str,Any])->dict[str,Any]:
    root=Path(plan['admission_root']); child=root/'first-child'; require(not root.exists() and not root.is_symlink(),f'admission root retained: {root}'); child.mkdir(parents=True)
    stdout=child/'stdout.log'; stderr=child/'stderr.log'; hist=child/'allocation-size.histogram'; folded=child/'allocation-stacks.folded'
    argv=[plan['tools']['heaptrack_print']['path'],'-f',str(profile),'-m','1','-n','1000','-s','1000','-a','-p','-T','-l','--disable-builtin-suppressions','--disable-embedded-suppressions','-H',str(hist),'-F',str(folded),'--flamegraph-cost-type','allocations']
    start={'schema':ADMISSION_SCHEMA,'status':'started','plan':validation['plan'],'runner':validation['runner'],'child':job['name'],'argv':argv,'cwd':validation['source']['root'],'environment':FIXED_ENV,'profile':descriptor(profile),'started_utc':utc_now(),'deadline_seconds':plan['first_child_admission']['deadline_seconds']}
    write_once(child/'start.json',start)
    run=run_process(argv,Path(validation['source']['root']),FIXED_ENV,stdout,stderr,plan['first_child_admission']['deadline_seconds'])
    process_ok=run['launch_error'] is None and run['returncode']==0 and not run['timed_out'] and run['cleanup']['process_group_clear'] and run['cleanup']['reaped']
    terminal={'schema':ADMISSION_SCHEMA,'status':'terminal','child':job['name'],'argv':argv,'cwd':validation['source']['root'],'environment':FIXED_ENV,'profile':descriptor(profile),'run':run,'stdout':descriptor(stdout) if stdout.is_file() else None,'stderr':descriptor(stderr) if stderr.is_file() else None,'generated_before_symbol_parse':[descriptor(p) for p in sorted(child.iterdir(),key=lambda p:p.name) if p.is_file() and p.name not in {'start.json','stdout.log','stderr.log'}],'process_terminal_precedes_symbol_parse':True,'process_ok':process_ok}
    # This receipt is durable before stdout is read or symbols are classified.
    write_once(child/'terminal.json',terminal)
    parse_error=None
    try:
        output=stdout.read_text(errors='replace') if stdout.is_file() else ''
        symbols=[s for s in RELEVANT_SYMBOLS if s in output]
    except BaseException as e:
        output=''; symbols=[]; parse_error=f'{type(e).__name__}: {e}'
    accepted=process_ok and parse_error is None and bool(symbols)
    result={'schema':ADMISSION_RESULT_SCHEMA,'status':'pass' if accepted else 'fail','child':job['name'],'process_terminal':descriptor(child/'terminal.json'),'report_oracle':report_receipt,'required_any_symbol':RELEVANT_SYMBOLS,'symbols_found':symbols,'number_or_box_clone_required':False,'process_terminal_precedes_symbol_parse':True,'symbol_parse_after_terminal':True,'parse_error':parse_error,'generated_after_symbol_parse':[descriptor(p) for p in sorted(child.iterdir(),key=lambda p:p.name) if p.is_file() and p.name not in {'start.json','terminal.json','stdout.log','stderr.log'}],'accepted':accepted}
    write_once(child/'result.json',result)
    return result

def run_child(plan:dict[str,Any],validation:dict[str,Any],job:dict[str,Any],capture:Path)->dict[str,Any]:
    folder=capture/'children'/job['name']; require(not folder.exists() and not folder.is_symlink(),f'refusing retained child: {folder}'); folder.mkdir(parents=True)
    argv=child_command(job,plan,folder)
    start={'schema':CHILD_SCHEMA,'status':'started','plan':validation['plan'],'runner':validation['runner'],'id':job['name'],'shape':job['shape'],'workload':job['workload'],'workers':1,'warmup':0,'samples':3,'argv':argv,'cwd':validation['source']['root'],'environment':FIXED_ENV,'binary':validation['binary']['binary'],'source':validation['source'],'started_utc':utc_now(),'deadline_seconds':900,'start_new_session':True,'process_group':True,'no_retry':True}
    write_once(folder/'start.json',start); run=run_process(argv,Path(validation['source']['root']),FIXED_ENV,folder/'stdout.log',folder/'stderr.log',900)
    arts=child_artifacts(folder); passed=run['launch_error'] is None and run['returncode']==0 and not run['timed_out'] and run['cleanup']['process_group_clear'] and run['cleanup']['reaped'] and arts['profile_actual'] is not None and arts['report_descriptor'] is not None
    terminal={'schema':CHILD_SCHEMA,'status':'pass' if passed else 'fail','plan':validation['plan'],'runner':validation['runner'],'id':job['name'],'shape':job['shape'],'workload':job['workload'],'workers':1,'warmup':0,'samples':3,'argv':argv,'cwd':validation['source']['root'],'environment':FIXED_ENV,'binary':validation['binary']['binary'],'source':validation['source'],'run':run,'terminal_before_report_or_profile_parse':True,'artifacts':arts,'analysis_started':False,'whole_process_profile':True}
    write_once(folder/'terminal.json',terminal)
    return terminal

def persist_report_oracle(plan:dict[str,Any],validation:dict[str,Any],job:dict[str,Any],terminal:dict[str,Any],capture:Path)->dict[str,Any]:
    folder=capture/'children'/job['name']; report=folder/'workload/report.json'; terminal_path=folder/'terminal.json'; receipt_path=folder/'report-oracle.json'
    base={'schema':REPORT_ORACLE_SCHEMA,'child':job['name'],'terminal':descriptor(terminal_path),'report':descriptor(report) if report.is_file() else None}
    try:
        oracle=report_oracle(job,report,validation['fixture_raw'])
        receipt={**base,'status':'pass','oracle':oracle}
    except BaseException as e:
        receipt={**base,'status':'fail','error':f'{type(e).__name__}: {e}'}
    write_once(receipt_path,receipt)
    return receipt

def run_capture(plan:dict[str,Any],validation:dict[str,Any])->int:
    capture=Path(plan['capture_root']); require(not capture.exists() and not capture.is_symlink(),f'refusing existing capture root: {capture}'); capture.mkdir(parents=True); (capture/'children').mkdir()
    jobs=plan['matrix']['jobs']; started=utc_now(); runner_start={'schema':RUNNER_SCHEMA,'status':'started','plan':validation['plan'],'runner':validation['runner'],'capture_root':str(capture),'admission_root':plan['admission_root'],'started_utc':started,'matrix':plan['matrix'],'children':[x['name'] for x in jobs],'serial':True,'no_retry':True,'terminal_before_each_child_report_or_profile_parse':True,'analysis_started':False}
    write_once(capture/'runner-start.json',runner_start); terminals=[]; admission=None; error=None
    try:
        for index,job in enumerate(jobs):
            terminal=run_child(plan,validation,job,capture); terminals.append(terminal)
            report_receipt=persist_report_oracle(plan,validation,job,terminal,capture)
            if terminal['status']!='pass': error=f'child {job["name"]} failed'; break
            if report_receipt['status']!='pass': error=f'child {job["name"]} report oracle failed'; break
            if index==0:
                profile=Path(terminal['artifacts']['profile_actual']['path'])
                admission=admission_print(plan,validation,job,capture/'children'/job['name'],profile,report_receipt)
                if admission['status']!='pass': error='first child symbol admission failed'; break
    except BaseException as e: error=f'{type(e).__name__}: {e}'
    status='pass' if error is None and len(terminals)==len(jobs) and all(x['status']=='pass' for x in terminals) and admission and admission['status']=='pass' else 'fail'
    outer={'schema':RUNNER_SCHEMA,'status':'terminal','capture_status':status,'plan':validation['plan'],'runner':validation['runner'],'capture_root':str(capture),'admission_root':plan['admission_root'],'started_utc':started,'terminal_utc':utc_now(),'children':[x['id'] for x in terminals],'expected_children':[x['name'] for x in jobs],'child_terminals_written':len(terminals),'failures':[x['id'] for x in terminals if x['status']!='pass'],'runner_error':error,'first_admission':admission,'terminal_before_each_child_report_or_profile_parse':True,'analysis_started':False,'no_retry':True}
    write_once(capture/'runner-terminal.json',outer); print(json.dumps({'status':status,'children':len(terminals),'runner_error':error,'capture_root':str(capture)},sort_keys=True)); return 0 if status=='pass' else 1

def main(argv:list[str])->int:
    parser=argparse.ArgumentParser(description=__doc__); mode=parser.add_mutually_exclusive_group(required=True); mode.add_argument('--validate-only',action='store_true'); mode.add_argument('--capture',action='store_true'); parser.add_argument('--go',action='store_true'); args=parser.parse_args(argv)
    try:
        plan,validation=validate_plan()
        if args.validate_only:
            require(not args.go,'--go only with --capture')
            public_validation={k:v for k,v in validation.items() if k!='fixture_raw'}
            print(json.dumps({'status':'validated','workload_started':False,'validation':public_validation},sort_keys=True)); return 0
        require(args.go,'capture is inert until explicit --go'); return run_capture(plan,validation)
    except (CaptureError,OSError) as e:
        print(f'0958 capture refused: {e}',file=sys.stderr); return 2
if __name__=='__main__': raise SystemExit(main(sys.argv[1:]))
