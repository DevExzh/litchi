#!/usr/bin/env python3
from __future__ import annotations
import hashlib,json,os,pathlib,shutil,stat,sys
from datetime import datetime,timezone
PACKET=pathlib.Path('/home/zhuhe/litchi-performance-0944')
PLAN_DIR=PACKET/'docs/performance/results/change-0944'

def sha(path):
 h=hashlib.sha256();n=0
 with path.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''):h.update(b);n+=len(b)
 return n,h.hexdigest()
def rel(path):return str(path.resolve().relative_to(PACKET.resolve()))
def desc(path):
 n,h=sha(path); return {'path':rel(path),'bytes':n,'sha256':h}
def sysdesc(path):
 p=pathlib.Path(path);n,h=sha(p);return {'path':str(p),'bytes':n,'sha256':h}
def write(path,obj):
 with path.open('x',encoding='utf-8') as f:json.dump(obj,f,indent=2,sort_keys=True);f.write('\n');f.flush();os.fsync(f.fileno())

def main():
 build_start=PACKET/'build/native-build-start.json'; build_term=PACKET/'build/native-build-terminal.json'; built=PACKET/'target-native/release/litchi-perf-baseline'
 for q in (build_start,build_term,built):
  if not q.is_file(): raise SystemExit(f'missing completed build input: {q}')
 (PLAN_DIR/'binaries').mkdir(parents=True,exist_ok=True)
 binary=PLAN_DIR/'binaries/profile-native'
 if binary.exists(): raise SystemExit(f'binary destination exists: {binary}')
 shutil.copy2(built,binary); binary.chmod(stat.S_IMODE(built.stat().st_mode))
 n,h=sha(binary)
 bt=json.loads(build_term.read_text())
 if bt.get('returncode')!=0 or bt.get('timed_out') is not False or bt.get('binary',{}).get('sha256')!=h: raise SystemExit('completed build terminal does not bind binary')
 ref=PACKET/'references/native-context-0900.json'; context=json.loads(ref.read_text())
 expected={}
 for shape in ('medium','dense-sparse'):
  rows=[r for r in context['rows'] if r.get('shape')==shape and r.get('lane')=='native']
  if not rows: raise SystemExit(f'missing retained expected row {shape}')
  row=rows[0]
  expected[shape]={'corpus':row['corpus'],'sink':row['sink'],'constants':row['constants'],'phases':row['phases'],'total':row['total']}

 default_tools=['/usr/bin/taskset','/usr/bin/valgrind','/usr/bin/callgrind_annotate','/usr/bin/x86_64-linux-gnu-nm','/usr/bin/x86_64-linux-gnu-objdump','/usr/libexec/valgrind/callgrind-amd64-linux','/usr/bin/valgrind.bin']
 tools={pathlib.Path(x).name:sysdesc(x) for x in default_tools}
 # Keep names stable for the backend paths with duplicate basenames resolved by path below.
 tools['callgrind_backend']=tools.pop('callgrind-amd64-linux'); tools['valgrind_bin']=tools.pop('valgrind.bin')
 plan={
  'schema':'litchi.xlsx-publication-profile-0944.v2','created_utc':datetime.now(timezone.utc).isoformat(),
  'head':'933ee45cbb5d2ef9dc29583f7ec890ad234e04fb','tree':'8b571b08a83c530c50fdbcb16345a9aa75883794',
  'source_root':'source-build','source_manifest':desc(PACKET/'source-build-manifest.json'),'full_source_manifest':desc(PACKET/'source-manifest.json'),
  'source_closure':{'included_prefixes':['.cargo/','crates/','test-data/','tools/perf-baseline/'],'included_root_files':['Cargo.toml','README.md','clippy.toml','deny.toml','rust-toolchain.toml','rustfmt.toml'],'tracked_3rdparty_entries':0,'file_count':10788},
  'accepted_0905':{'commit':'8267929233736e3f4f4612eee94e47016a74685e','path':'source-build/crates/xml-minifier/src/audit.rs','sha256':'eddc437156c976c55ce36b0bf8f0ef47cc4943d296735a1ac8d6ef120b8754e7'},
  'binary':{'path':rel(binary),'bytes':n,'sha256':h,'mode_bits':stat.S_IMODE(binary.stat().st_mode),'executable':os.access(binary,os.X_OK)},
  'build_start':desc(build_start),'build':desc(build_term),'build_process_group_audit':desc(PACKET/'build/build-process-group-audit.json'),
  'case':'xlsx_source_backed_cell_values_one_percent_edit_save','shapes':['medium','dense-sparse'],'cpu':12,'repeats':2,'samples':1,'warmup':0,
  'environment':{'PATH':'/usr/bin:/bin','LANG':'C','LC_ALL':'C','TZ':'UTC','RUST_BACKTRACE':'0'},
  'owner':'litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream','measured_parent':'litchi_perf_baseline::run_xlsx_cell_values_edit_save','lifecycle_parent':'litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates','topology_owner':'litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream',
  'options':['--tool=callgrind','--collect-atstart=no','--toggle-collect=litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream','--zero-before=litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream','--dump-after=litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream'],
  'schedule':[{'index':0,'repeat':0,'shape':'medium'},{'index':1,'repeat':0,'shape':'dense-sparse'},{'index':2,'repeat':1,'shape':'dense-sparse'},{'index':3,'repeat':1,'shape':'medium'}],
  'expected_measured_calls':1,'historical_expected_lifecycle_calls':5,'first_child_admission':'Capture index 0 then require raw Callgrind owner/lifecycle and semantic/reopen/preservation admission before index 1; no retries or replacement rows.',
  'scope':'Callgrind Ir attribution for the publish_multi_commit_to_stream owner only. Returned snapshot destruction is outside the owner scope and remains a native-timer-only context. No native latency, cycles, allocation, RSS, or speedup conclusion.',
  'source_scope':{'generator':'litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1','expected_shapes':['medium','dense-sparse'],'reopen_and_preservation':'Harness reopens output and validates updated values, semantic output identity, OPC part/relationship identity, untouched member digests, sink bound, and source/cache counters outside timed phases.'},
  'expected_retained_context':desc(ref),'expected':expected,
  'native_context':desc(ref),'prior_semantic_validator':desc(PACKET/'docs/performance/results/change-0944/helpers/retained-0705-semantic-analyze.py'),
  'source_checker':desc(PACKET/'tools/source_check_v2.py'),'semantic_validator':desc(PACKET/'tools/semantic_v2.py'),'reader':desc(PACKET/'docs/performance/results/change-0944/change-0845/reader_v2.py'),'raw_analyzer':desc(PACKET/'tools/raw_analyzer_v2.py'),'capture_runner':desc(PACKET/'tools/capture_v2.py'),'retained_helpers':desc(PACKET/'docs/performance/results/change-0944/retained-helper-custody.json'),'tools':tools,
 }
 write(PLAN_DIR/'plan.json',plan)
 print(json.dumps({'plan':rel(PLAN_DIR/'plan.json'),'binary':plan['binary'],'plan_sha256':hashlib.sha256((PLAN_DIR/'plan.json').read_bytes()).hexdigest()},sort_keys=True))
if __name__=='__main__': main()
