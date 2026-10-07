#!/usr/bin/env python3
"""Raw Callgrind admission/reduction adapter for the frozen 0944 matrix."""
from __future__ import annotations
import hashlib, importlib.util, json, pathlib, sys
from typing import Any
PACKET=pathlib.Path(__file__).resolve().parent.parent
ADAPTER=PACKET/'docs/performance/results/change-0944/analyze_v2_adapter.py'
OWNER='litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream'
MEASURED='litchi_perf_baseline::run_xlsx_cell_values_edit_save'
LIFECYCLE='litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates'

def reject_pairs(pairs):
 d={}
 for k,v in pairs:
  if k in d: raise ValueError(f'duplicate key: {k}')
  d[k]=v
 return d

def load(path): return json.loads(path.read_text(encoding='utf-8'),object_pairs_hook=reject_pairs)
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def packet_rel(path): return str(path.resolve().relative_to(PACKET.resolve()))
def load_adapter():
 spec=importlib.util.spec_from_file_location('retained_callgrind_adapter_0901',ADAPTER)
 if spec is None or spec.loader is None: raise ValueError('cannot load retained Callgrind adapter')
 mod=importlib.util.module_from_spec(spec);sys.modules[spec.name]=mod;spec.loader.exec_module(mod);mod.HERE=PACKET;return mod
MOD=load_adapter()
def require(c,m):
 if not c: raise ValueError(m)
def analyze(plan_path: pathlib.Path, folder: pathlib.Path, row:dict[str,Any])->dict[str,Any]:
 plan=load(plan_path); require(plan.get('owner')==OWNER,'owner differs'); require(plan.get('measured_parent')==MEASURED,'measured parent differs'); require(plan.get('lifecycle_parent')==LIFECYCLE,'lifecycle parent differs')
 folder=folder.resolve(); require(folder.is_dir(),'capture folder missing')
 all_dumps=[p for p in folder.iterdir() if p.name.startswith('callgrind.out.') and p.name != 'callgrind.out']
 require(all(not p.is_symlink() and p.is_file() for p in all_dumps),'Callgrind artifact set contains a non-regular entry')
 malformed=[p.name for p in all_dumps if not p.name[len('callgrind.out.'):].isdigit()]
 require(not malformed,f'malformed numbered Callgrind suffixes: {malformed}')
 numbered=sorted(all_dumps, key=lambda p:int(p.suffix[1:]))
 require(numbered,'no numbered Callgrind dumps')
 nums=[int(p.suffix[1:]) for p in numbered]; require(nums==list(range(1,len(nums)+1)),f'non-contiguous dump sequence {nums}')
 start_doc=load(folder/'start.json'); frozen_argv=start_doc.get('argv'); require(isinstance(frozen_argv,list) and all(isinstance(x,str) for x in frozen_argv),'start argv is missing')
 require(start_doc.get('row')==row,'start row differs')
 try: warmup_index=frozen_argv.index('--warmup')
 except ValueError: raise ValueError('start argv has no warmup boundary')
 require(warmup_index>0,'start argv has no binary before warmup')
 expected_command=frozen_argv[warmup_index-1:]
 require(pathlib.PurePosixPath(expected_command[0]).name==pathlib.PurePosixPath(plan['binary']['path']).name,'start argv binary basename differs from frozen plan')
 parts=[]; measured=[]; lifecycle=0
 for path in numbered:
  parsed=MOD.READER.parse_raw(path); h=parsed['header']; part=int(path.suffix[1:])
  require(h.get('part')==part,f'{path} part mismatch'); require(h.get('events')==['Ir'],f'{path} events differ'); require(h.get('trigger')=='--dump-after='+OWNER,f'{path} trigger mismatch'); require(h.get('command','').split()==expected_command,f'{path} header command differs from frozen child argv'); require(h.get('summary_ir')==h.get('totals_ir') and h.get('summary_ir',0)>0,f'{path} summary/totals invalid'); require(sum(f.get('self_ir',0) for f in parsed.get('functions',{}).values())==h['summary_ir'],f'{path} exclusive self rows do not conserve')
  owner=MOD.owner_rows(parsed,OWNER,MEASURED,LIFECYCLE); owner.update({'path':packet_rel(path),'sha256':sha(path),'part':part,'summary_ir':h['summary_ir']})
  if owner['role']=='measured': measured.append((path,parsed,owner))
  else: lifecycle += 1
  parts.append({'file':packet_rel(path),'part':part,'sha256':sha(path),'header':h,'owner':owner})
 require(len(measured)==1,f'measured owner count is {len(measured)}')
 terminal=folder/'callgrind.out'; require(terminal.is_file(),'terminal dump missing')
 term=MOD.READER.parse_raw(terminal,allow_empty=True); th=term['header']; require(th.get('part')==nums[-1]+1,'terminal part has a gap'); require(th.get('events')==['Ir'],'terminal events differ'); require(th.get('command','').split()==expected_command,'terminal header command differs from frozen start argv'); require(th.get('trigger')=='Program termination','terminal trigger mismatch'); require(th.get('summary_ir')==0 and th.get('totals_ir')==0 and not term.get('functions'),'terminal is not empty')
 selected_path,selected_parsed,selected_owner=measured[0]
 topology=MOD.topology_partition(selected_parsed,selected_owner['selected']['callee_id'])
 nested=MOD.nested_diagnostics(selected_parsed)
 self_rows=MOD.top_self_rows(selected_parsed)
 collected=MOD.stderr_collected(folder,parts)
 return {'schema':'litchi.xlsx-publication-profile-raw-analysis-0944.v2','status':'pass','plan':{'path':packet_rel(plan_path),'sha256':sha(plan_path)},'index':row['index'],'repeat':row['repeat'],'shape':row['shape'],'parts':parts,'terminal':{'file':packet_rel(terminal),'part':th['part'],'sha256':sha(terminal),'summary_ir':0,'totals_ir':0,'trigger':th['trigger']},'classification':{'measured_count':1,'lifecycle_count':lifecycle,'measured_part':selected_owner['part'],'measured_file':selected_owner['path'],'lifecycle_prediction':plan.get('historical_expected_lifecycle_calls')},'measured':selected_owner,'measured_exclusive_self':self_rows,'stderr_collected':collected,'topology':topology,'nested_diagnostics':nested,'command_binding':{'frozen_start_argv':frozen_argv,'header_command':expected_command,'terminal_header_command':th.get('command','').split(),'origin_root':plan.get('origin_root')},'validation':{'all_numbered_dumps_retained':True,'all_numbered_dumps_have_known_owner_parent':True,'one_measured_owner_call':True,'lifecycle_calls_retained_separately':True,'numbered_summary_equals_totals_for_all_parts':True,'actual_lifecycle_count_recorded':True,'terminal_before_parse':True,'nested_rows_not_summed':True}}
def main():
 if len(sys.argv)!=4: raise SystemExit('usage: raw_analyzer_v2.py PLAN FOLDER ROW_JSON')
 plan=pathlib.Path(sys.argv[1]); folder=pathlib.Path(sys.argv[2]); row=load(pathlib.Path(sys.argv[3]))
 try:
  result=analyze(plan,folder,row); print(json.dumps(result,indent=2,sort_keys=True))
 except Exception as e: print(f'raw_analyzer_v2: {e}',file=sys.stderr); return 1
 return 0
if __name__=='__main__': raise SystemExit(main())
