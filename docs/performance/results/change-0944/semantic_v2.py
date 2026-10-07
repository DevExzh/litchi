#!/usr/bin/env python3
"""Current packet semantic/reopen/preservation gate using retained 0705 rules."""
from __future__ import annotations
import hashlib, importlib.util, json, pathlib, sys
from typing import Any
PACKET=pathlib.Path(__file__).resolve().parent.parent
SEM=PACKET/'docs/performance/results/change-0944/helpers/retained-0705-semantic-analyze.py'
HEX=set('0123456789abcdef')
class EvidenceError(ValueError): pass
def reject_pairs(pairs):
 d={}
 for k,v in pairs:
  if k in d: raise EvidenceError(f'duplicate JSON key: {k}')
  d[k]=v
 return d
def load(path):
 try:return json.loads(path.read_text(encoding='utf-8'),object_pairs_hook=reject_pairs)
 except Exception as e: raise EvidenceError(f'cannot parse {path}: {e}') from e
def relpath(value,label):
 if not isinstance(value,str) or not value or value.startswith('/') or '\\x00' in value: raise EvidenceError(f'{label} is not relative')
 p=pathlib.PurePosixPath(value)
 if '..' in p.parts or '.' in p.parts: raise EvidenceError(f'{label} escapes packet')
 return PACKET/pathlib.Path(*p.parts)
def sha(path):
 if not path.is_file() or path.is_symlink(): raise EvidenceError(f'missing regular file {path}')
 return hashlib.sha256(path.read_bytes()).hexdigest()
def load_retained():
 spec=importlib.util.spec_from_file_location('retained_semantic_0705',SEM)
 if spec is None or spec.loader is None: raise EvidenceError('cannot load retained semantic validator')
 m=importlib.util.module_from_spec(spec);sys.modules[spec.name]=m;spec.loader.exec_module(m);return m
RETAINED=load_retained()
def require(cond,msg):
 if not cond: raise EvidenceError(msg)
def validate(plan:dict[str,Any], report:dict[str,Any], shape:str)->dict[str,Any]:
 require(report.get('schema_version')==1,'report schema differs')
 b=report.get('binary_identity'); require(isinstance(b,dict),'binary identity missing')
 require(b.get('binary_sha256')==plan['binary']['sha256'],'binary SHA does not bind plan')
 require(b.get('binary_bytes')==plan['binary']['bytes'],'binary size does not bind plan')
 require(b.get('profile')=='release' and b.get('executable') is True,'binary profile/executable gate failed')
 env=report.get('environment'); require(isinstance(env,dict),'report environment missing')
 require(env.get('cpu_affinity')=='12','CPU affinity gate failed')
 require(env.get('git_revision') is None and env.get('git_worktree_dirty') is None and env.get('rustc_version') is None,'detached source runtime identity gate failed')
 tool=report.get('tool'); require(isinstance(tool,dict),'report tool missing')
 require(tool.get('profile')=='release' and tool.get('instrumentation')=='none','runtime instrumentation gate failed')
 cfg=report.get('configuration'); require(isinstance(cfg,dict),'report configuration missing')
 require(cfg.get('cases')==[plan['case']] and cfg.get('xlsx_cell_crud_shapes')==[shape],'case/shape gate failed')
 require(cfg.get('samples_per_case')==1 and cfg.get('warmup_iterations_per_case')==0,'sample policy gate failed')
 rows=report.get('results'); require(isinstance(rows,list) and len(rows)==1,'result cardinality gate failed')
 row=rows[0]; require(isinstance(row,dict) and row.get('case')==plan['case'],'result identity gate failed')
 retained=RETAINED
 corpus=retained.validate_corpus(row.get('corpus'),shape)
 retained.validate_sink(row.get('sink'),shape)
 phases,constants=retained.validate_source(row,corpus,shape,1,compatibility=False)
 native_path=relpath(plan['native_context']['path'],'native_context.path')
 native=load(native_path)
 reference=next((v for v in native.get('rows',[]) if v.get('shape')==shape and v.get('lane')=='native'),None)
 require(reference is not None,'native semantic context row missing')
 require(corpus==reference['corpus'] and constants==reference['constants'] and row['sink']==reference['sink'],'corpus/source/sink semantic identity differs from retained context')
 summary=row['source']['xlsx_cell_values']
 for key in ('open_ns','plan_ns','commit_ns','publication_ns','reopen_ns'):
  val=summary.get(key); require(isinstance(val,list) and len(val)==1 and type(val[0]) is int and val[0]>=0,f'{key} reopen/timing vector gate failed')
 require('reopen' in summary.get('timing_scope','') and 'excluded' in summary.get('timing_scope',''),'reopen timing scope is not explicit')
 for key in ('output_sha256','semantic_sha256','untouched_member_sha256'):
  val=summary.get(key); require(isinstance(val,list) and len(val)==1 and isinstance(val[0],str) and len(val[0])==64 and set(val[0])<=HEX,f'{key} preservation digest gate failed')
 require(type(summary.get('untouched_member_count')) is int and summary['untouched_member_count']>0,'untouched member count gate failed')
 require(row.get('output_sha256')==summary['output_sha256'][0],'output digest binding failed')
 metrics=row.get('operation_metrics'); require(isinstance(metrics,dict),'operation metrics missing')
 require(metrics.get('sample_count')==1 and metrics.get('sample_indices')==[0],'operation sample alignment gate failed')
 require(metrics.get('latency_claim')=='comparable_timed_operation','operation metric identity changed')
 # The harness's actual reopen, package identity, semantic hash, and untouched
 # member checks are inside the binary; these retained fields are the durable
 # evidence that those gates ran and that reopen is outside the timed scope.
 return {'status':'pass','shape':shape,'corpus':corpus,'constants':constants,'sink':row['sink'],'timing_policy':'Instrumented elapsed durations retained but excluded from Callgrind/native conclusions. Reopen and preservation gates run inside the harness outside the timed phases.','native_reference':plan['native_context'],'reopen_ns':summary['reopen_ns'][0],'preservation':{'output_sha256':summary['output_sha256'][0],'semantic_sha256':summary['semantic_sha256'][0],'untouched_member_count':summary['untouched_member_count'],'untouched_member_sha256':summary['untouched_member_sha256'][0]}}
def main():
 if len(sys.argv)!=4: raise SystemExit('usage: semantic_v2.py PLAN REPORT SHAPE')
 plan=load(pathlib.Path(sys.argv[1])); report=load(pathlib.Path(sys.argv[2])); shape=sys.argv[3]
 try: print(json.dumps(validate(plan,report,shape),indent=2,sort_keys=True))
 except (EvidenceError,AssertionError,StopIteration,KeyError) as e: print(f'semantic_v2: {e}',file=sys.stderr); return 1
 return 0
if __name__=='__main__': raise SystemExit(main())
