#!/usr/bin/env python3
"""Offline positive/negative controls for the frozen 0944 adapters."""
from __future__ import annotations
import copy, hashlib, importlib.util, json, pathlib, shutil, sys
PACKET=pathlib.Path('/home/zhuhe/litchi-performance-0944')
PLAN=PACKET/'docs/performance/results/change-0944/plan.json'
OLD=pathlib.Path('/home/zhuhe/code/litchi/docs/performance/results/change-0901/capture/00-medium')

def load_mod(path,name):
 spec=importlib.util.spec_from_file_location(name,path); assert spec and spec.loader
 m=importlib.util.module_from_spec(spec);sys.modules[name]=m;spec.loader.exec_module(m);return m

def expect_reject(name,fn,results):
 try: fn()
 except Exception as e: results.append({'name':name,'status':'pass','expected':'reject','error':type(e).__name__+': '+str(e)[:180]})
 else: raise AssertionError(name+' unexpectedly accepted invalid state')

def main():
 plan=json.loads(PLAN.read_text()); source=load_mod(PACKET/'tools/source_check_v2.py','audit_source_check'); sem=load_mod(PACKET/'tools/semantic_v2.py','audit_semantic'); raw=load_mod(PACKET/'tools/raw_analyzer_v2.py','audit_raw')
 results=[]
 # Custody/schema controls.
 expect_reject('duplicate-json-key',lambda: source.reject_pairs([('x',1),('x',2)]),results)
 for name,value in [('absolute-path','/tmp/x'),('parent-path','../x'),('nul-path','a\x00b')]: expect_reject(name,lambda value=value: source.safe_rel(value,name),results)
 source.check(PLAN);results.append({'name':'current-source-plan-positive','status':'pass','expected':'accept'})
 # Retained 0901 reports are positive semantic controls with their own frozen binary identity.
 medium=json.loads((OLD/'report.json').read_text()); dense=json.loads((OLD.parent/'01-dense-sparse/report.json').read_text())
 for shape,report in [('medium',medium),('dense-sparse',dense)]:
  synthetic=copy.deepcopy(plan);synthetic['binary']=dict(plan['binary']); synthetic['binary'].update({'bytes':report['binary_identity']['binary_bytes'],'sha256':report['binary_identity']['binary_sha256']})
  sem.validate(synthetic,report,shape);results.append({'name':'retained-semantic-positive-'+shape,'status':'pass','expected':'accept'})
 bad=copy.deepcopy(medium);bad['binary_identity']['binary_sha256']='0'*64
 synthetic=copy.deepcopy(plan);synthetic['binary']=dict(plan['binary']);synthetic['binary'].update({'bytes':medium['binary_identity']['binary_bytes'],'sha256':medium['binary_identity']['binary_sha256']})
 expect_reject('semantic-binary-digest',lambda:sem.validate(synthetic,bad,'medium'),results)
 bad=copy.deepcopy(medium);bad['results'][0]['source']['xlsx_cell_values']['reopen_ns']=[True]
 expect_reject('semantic-bool-reopen',lambda:sem.validate(synthetic,bad,'medium'),results)
 bad=copy.deepcopy(medium);bad['results'][0]['output_sha256']='0'*64
 expect_reject('semantic-output-identity',lambda:sem.validate(synthetic,bad,'medium'),results)
 bad=copy.deepcopy(medium);del bad['results'][0]['operation_metrics']
 expect_reject('semantic-missing-operation-metrics',lambda:sem.validate(synthetic,bad,'medium'),results)
 # Offline raw positive and two previously weak controls: suffix and terminal part gap.
 base=PACKET/'controls/retained-medium';
 if base.exists(): shutil.rmtree(base)
 shutil.copytree(OLD,base)
 oldbin='/home/zhuhe/code/litchi/docs/performance/results/change-0894/binaries/baseline-native'; newbin=str(PACKET/'controls/retained-binary'); oldreport=str(OLD/'report.json'); newreport=str(base/'report.json')
 for path in base.glob('callgrind.out*'):
  text=path.read_text(errors='strict'); text=text.replace(oldbin,newbin).replace(oldreport,newreport); path.write_text(text)
 start_path=base/'start.json'; start_path.write_text(start_path.read_text().replace(oldbin,newbin).replace(oldreport,newreport))
 rawplan=copy.deepcopy(plan);rawplan['binary']={'path':'controls/retained-binary'}
 rawplan_path=PACKET/'controls/raw-positive-plan.json';rawplan_path.write_text(json.dumps(rawplan,indent=2,sort_keys=True)+'\n')
 row={'index':0,'repeat':0,'shape':'medium'}
 raw.analyze(rawplan_path,base,row)
 results.append({'name':'retained-raw-positive','status':'pass','expected':'accept'})
 extra=PACKET/'controls/retained-extra-suffix';shutil.rmtree(extra,ignore_errors=True);shutil.copytree(base,extra);shutil.copy2(extra/'callgrind.out.1',extra/'callgrind.out.1.extra')
 expect_reject('raw-extra-suffix',lambda:raw.analyze(rawplan_path,extra,row),results)
 gap=PACKET/'controls/retained-terminal-gap';shutil.rmtree(gap,ignore_errors=True);shutil.copytree(base,gap); gp=gap/'callgrind.out';gt=gp.read_text();gp.write_text(gt.replace('part: 7','part: 8',1))
 expect_reject('raw-terminal-gap',lambda:raw.analyze(rawplan_path,gap,row),results)
 out=PACKET/'controls/audit-controls-v2.json';out.parent.mkdir(parents=True,exist_ok=True);doc={'schema':'litchi-xlsx-profile-audit-controls-v2','status':'pass','control_count':len(results),'controls':results};out.write_text(json.dumps(doc,indent=2,sort_keys=True)+'\n');print(json.dumps({'status':'pass','control_count':len(results),'path':str(out.relative_to(PACKET))},sort_keys=True))
if __name__=='__main__': main()
