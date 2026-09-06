#!/usr/bin/env python3
"""Replay the six aligned pilots and reject altered fixture/resource evidence."""
import argparse,copy,hashlib,importlib.util,json,shutil,tempfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent

def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def run():
 spec=importlib.util.spec_from_file_location('odp439oracle',ROOT/'oracle/verify-report.py');oracle=importlib.util.module_from_spec(spec);spec.loader.exec_module(oracle)
 pilots=ROOT/'pilots/after/aligned';controls=[]
 for mode in ['normal','allocator']:
  for shape in ['tiny','medium','large']:
   path=pilots/f'{mode}-{shape}.json';oracle.validate_report(path,mode,shape,samples=3,warmups=1);controls.append({'path':str(path.relative_to(ROOT)),'sha256':sha(path)})
 path=pilots/'allocator-tiny.json';base=json.loads(path.read_text());source_hash=sha(path)
 def summary(r):return r['results'][0]['source']['odp_append']
 def alloc(r):return r['results'][0]['operation_metrics']['allocation']
 def change(r,key):summary(r)[key]='0'*64
 mutations={
  'source_semantic':lambda r:change(r,'source_semantic_sha256'),
  'output_order':lambda r:change(r,'output_order_sha256'),
  'opaque_formula':lambda r:change(r,'opaque_sha256'),
  'text_projection':lambda r:change(r,'source_text_projection_sha256'),
  'manifest_gate':lambda r:summary(r).__setitem__('output_manifest_bindings_verified',False),
  'untouched_gate':lambda r:summary(r).__setitem__('untouched_members_verified',False),
  'timing_open_excluded':lambda r:summary(r).__setitem__('timing_scope','Snapshot::from_bytes outside; transaction add commit write inside; oracle drop outside'),
  'chronological_order':lambda r:summary(r)['lifecycle_ns'].reverse(),
  'allocation_balance':lambda r:alloc(r)['live_bytes_after']['values'].__setitem__(0,alloc(r)['live_bytes_after']['values'][0]+1),
  'allocation_peak':lambda r:alloc(r)['region_peak_live_bytes']['values'].__setitem__(0,0),
  'allocation_failure':lambda r:alloc(r)['failed_allocation_calls']['values'].__setitem__(0,1),
  'allocator_revision':lambda r:r['tool'].__setitem__('allocator_counter_revision','old'),
  'binary_profile':lambda r:r['binary_identity'].__setitem__('profile','debug'),
  'sink_count':lambda r:r['results'][0]['operation_metrics']['sink']['accepted_bytes']['values'].__setitem__(0,1),
 }
 records=[]
 with tempfile.TemporaryDirectory(prefix='litchi-0439-oracle-probes-') as tmp:
  report=Path(tmp)/'report.json';shutil.copyfile(pilots/'allocator-tiny-catalog.json',Path(tmp)/'report-catalog.json')
  for name,mutate in mutations.items():
   value=copy.deepcopy(base);mutate(value);assert value!=base,name;report.write_text(json.dumps(value))
   try:oracle.validate_report(report,'allocator','tiny',samples=3,warmups=1)
   except oracle.ValidationError as error:records.append({'name':name,'status':'rejected','reason':str(error)})
   else:raise AssertionError('oracle accepted mutation '+name)
 assert sha(path)==source_hash
 return {'change':439,'status':'pass','oracle_sha256':sha(ROOT/'oracle/verify-report.py'),'driver_sha256':sha(Path(__file__)),'controls':controls,'probes':records,'source_unchanged':True}
if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--check',action='store_true');args=parser.parse_args();value=run();target=ROOT/'oracle-probes.json'
 if args.check:assert json.loads(target.read_text())==value
 else:
  with target.open('x') as stream:stream.write(json.dumps(value,indent=2)+'\n')
 print(json.dumps({'status':'pass','controls':len(value['controls']),'rejected':len(value['probes'])}))
