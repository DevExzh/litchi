#!/usr/bin/env python3
"""Analyze per-process baseline phases without subtracting independent scopes."""
import hashlib,json,math,statistics
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def stats(v):
 v=sorted(v);return dict(n=len(v),p50=statistics.median(v),mean=statistics.mean(v),p95=v[math.ceil(.95*len(v))-1],p99=v[math.ceil(.99*len(v))-1],maximum=v[-1])
def main():
 contract=read(P/'oracle-contract.json');plan=read(P/'plan.json');f=read(P/'freeze.json');m=read(P/'captures/manifest.json');assert m['status']=='complete' and m['freeze_sha256']==sha(P/'freeze.json');assert m['bindings_start']==m['bindings_end']==f['bindings']
 for rel,h in f['bindings'].items():assert sha(ROOT/rel)==h,rel
 for b in f['binaries']:
  path=Path(b['path'])
  if path.exists():assert sha(path)==b['sha256'] and path.stat().st_size==b['bytes']
  else:
   c=read(P/'cleanup.json');assert c['removed'] and b in c['identities']
 assert len(m['runs'])==81
 seen=set();rows=[];identities={};allocs={};route_outputs={};cases={c['case']:c for c in read(P/'cases.json')}
 for r in m['runs']:
  key=(r['lane'],r['cycle'],r['case'],r['operation'],r['policy'],r['repeat']);assert key not in seen;seen.add(key);assert r['exit_code']==0
  path=P/'captures'/r['output'];assert sha(path)==r['sha256'] and sha(P/'captures'/r['stderr'])==r['stderr_sha256'];x=read(path);native=r['lane']=='native';samples=plan['samples'] if native else 1;warmups=plan['warmups'] if native else 0
  assert (x['case'],x['operation'],x['policy'],x['samples_requested'],x['warmups'])==(r['case'],r['operation'],r['policy'],samples,warmups)
  assert x['schema_version']==1 and x['input']==cases[r['case']]['path'] and x['format']==cases[r['case']]['format']
  assert x['source_inventory']['file_bytes']==cases[r['case']]['bytes']
  assert x['scope']==('public_format_open_edit_commit' if r['operation']=='format' else 'common_container_open_replace_and_validate_finish_control')
  assert x['source_sha256']==cases[r['case']]['sha256'] and x['allocator_instrumented']==(not native) and x['timing_claim']==native
  cc=contract[r['case']];assert x['directory_metadata_fields']==cc['directory_metadata_fields'] and x['allocation_ownership_contract']==cc['allocation_ownership_contract']
  assert x['expected_oracle']['semantic_witness']==cc['semantic_witness']
  assert x['changed_length_proof']['format_specific_semantic_length_proven']
  assert [o['name'] for o in x['oracle_controls']]==cc['control_names'] and all(o['rejected'] and o['status']=='rejected' and o['failure_reasons'] for o in x['oracle_controls'])
  assert x['policy_application_scope']==('not_applied_public_format_route' if r['operation']=='format' else 'common_container_editor')
  assert x['policy_applied']==(r['operation']=='container') and x['changed_length_proof']['logical_stream_length_change_proven']
  assert all(v for v in x['expected_oracle'].values() if isinstance(v,bool)) and not x['expected_oracle']['failure_reasons'];assert len(x['samples'])==samples
  ident={k:x[k] for k in ['source_sha256','expected_output_sha256','replacements_sha256','source_inventory','expected_output_inventory','replacements','changed_length_proof']};assert ident==cc['identity'] and ident==identities.setdefault(r['case'],ident)
  binary=next(b['path'] for b in f['binaries'] if Path(b['path']).name==('ole_format_save_probe' if native else 'ole_format_save_probe_alloc'))
  assert r['command']==['taskset','-c',str(plan['cpu']),binary,'--case',r['case'],'--input',cases[r['case']]['path'],'--operation',r['operation'],'--policy',r['policy'],'--samples',str(samples),'--warmups',str(warmups)]
  arrays={};phasefractions={};output_hash=None
  for n,s in enumerate(x['samples']):
   assert s['index']==n and all(v for v in s['oracle'].values() if isinstance(v,bool)) and not s['oracle']['failure_reasons'];assert s['output_inventory']['streams']==x['expected_output_inventory']['streams'];assert s['oracle']['semantic_witness']==cc['semantic_witness']
   raw=s['oracle']['raw_directory'];assert raw['source_expected_difference_bytes']==0
   if r['operation']=='format' or r['policy']=='reuse':assert raw['source_output_difference_bytes']==raw['expected_output_difference_bytes']==0
   if output_hash is None:output_hash=s['output_sha256']
   assert output_hash==s['output_sha256'];rk=(r['case'],r['operation'],r['policy']);assert s['output_sha256']==route_outputs.setdefault(rk,s['output_sha256'])
   if native:
    assert 'allocations' not in s;times=s['phase_ns'];assert all(isinstance(t,int) and t>=0 for t in times.values());assert set(times)==({'whole_ns'} if r['operation']=='format' else {'open_ns','stage_ns','finish_ns','whole_ns'})
    if r['operation']=='container':assert sum(times[k] for k in ['open_ns','stage_ns','finish_ns'])<=times['whole_ns']
    for k,v in times.items():arrays.setdefault(k,[]).append(v)
    if r['operation']=='container':
     for k in ['open_ns','stage_ns','finish_ns']:phasefractions.setdefault(k,[]).append(times[k]/times['whole_ns']*100)
   else:
    assert 'phase_ns' not in s;v=s['allocations'];ak=(r['case'],r['operation'],r['policy']);assert v==allocs.setdefault(ak,v),'allocation repeats differ'
  rows.append(dict(lane=r['lane'],cycle=r['cycle'],case=r['case'],operation=r['operation'],policy=r['policy'],repeat=r['repeat'],output_sha256=output_hash,timing={k:stats(v) for k,v in arrays.items()},phase_percent={k:stats(v) for k,v in phasefractions.items()},allocations=x['samples'][0].get('allocations')))
 expected={(lane,cy,c,op,pol,r) for lane in ['native','allocation'] for cy in range(2 if lane=='native' else 1) for c in cases for op,pol in [('format','reuse'),('container','reuse'),('container','rewrite')] for r in range(3)};assert seen==expected
 out=dict(disposition='unchanged production baseline; no speedup claim',processes=rows,case_identities=identities);(P/'analysis.json').write_text(json.dumps(out,indent=2)+'\n');print('PASS 81 processes; exact oracles/identity and per-process statistics')
if __name__=='__main__':main()
