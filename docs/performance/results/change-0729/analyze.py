#!/usr/bin/env python3
"""Attribute actual DOC lifecycles and flag observer/control differences."""
import hashlib,json,math,statistics
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
OUTER=['open_ns','edit_ns','replace_ns','commit_ns','output_copy_ns']
PHASES={'open':['StrictOwnerValidation','PublicReaderValidation','SourceRetention'],'commit':['Finish','StrictOwnerValidation','PublicReaderValidation','SourceRetention','Patch']}
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def stats(v):
 v=sorted(v);return dict(n=len(v),p50=statistics.median(v),mean=statistics.mean(v),p95=v[math.ceil(.95*len(v))-1],p99=v[math.ceil(.99*len(v))-1],maximum=v[-1])
def oracle(o,witness):
 assert all(v for v in o.values() if isinstance(v,bool)) and not o['failure_reasons'] and o['semantic_witness']==witness
 raw=o['raw_directory'];assert raw['source_expected_difference_bytes']==raw['source_output_difference_bytes']==raw['expected_output_difference_bytes']==0

def sample_times(s,route):
 whole=s['whole_ns'];assert type(whole)==int and whole>0
 times={'whole_ns':whole};fractions={};parent_fractions={}
 if route=='ordinary-opaque':assert s['split'] is None
 else:
  split=s['split'];assert set(split)==set(OUTER+['split_sum_ns','whole_residual_ns']);assert all(type(v)==int and v>=0 for v in split.values())
  assert split['split_sum_ns']==sum(split[k] for k in OUTER) and split['whole_residual_ns']==whole-split['split_sum_ns']
  times.update(split)
  for k in OUTER+['whole_residual_ns']:fractions[k]=split[k]/whole*100
 if route!='profiled-clock':assert s['diagnostics'] is None and s['observer_clock_control_ns'] is None
 else:
  assert type(s['observer_clock_control_ns'])==int and s['observer_clock_control_ns']>=0;times['observer_clock_control_ns']=s['observer_clock_control_ns']
  assert set(s['diagnostics'])==set(PHASES) and s['split']['open_ns']>0 and s['split']['commit_ns']>0;last=0
  for owner,phases in PHASES.items():
   d=s['diagnostics'][owner];assert d['balanced'] and d['sequence_ok'] and d['overflow'] is False and d['expected_phases']==phases
   assert d['event_count']==len(d['events'])==2*len(phases) and len(d['spans'])==len(phases)
   total=0
   for i,phase in enumerate(phases):
    a,b=d['events'][2*i:2*i+2];span=d['spans'][i]
    assert (a['kind'],a['phase'],a['outcome'])==('started',phase,'started') and (b['kind'],b['phase'],b['outcome'])==('finished',phase,'success')
    assert type(a['t_ns'])==type(b['t_ns'])==int and last<=a['t_ns']<=b['t_ns']<=whole;last=b['t_ns']
    duration=b['t_ns']-a['t_ns'];assert span==dict(phase=phase,outcome='success',start_ns=a['t_ns'],finish_ns=b['t_ns'],duration_ns=duration)
    key=owner+'.'+phase;times[key]=duration;fractions[key]=duration/whole*100;parent_fractions[key]=duration/s['split'][owner+'_ns']*100;total+=duration
   assert total<=s['split'][owner+'_ns']
 return times,fractions,parent_fractions

def main():
 plan=read(P/'plan.json');contract=read(P/'oracle-contract.json');f=read(P/'freeze.json');m=read(P/'captures/manifest.json');cases={x['case']:x for x in read(P/'cases.json')}
 assert m['status']=='complete' and m['freeze_sha256']==sha(P/'freeze.json') and m['bindings_start']==m['bindings_end']==f['bindings']
 for rel,h in f['bindings'].items():assert sha(ROOT/rel)==h,rel
 assert len(f['binaries'])==1;b=f['binaries'][0];path=Path(b['path'])
 if path.exists():assert sha(path)==b['sha256'] and path.stat().st_size==b['bytes']
 else:
  cleanup=read(P/'cleanup.json');assert cleanup['removed'] and b in cleanup['identities']
 assert len(m['runs'])==len(plan['schedule'])==72;rows=[];outputs={};identities={}
 for r,planned in zip(m['runs'],plan['schedule']):
  assert all(r[k]==v for k,v in planned.items()) and r['exit_code']==0
  c=cases[r['case']];cc=contract[r['case']];path=P/'captures'/r['output'];assert sha(path)==r['sha256'] and sha(P/'captures'/r['stderr'])==r['stderr_sha256']
  name=f"c{r['cycle']}-{r['case']}-{r['route']}-r{r['repeat']}.json";assert r['output']==name and r['stderr']==name+'.stderr'
  assert r['command']==['taskset','-c',str(plan['cpu']),b['path'],'--case',r['case'],'--input',c['path'],'--route',r['route'],'--samples',str(plan['samples']),'--warmups',str(plan['warmups'])]
  x=read(path);assert x['schema_version']==1 and x['mode']=='doc_public_phase_attribution' and x['format']=='doc' and x['operation']=='format' and x['scope']=='public_doc_open_edit_replace_commit_output_copy'
  assert (x['case'],x['route'],x['input'],x['samples_requested'],x['warmups'])==(r['case'],r['route'],c['path'],plan['samples'],plan['warmups'])
  assert x['timing_claim'] is True and x['allocator_instrumented'] is False and x['text_utf16_units']==45
  assert x['source_sha256']==c['sha256'] and x['source_inventory']['file_bytes']==c['bytes']
  ident={k:x[k] for k in cc['identity']};assert ident==cc['identity'] and ident==identities.setdefault(r['case'],ident)
  for k,v in cc['headers'].items():assert x[k]==v
  assert x['directory_metadata_fields']==cc['directory_metadata_fields'] and x['allocation_ownership_contract']==cc['allocation_ownership_contract']
  oracle(x['expected_oracle'],cc['semantic_witness']);assert [o['name'] for o in x['oracle_controls']]==cc['control_names']
  assert all(o['status']=='rejected' and o['rejected'] is True and o['failure_reasons'] for o in x['oracle_controls'])
  assert len(x['samples'])==plan['samples'];arrays={};fractions={};parent_fractions={}
  for index,s in enumerate(x['samples']):
   assert s['index']==index and s['route']==r['route'];oracle(s['oracle'],cc['semantic_witness'])
   assert s['output_inventory']['streams']==x['expected_output_inventory']['streams'] and s['output_sha256']==x['expected_output_sha256']==outputs.setdefault(r['case'],s['output_sha256'])
   times,fs,pfs=sample_times(s,r['route'])
   for src,dest in [(times,arrays),(fs,fractions),(pfs,parent_fractions)]:
    for k,v in src.items():dest.setdefault(k,[]).append(v)
  rows.append(dict(planned,output_sha256=outputs[r['case']],timing={k:stats(v) for k,v in arrays.items()},whole_percent={k:stats(v) for k,v in fractions.items()},parent_percent={k:stats(v) for k,v in parent_fractions.items()}))
 comparisons=[];lookup={(r['cycle'],r['repeat'],r['case'],r['route']):r for r in rows}
 for cycle in range(plan['cycles']):
  for repeat in range(plan['rounds']):
   for case in plan['cases']:
    for before,after in plan['comparisons']:
     a=lookup[cycle,repeat,case,before];b=lookup[cycle,repeat,case,after];metrics={}
     for metric in ['p50','mean']:
      av=a['timing']['whole_ns'][metric];bv=b['timing']['whole_ns'][metric];delta=(bv/av-1)*100;metrics[metric]=dict(before_ns=av,after_ns=bv,delta_percent=delta,flag=abs(delta)>plan['observer_flag_percent'])
     comparisons.append(dict(cycle=cycle,repeat=repeat,case=case,before=before,after=after,metrics=metrics))
 out=dict(disposition='unchanged source public DOC attribution; no production speedup claim',processes=rows,comparisons=comparisons,case_identities=identities)
 (P/'analysis.json').write_text(json.dumps(out,indent=2)+'\n');print('PASS 72 processes; exact DOC oracles/events and per-process attribution')
if __name__=='__main__':main()
