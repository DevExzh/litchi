"""Independent valid-tag lexical scan, histogram census and semantic identity audit."""
import collections,json,sys
from pathlib import Path
import custody as c
p=c.P

def lexical(row):
 data=bytes(row['tag_bytes']);name=bytes(row['element_name_bytes']);assert data.startswith(name) and name
 i=len(name);keys=[];n=len(data);space=b' \t\r\n'
 while True:
  while i<n and data[i] in space:i+=1
  if i==n:break
  start=i
  while i<n and data[i] not in space+b'=':i+=1
  key=data[start:i];assert key and key not in keys
  while i<n and data[i] in space:i+=1
  assert i<n and data[i]==61;i+=1
  while i<n and data[i] in space:i+=1
  assert i<n and data[i] in (34,39);quote=data[i];i+=1
  end=data.find(bytes([quote]),i);assert end>=i;i=end+1;keys.append(key)
 return len(keys)

def identity(r):
 s=r['samples'][0]
 result={k:r[k] for k in ['shape','mode','slides','shapes_per_slide','source','fixture','marker']}|{'sample':{k:s[k] for k in ['source_sha256','output','verification']}}
 result['sample']['metrics']={k:v for k,v in s['metrics'].items() if k!='elapsed_ns'}
 return result

historical={}
for row in c.read(p.parent/'change-0794/qualification/receipts.json'):
 r=c.read(row['report']['path']);historical[(r['shape'],r['mode'])]=identity(r)
controls={}
for receipt in c.read(p/'control/receipts.json'):
 r=c.read(receipt['report']['path']);assert len(r['samples'])==1 and 'census' not in r['samples'][0]
 key=(r['shape'],r['mode']);controls[key]=identity(r);assert controls[key]==historical[key]
assert len(controls)==15
outputs=[];repeats={}
for receipt in c.read(p/'census/receipts.json'):
 r=c.read(receipt['report']['path']);assert len(r['samples'])==1;key=(r['shape'],r['mode']);assert identity(r)==controls[key]
 d=r['samples'][0]['census'];assert not d['counter_saturated'] and d['live_instances_at_finish']==0
 assert d['instance_identity_qualified'] and d['aggregation_conserved']
 counts=collections.Counter();tags=collections.Counter();drops=clones=starts=total=next_calls=successful=errors=ends=early=partial=never=0
 for row in d['rows']:
  f=row['frequency'];assert isinstance(f,int) and f>0;total+=f
  assert not row['counter_saturated'] and row['lexical_scan_completed'] and row['dropped'] and not row['live_at_finish']
  available=lexical(row);assert available==row['lexical_attribute_count']==row['lexical_item_count'] and row['lexical_error_count']==0
  assert row['next_calls']==row['successful_yields']+row['error_yields']+row['end_yields']
  consumed=row['starting_successful_yields']+row['successful_yields'];assert consumed<=available
  assert row['partial_consumption']==(consumed<available)
  expected='error' if row['error_yields'] else ('exhausted' if row['end_yields'] else 'dropped')
  assert row['termination']==expected and row['early_drop']==(expected=='dropped')
  drops+=f;clones+=f*row['is_clone'];starts+=f*(not row['is_clone']);next_calls+=f*row['next_calls'];successful+=f*row['successful_yields'];errors+=f*row['error_yields'];ends+=f*row['end_yields'];early+=f*row['early_drop'];partial+=f*row['partial_consumption'];never+=f*(row['next_calls']==0)
  counts[str(available)]+=f;tags[bytes(row['element_name_bytes']).hex()]+=f
 assert total==d['raw_instance_rows']==d['iterator_starts']+d['iterator_clones']==d['iterator_drops']
 assert (starts,clones,drops)==(d['iterator_starts'],d['iterator_clones'],d['iterator_drops'])
 if key in repeats:assert d==repeats[key]
 else:repeats[key]=d
 outputs.append({'shape':key[0],'mode':key[1],'block':receipt['block'],'instances':total,'starts':starts,'clones':clones,'drops':drops,'next_calls':next_calls,'successful_yields':successful,'error_yields':errors,'end_yields':ends,'early_drops':early,'partial_consumption':partial,'never_advanced':never,'attribute_count_histogram':dict(sorted(counts.items(),key=lambda x:int(x[0]))),'element_name_hex_histogram':dict(sorted(tags.items()))})
assert len(outputs)==30 and len(repeats)==15
result={'schema':'litchi.performance.0798.independent-audit.v1','reports':45,'samples':45,'rows':outputs,'native_timing_claim':False,'lexical_scope':'Generated well-formed tag bodies only; malformed row would fail this independent audit'}
if '--check' in sys.argv:assert c.read(p/'root-audit.json')==result
else:assert not (p/'root-audit.json').exists();c.write(p/'root-audit.json',result)
print('Independent valid-tag lexical, count, repeat and semantic audit PASS')
