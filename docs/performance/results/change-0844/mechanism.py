"""Attribute allocation changes to insertion stacks, without phase-time claims."""
import driver as d

def verify():
 a=d.read(d.P/'analysis.json');rows=[]
 for item in a['heap_traces']:
  owner=item['attribution']['owner_attribution'];selected=[]
  for row in owner['rows']:
   if any('litchi_opc::members::PartNameIndex::insert' in (f['function'] or '') for f in row['frames_leaf_to_root']):selected.append(row)
  assert selected
  events=sum(x['allocation_events'] for x in selected);requested=sum(x['requested_bytes'] for x in selected)
  assert events<=owner['allocation_events'] and requested<=owner['requested_bytes']
  rows.append(dict(row=item['row'],insertion_allocation_events=events,insertion_requested_bytes=requested,allocation_info_indexes=[x['allocation_info_index'] for x in selected],owner_allocation_events=owner['allocation_events'],owner_requested_bytes=owner['requested_bytes']))
 comparison=[]
 for shape in ['shared','root','deep-shared','deep-unique']:
  for parts in [4,64,1024,4096]:
   for block in [0,1]:
    legs={leg:next(x for x in rows if x['row']['shape']==shape and x['row']['parts']==parts and x['row']['block']==block and x['row']['leg']==leg) for leg in ['before','after']}
    comparison.append(dict(shape=shape,parts=parts,block=block,insertion_events_removed=legs['before']['insertion_allocation_events']-legs['after']['insertion_allocation_events'],owner_events_removed=legs['before']['owner_allocation_events']-legs['after']['owner_allocation_events'],insertion_requested_bytes_removed=legs['before']['insertion_requested_bytes']-legs['after']['insertion_requested_bytes'],owner_requested_bytes_removed=legs['before']['owner_requested_bytes']-legs['after']['owner_requested_bytes']))
 return dict(status='pass',rows=rows,comparison=comparison,analysis=d.desc(d.P/'analysis.json'),reducer=d.desc(d.P/'mechanism.py'),scope='Exact owner ancestry and concrete PartNameIndex::insert stack frames; requested allocations only, not live memory, physical copies or elapsed phases.')
if __name__=='__main__':
 v=verify();p=d.P/'mechanism.json'
 if p.exists():assert d.read(p)==v
 else:d.write(p,v)
 print('Insertion allocation attribution PASS',len(v['rows']),'traces')
