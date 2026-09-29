"""Identify concrete current-source ancestor staging allocation stacks."""
import driver as d

def verify():
 a=d.read(d.P/'analysis.json');rows=[]
 for x in a['heaps']:
  if x['row']['parts']!=4096:continue
  selected=[]
  for line in [201,209]:
   stacks=[r for r in x['attribution']['owner_attribution']['rows'] if any(f['line']==line and (f['file'] or '').endswith('/members.rs') and 'PartNameIndex::insert' in (f['function'] or '') for f in r['frames_leaf_to_root'])]
   assert len(stacks)==1
   r=stacks[0];assert r['allocation_events']==4096
   selected.append(dict(line=line,allocation_events=r['allocation_events'],requested_bytes=r['requested_bytes'],allocation_info_index=r['allocation_info_index']))
  rows.append(dict(row=x['row'],stacks=selected))
 assert len(rows)==2
 return dict(status='pass',rows=rows,analysis=d.desc(d.P/'analysis.json'),source=d.desc(d.ROOT/'crates/litchi-opc/src/members.rs'),reducer=d.desc(d.P/'ancestor_allocation.py'),scope='Allocation attribution only; no native speedup prediction or production candidate.')
if __name__=='__main__':
 v=verify();p=d.P/'ancestor-allocation.json'
 if p.exists():assert d.read(p)==v
 else:d.write(p,v)
 print('Ancestor allocation attribution PASS')
