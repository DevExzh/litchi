"""Qualify observer support without counting preflight as formal evidence."""
import driver as d
import qualify as q
import reader as r
import heap_reader as h

def verify():
 rows=[]
 for lane in ['profiles','heap']:
  folder=d.P/'observer-preflight'/lane;item=q.receipt(folder);row=item['row']
  expected=d.read(q.folder_for('qualification',row)/'report.json')
  validated=r.validate_report(folder/'report.json',4,'publish',1,0,expected=expected)
  detail={}
  if lane=='profiles':
   assert {p.name for p in folder.glob('profile.callgrind*')}=={'profile.callgrind','profile.callgrind.1'}
   command=d.read(q.bind(d.read(folder/'receipt.json')['command']))['argv']
   detail=r.validate_profile(folder/'profile.callgrind.1',folder/'profile.callgrind',command=command)
  else:
   paths=list(folder.glob('trace*.zst'));assert len(paths)==1
   value=h.parse_zst(paths[0]);assert value==d.read(d.P/'heap-preflight-validation.json')['attribution'] and value['owner_attribution']['allocation_events']>0
   detail=dict(trace=d.desc(paths[0]),owner_allocation_events=value['owner_attribution']['allocation_events'])
  rows.append(dict(**item,**validated,observer=detail))
 return dict(status='pass',rows=rows,reader=d.desc(d.P/'reader.py'),heap_reader=d.desc(d.P/'heap_reader.py'),checker=d.desc(d.P/'observer_check.py'),scope='Observer compatibility only; excluded from formal72reports1464samples.')
if __name__=='__main__':
 value=verify();p=d.P/'observer-preflight-validation.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Both observer preflights independently validated')
