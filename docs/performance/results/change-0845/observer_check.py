"""Bind observer reports and full-stack allocation evidence outside formal lanes."""
import driver as d
import qualify as q
import reader as r
import heap_reader as h
import sys

def verify(leg):
 folder=d.P/'observer-preflight'/leg;item=q.receipt(folder);row=item['row']
 expected=d.read(q.folder_for('qualification',dict(row,block=0))/'report.json')
 report=r.validate_report(folder/'report.json',4,'publish',1,0,shape='shared',expected=expected)
 paths=list(folder.glob('trace*.zst'));assert len(paths)==1
 value=h.parse_zst(paths[0]);assert value['owner_attribution']['allocation_events']>0
 return dict(status='pass',receipt=item,report=report,attribution=dict(owner_attribution=value['owner_attribution']),trace=d.desc(paths[0]),reader=d.desc(d.P/'reader.py'),heap_reader=d.desc(d.P/'heap_reader.py'),checker=d.desc(d.P/'observer_check.py'),scope='Separate observer compatibility process; excluded from512formalreports211328samples.')
if __name__=='__main__':
 leg=sys.argv[1];assert leg in ['before','after'];value=verify(leg);p=d.P/f'observer-preflight-{leg}.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Observer preflight PASS',leg)
