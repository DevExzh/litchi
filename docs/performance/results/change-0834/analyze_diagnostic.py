"""Independently reconstruct payload overlaps from the retained failed replay."""
from pathlib import Path
import hashlib,json
P=Path(__file__).resolve().parent

def save(name,value):
 with (P/name).open('x') as f:json.dump(value,f,indent=2,sort_keys=True);f.write('\n')

def overlap(read,ranges):
 start=read['offset'];end=start+read['returned_length']
 return sum(max(0,min(end,r['end'])-max(start,r['start'])) for r in ranges)

def main():
 path=P/'commands/diagnostic-pptx-cold-verified/output.log'
 line=path.read_text().strip()
 outer=json.loads(line.removeprefix('Error: '))
 inner=json.loads(outer.split(': Error: ',1)[1].strip())
 d=json.loads(inner.split('; diagnostic=',1)[1])
 reads=d['raw_reads'];boundary=d['open_read_count'];assert 0<=boundary<=len(reads)
 assert all(0<=r['returned_length']<=r['requested_length'] and r['offset']+r['returned_length']<=d['source_bytes'] for r in reads)
 assert len(reads)==d['counters']['read_calls']
 assert sum(r['returned_length'] for r in reads)==d['counters']['read_bytes']
 ranges=d['payload_ranges'];groups={'slide':ranges['slides'],'selected_slide':[ranges['selected_slide']], 'unselected_slide':ranges['unselected_slides'],'media':ranges['media']}
 phases={}
 for phase,items in [('all',reads),('open',reads[:boundary]),('query',reads[boundary:])]:
  row={'read_calls':len(items),'read_bytes':sum(r['returned_length'] for r in items)}
  for name,spans in groups.items():
   values=[overlap(r,spans) for r in items]
   row[name+'_payload_read_calls']=sum(v>0 for v in values)
   row[name+'_payload_read_bytes']=sum(values)
  phases[phase]=row
 assert phases['all']==d['counters']
 tail={'offset':d['source_bytes']-65536,'requested_length':65536,'returned_length':65536}
 indices=[i for i,r in enumerate(reads) if r==tail]
 assert len(indices)==1 and indices[0]<boundary
 for name,spans in groups.items():
  assert phases['open'][name+'_payload_read_bytes']==overlap(tail,spans)
 assert phases['query']['unselected_slide_payload_read_bytes']==0
 assert phases['query']['media_payload_read_bytes']==0
 selected=ranges['selected_slide']
 assert phases['query']['selected_slide_payload_read_bytes']==selected['end']-selected['start']
 save('diagnostic-decoded.json',d)
 save('diagnostic-analysis.json',dict(status='pass',child_mode=d['child_mode'],source_bytes=d['source_bytes'],source_sha256=d['source_sha256'],open_read_count=boundary,tail_read_index=indices[0],tail=tail,phases=phases,log_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),claim='Failed untimed replay only; no timed cold performance claim.'))
 print(json.dumps({'child_mode':d['child_mode'],'tail':tail,'phases':phases},indent=2))

if __name__=='__main__':main()
