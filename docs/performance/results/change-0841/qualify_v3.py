"""Bind exact commands/artifacts and independently validate workload results."""
import driver as d
import sys

def bind(value):
 assert d.desc(value['path'])==value,value['path']
 return d.Path(value['path'])

def receipt(folder):
 end=d.read(folder/'receipt.json');start=d.read(folder/'started.json');assert end['exit_code']==0
 for x in end['artifacts']:bind(x)
 assert {d.Path(x['path']).name for x in end['artifacts']}=={p.name for p in folder.iterdir() if p.is_file() and p.name!='receipt.json'}
 command=d.read(bind(end['command']));assert command['exit_code']==0
 cstart=d.read(d.Path(end['command']['path']).parent/'started.json');bind(command['log'])
 assert command['argv']==cstart['argv']==start['argv']
 assert start['started_unix']<=command['started_unix']<=command['finished_unix']<=end['finished_unix']
 row=end['row'];assert row==start['row'] and end['lane']==start['lane']
 build=d.read(bind(start['build']));assert build['binary']==start['binary']
 assert build==d.read(d.P/f"build-{row['leg']}.json")
 frozen=d.read(bind(build['freeze']));assert frozen['source']==d.read(d.P/f"source-{row['leg']}.json")['source']
 assert all(d.sha(d.P/n)==h for n,h in frozen['probe'].items())
 for name in ['driver','builder','plan']:bind(frozen[name])
 argv=['taskset','-c',str(d.read(d.P/'plan.json')['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(folder/'time.txt')]
 if end['lane']=='profiles':argv+=['valgrind',*d.read(d.P/'plan.json')['profile_options'],'--callgrind-out-file='+str(folder/'profile.callgrind')]
 argv+=[build['binary']['path'],'--mode',row['mode'],'--shape',row['shape'],'--samples',str(row['samples']),'--warmup',str(row['warmup']),'--output',str(folder/'report.json')]
 if end['lane']=='qualification':argv+=['--artifact',str(folder/'artifact.ppt')]
 assert argv==command['argv'],folder
 counts=(folder/'time.txt').read_text().split();assert len(counts)==3 and all(x.isdigit() for x in counts)
 return dict(lane=end['lane'],row=row,rss_kib=int(counts[0]),minor_faults=int(counts[1]),major_faults=int(counts[2]),started=command['started_unix'],finished=command['finished_unix'],folder=folder)

def folders(lane,leg=None):
 plan=d.read(d.P/'plan.json');expected=[x for x in plan[lane] if leg is None or x['leg']==leg]
 actual=[]
 for row in expected:
  name=f"{lane}-{row.get('block',0):02}-{row['shape']}-{row['mode']}-{'leg-a' if row['leg']=='before' else 'leg-b'}"
  item=receipt(d.P/'runs'/name);assert item['row']==row;actual.append(item)
 return actual

def validate_items(lane,leg=None):
 import readers_v2 as r
 rows=folders(lane,leg);out=[]
 for item in rows:
  row=item['row'];folder=item['folder']
  if lane=='qualification':artifact=folder/'artifact.ppt';expected=None
  else:
   artifact=None
   qual=d.P/'runs'/f"qualification-00-{row['shape']}-{row['mode']}-{'leg-a' if row['leg']=='before' else 'leg-b'}"/'report.json'
   expected=d.read(qual)['artifact']
  validated=r.validate_report(folder/'report.json',row['shape'],row['mode'],row['samples'],row['warmup'],artifact_path=artifact,expected_output=expected)
  summary={k:validated[k] for k in ['source','output','wall_ns','p50_ns','p95_ns','p99_ns','mean_ns']}
  if lane=='qualification':
   summary['artifact']=d.desc(artifact)
  if lane=='profiles':
   raw=r.validate_profile(folder/'profile.callgrind.1',folder/'profile.callgrind',shape=row['shape'],report_path=folder/'report.json', command=d.read(bind(d.read(folder/'receipt.json')['command']))['argv'])
   command=d.read(bind(d.read(folder/'receipt.json')['command']))['argv']
   expected_cmd=' '.join(command[command.index(d.read(d.P/f"build-{row['leg']}.json")['binary']['path']):])
   parsed=r.parse_raw(folder/'profile.callgrind.1');assert parsed['header']['command'].lstrip()==expected_cmd
   summary['profile']=raw
  out.append(dict(**{k:v for k,v in item.items() if k!='folder'},folder=str(folder),**summary))
 return out

def qualification():
 rows=validate_items('qualification')
 for case in d.read(d.P/'plan.json')['cases']:
  pair=[x for x in rows if all(x['row'][k]==v for k,v in case.items())];assert len(pair)==2
  assert pair[0]['source']==pair[1]['source'] and pair[0]['output']==pair[1]['output']
  assert d.Path(pair[0]['artifact']['path']).read_bytes()==d.Path(pair[1]['artifact']['path']).read_bytes()
 # All modes and legs of a shape must produce exactly identical bytes.
 for shape in {x['row']['shape'] for x in rows}:
  selected=[x for x in rows if x['row']['shape']==shape]
  assert len({x['output']['sha256'] for x in selected})==1
 return dict(status='pass',reports=len(rows),rows=rows,reader=d.desc(d.P/'readers_v2.py'))

if __name__=='__main__':
 assert sys.argv[1:] == ['qualification']
 result=qualification();d.write(d.P/'qualification-v2.json',result);print('qualification PASS',result['reports'])
