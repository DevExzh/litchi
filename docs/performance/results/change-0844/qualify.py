"""Bind actual commands/artifacts before independent result validation."""
import driver as d
import sys

def bind(item):
 assert d.desc(item['path'])==item,item['path']
 return d.Path(item['path'])

def folder_for(lane,row):return d.P/'runs'/f"{lane}-{row['block']:02}-{row['shape']}-{row['parts']:04}-{row['mode']}-{row['leg']}"

def receipt(folder):
 end=d.read(folder/'receipt.json');start=d.read(folder/'started.json');assert end['exit_code']==0
 for item in end['artifacts']:bind(item)
 assert {d.Path(x['path']).name for x in end['artifacts']}=={p.name for p in folder.iterdir() if p.is_file() and p.name!='receipt.json'}
 command=d.read(bind(end['command']));cstart=d.read(d.Path(end['command']['path']).parent/'started.json');bind(command['log'])
 assert command['exit_code']==0 and command['argv']==cstart['argv']==start['argv']
 assert start['started_unix']<=command['started_unix']==cstart['started_unix']<=command['finished_unix']<=end['finished_unix']
 row=end['row'];lane=end['lane'];assert row==start['row'] and lane==start['lane']
 build=d.read(bind(start['build']));assert build==d.read(d.P/f"build-{row['leg']}.json") and build['binary']==start['binary']
 frozen=d.read(bind(build['freeze']));assert frozen['source']==d.read(d.P/f"source-{row['leg']}.json")['source']
 assert all(d.sha(d.P/n)==h for n,h in frozen['probe'].items())
 for key in ['driver','builder','plan']:bind(frozen[key])
 plan=d.read(d.P/'measurement-plan.json');assert row in plan[lane]
 argv=['taskset','-c',str(plan['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(folder/'time.txt')]
 if lane=='heap':argv+=['heaptrack','-o',str(folder/'trace')]
 argv+=[build['binary']['path'],'--shape',row['shape'],'--parts',str(row['parts']),'--mode',row['mode'],'--samples',str(row['samples']),'--warmup',str(row['warmup']),'--output',str(folder/'report.json')]
 if lane=='qualification':argv+=['--artifact',str(folder/'artifact.zip')]
 assert command['argv']==argv
 counts=(folder/'time.txt').read_text().split();assert len(counts)==3 and all(x.isdigit() for x in counts) and int(counts[0])>0
 return dict(lane=lane,row=row,folder=str(folder),started=command['started_unix'],finished=command['finished_unix'],rss_kib=int(counts[0]),minor_faults=int(counts[1]),major_faults=int(counts[2]))

def validate_items(lane,leg=None):
 import reader as r
 out=[]
 for row in d.read(d.P/'measurement-plan.json')[lane]:
  if leg is not None and row['leg']!=leg:continue
  folder=folder_for(lane,row);item=receipt(folder)
  expected=None if lane=='qualification' else d.read(folder_for('qualification',dict(row,block=0))/'report.json')
  validated=r.validate_report(folder/'report.json',row['parts'],row['mode'],row['samples'],row['warmup'],artifact_path=folder/'artifact.zip' if lane=='qualification' else None,source_path=folder/'source.zip' if lane=='qualification' else None,expected=expected,shape=row['shape'])
  out.append(dict(**item,**validated))
 return out

def qualification(leg=None):
 rows=validate_items('qualification',leg)
 for shape in ['shared','root','deep-shared','deep-unique']:
  for parts in [4,64,1024,4096]:
   selected=[x for x in rows if x['row']['parts']==parts and x['row']['shape']==shape];assert len(selected)==(2 if leg else 4)
   for name in ['source.zip','artifact.zip']:
    expected=(d.Path(selected[0]['folder'])/name).read_bytes()
    assert all((d.Path(x['folder'])/name).read_bytes()==expected for x in selected)
 return dict(status='pass',reports=len(rows),rows=rows,reader=d.desc(d.P/'reader.py'))

if __name__=='__main__':
 leg=sys.argv[1] if len(sys.argv)>1 else None;assert leg in [None,'before','after']
 value=qualification(leg);out=d.P/f"qualification{'-'+leg if leg else ''}.json"
 if out.exists():assert d.read(out)==value
 else:d.write(out,value)
 print('Independent qualification PASS',len(value['rows']),'reports')
