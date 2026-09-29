"""Bind exact commands/artifacts and independently validate workload results."""
import driver as d
import sys

def bind(value):
 assert d.desc(value['path'])==value,value['path']
 return d.Path(value['path'])

def receipt(folder):
 end=d.read(folder/'receipt.json');start=d.read(folder/'started.json');assert end['exit_code']==0
 for x in end['artifacts']:bind(x)
 assert {d.Path(x['path']).relative_to(folder).as_posix() for x in end['artifacts']}=={p.relative_to(folder).as_posix() for p in folder.rglob('*') if p.is_file() and p.name!='receipt.json'}
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
 argv+=[build['binary']['path'],'--mode',row['mode'],'--shape',row['shape'],'--samples',str(row['samples']),'--warmup',str(row['warmup']),'--output',str(folder/'report.json')]
 if end['lane']=='qualification':argv+=['--artifact',str(folder/'artifact.zip')]
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
 import reader as r
 rows=folders(lane,leg);out=[]
 for item in rows:
  row=item['row'];folder=item['folder']
  artifact_path=folder/'artifact.zip' if lane=='qualification' else None
  expected=None
  if lane=='native':
   qual=d.P/'runs'/f"qualification-00-{row['shape']}-{row['mode']}-{'leg-a' if row['leg']=='before' else 'leg-b'}"/'report.json'
   expected=d.read(qual)
  validated=r.validate_report(folder/'report.json',row['shape'],row['mode'],row['samples'],row['warmup'],artifact_path=artifact_path,expected=expected)
  out.append(dict(**{k:v for k,v in item.items() if k!='folder'},folder=str(folder),**validated))
 return out

def qualification():
 rows=validate_items('qualification')
 for case in d.read(d.P/'plan.json')['cases']:
  pair=[x for x in rows if all(x['row'][k]==v for k,v in case.items())];assert len(pair)==2
  assert pair[0]['source']==pair[1]['source'] and pair[0]['output']==pair[1]['output']
  a=d.Path(pair[0]['folder'])/'artifact.zip';b=d.Path(pair[1]['folder'])/'artifact.zip'
  assert a.read_bytes()==b.read_bytes()
 import reader as r
 import io,zipfile
 unchanged=[]
 def local_spans(data):
  with zipfile.ZipFile(io.BytesIO(data)) as z:
   infos=sorted(z.infolist(),key=lambda x:x.header_offset)
   return {v.filename:data[v.header_offset:(infos[i+1].header_offset if i+1<len(infos) else z.start_dir)] for i,v in enumerate(infos)}
 for shape in ['tiny','xml','random','repeated']:
  for leg in ['before','after']:
   group={x['row']['mode']:x for x in rows if x['row']['shape']==shape and x['row']['leg']==leg}
   for row in group.values():
    assert row['source']==group['write']['output']
    assert row['expected_changed_archive']==group['changed']['output']
    assert row['expected_changed_logical']==group['changed']['logical']
   blobs={mode:(d.Path(x['folder'])/'artifact.zip').read_bytes() for mode,x in group.items()}
   assert blobs['write']==blobs['read']
   old={x['name']:x for x in r.archive_manifest(blobs['write'])['members']}
   new={x['name']:x for x in r.archive_manifest(blobs['changed'])['members']}
   assert set(old)==set(new)
   oldspans,newspans=local_spans(blobs['write']),local_spans(blobs['changed'])
   target='custom/payload.bin';assert old[target]['logical']!=new[target]['logical']
   names=sorted(set(old)-{target})
   for name in names:
    assert oldspans[name]==newspans[name],(shape,leg,name,'local record changed')
    for field in ['central','logical','raw','crc32']:assert old[name][field]==new[name][field]
   unchanged.append(dict(shape=shape,leg=leg,unchanged_members=names,exact_local_spans=True,exact_compressed_payloads=True,exact_central_metadata_except_offsets=True))
 return dict(status='pass',reports=len(rows),rows=rows,unchanged=unchanged,reader=d.desc(d.P/'reader.py'))

if __name__=='__main__':
 assert sys.argv[1:]==['qualification']
 result=qualification();d.write(d.P/'qualification.json',result);print('qualification PASS',result['reports'])
