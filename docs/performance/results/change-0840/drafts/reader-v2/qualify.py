"""Independent custody validation around report and CFB content readers."""
import driver as d
import readers as r

def bind(x):
 assert d.desc(x['path'])==x,x['path']

def validate(folder):
 start=d.read(folder/'started.json');receipt=d.read(folder/'receipt.json')
 assert receipt['exit_code']==0 and receipt['finished_unix']>=start['started_unix']
 expected={str(p) for p in folder.iterdir() if p.is_file() and p.name!='receipt.json'}
 assert {x['path'] for x in receipt['files']}==expected
 for x in receipt['files']:bind(x)
 bind(start['build']);build=d.read(start['build']['path']);assert build['binary']==start['binary'];bind(build['freeze']);bind(build['receipt'])
 plan=d.read(d.P/'plan.json');lane=start['lane'];cfg=plan['qualification' if lane=='allocation-preflight' else lane];case=start['case']
 assert case in plan['cases'] and start['leg'] in ['before','after']
 kind='allocation' if lane in ['allocation','allocation-preflight'] else 'native'
 assert start['build']['path']==str(d.P/f"build-{start['leg']}-{kind}.json")
 args=['taskset','-c',str(plan['native']['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(folder/'time.txt'),start['binary']['path'],'--case',case,'--samples',str(cfg['samples']),'--warmup',str(cfg['warmup']),'--output',str(folder/'report.json')]
 if lane=='observer':args+=['--observe']
 if lane in ['qualification','observer']:args+=['--artifact',str(folder/'output.cfb')]
 assert start['argv']==args
 value=r.validate_report(folder/'report.json',case,cfg['samples'],cfg['warmup'],observer=lane=='observer',allocation=lane in ['allocation','allocation-preflight'])
 counters=(folder/'time.txt').read_text().split();assert len(counters)==3 and all(x.isdigit() for x in counters)
 value.update(lane=lane,leg=start['leg'],case=case,block=start['block'],rss_kib=int(counters[0]),minor_faults=int(counters[1]),major_faults=int(counters[2]),started_unix=start['started_unix'],finished_unix=receipt['finished_unix'],folder=str(folder))
 if lane in ['qualification','observer']:
  data=(folder/'output.cfb').read_bytes();value['cfb']=r.validate_cfb_artifact(case,data)
  assert value['artifact_identity']['output_sha256']==d.sha(folder/'output.cfb') and value['artifact_identity']['output_bytes']==len(data)
  value['cfb'].pop('_stream_bytes',None)
 return value

def all_rows():
 rows=[validate(p.parent) for p in sorted((d.P/'runs').rglob('receipt.json'))]
 ordered=sorted(rows,key=lambda x:x['started_unix'])
 assert all(a['finished_unix']<=b['started_unix'] for a,b in zip(ordered,ordered[1:]))
 for case in d.read(d.P/'plan.json')['cases']:
  subset=[x for x in rows if x['case']==case]
  # Artifact equality is checked independently from each report's identity.
  artifacts=[d.Path(x['folder'])/'output.cfb' for x in subset if x['lane'] in ['qualification','observer']]
  if artifacts:
   first=artifacts[0].read_bytes()
   assert all(p.read_bytes()==first for p in artifacts),case
   assert all(x['artifact_identity']['output_sha256']==d.sha(artifacts[0]) and x['artifact_identity']['output_bytes']==len(first) for x in subset),case
  if subset:assert all(x['corpus_identity']==subset[0]['corpus_identity'] for x in subset),case
 return rows

if __name__=='__main__':
 rows=all_rows();print('qualification PASS',len(rows),'reports')
