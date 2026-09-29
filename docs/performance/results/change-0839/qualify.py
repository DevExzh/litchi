"""Independent correctness and custody replay for every retained process."""
import driver as d
import readers as r
import gzip,json,sys
PLAN=d.read(d.P/'plan.json')
def bind(x):
 p=d.Path(x['path']);marker='/docs/performance/results/change-0839/'
 if marker in str(p):p=d.P/str(p).split(marker,1)[1]
 assert p.is_file() and not p.is_symlink() and p.stat().st_size==x['bytes'] and d.sha(p)==x['sha256'],p

def validate(folder):
 s=d.read(folder/'started.json');receipt=d.read(folder/'receipt.json')
 assert receipt['exit_code']==0 and receipt['argv']==s['argv'] and receipt['started_unix']==s['started_unix'] and receipt['finished_unix']>=s['started_unix']
 for x in receipt['artifacts']:bind(x)
 kind=s['kind'];leg=s['leg'];name=leg if kind=='native' else leg+'-'+kind
 build=d.read(d.P/f'build-{name}.json');assert s['binary']==build['binary']
 argv=s['argv'];program=build['binary']['path'];assert argv.count(program)==1
 tail=argv[argv.index(program)+1:];expected=[]
 for k,v in s['case'].items():expected+=['--'+k.replace('_','-'),str(v)]
 assert tail[:-2]==expected+['--samples',str(s['samples']),'--warmup',str(s['warmup'])]
 assert tail[-2]=='--output' and tail[-1].endswith('/'+folder.name+'/report.json')
 affinity=','.join(map(str,PLAN['affinity'])) if s['affinity']=='all' else str(PLAN['affinity'][0])
 assert argv[:3]==['taskset','-c',affinity]
 result=r.validate_report(folder/'report.json',s['case'],s['samples'],s['warmup'],feature=kind=='observer',allocation=kind=='allocator')
 fields=['case','corpus_fingerprint','verification_ok','resources','source_metrics','p50_ns','p95_ns','p99_ns','allocation_metrics']
 row={k:result[k] for k in fields};row.update({k:s[k] for k in ['lane','leg','block','index','kind','affinity','protocol','samples','warmup']})
 row['path']=str(folder.relative_to(d.P));row['started_unix']=s['started_unix'];row['finished_unix']=receipt['finished_unix']
 if s['lane'].startswith('memory'):
  snapshots=json.loads(gzip.decompress((folder/'snapshots.json.gz').read_bytes()));rss=r.parse_snapshot_series(snapshots)
  seq=[('startup',2**64-1),('corpus_ready',2**64-1),('warmup_done',2**64-1)]
  for sample in sorted({0,s['samples']-1}):seq +=[(n,sample) for n in ['package_ready','after_preload','after_operation','after_batch_drop','after_package_drop']]
  seq +=[(n,2**64-1) for n in ['samples_done','report_written','report_dropped']]
  assert [(x['phase'],x['sample']) for x in rss]==seq
  usage=d.read(folder/'usage.json');assert usage['exit_code']==0 and usage['child_pid']==rss[0]['pid']
  for raw,parsed in zip(snapshots,rss):
   stat=raw['raw']['stat'];assert parsed['starttime']==int(stat[stat.rfind(')')+2:].split()[19])
  if s['lane']=='memory':
   expected_launcher=d.read(d.P/'launcher-quality.json')['binary']
   assert s['launcher']==expected_launcher and all(x['launcher']==expected_launcher for x in snapshots)
  row['residency']=rss;row['usage']=usage
  row['memory_metrics']={phase:max(x['smaps_sum_rss_kib'] for x in rss if x['phase']==phase) for phase in ['after_preload','after_operation','after_batch_drop','after_package_drop']}
  row['memory_metrics']['maximum observed RSS']=max(x['smaps_sum_rss_kib'] for x in rss)
 else:
  values=(folder/'time.txt').read_text().split();assert len(values)==3 and all(x.isdigit() for x in values)
  row['rss_kib'],row['minor_faults'],row['major_faults']=map(int,values)
 return row

def all_rows():
 rows=[validate(p) for p in sorted((d.P/'runs').glob('*/*')) if p.is_dir()]
 ordered=sorted(rows,key=lambda x:x['started_unix'])
 for a,b in zip(ordered,ordered[1:]):assert a['finished_unix']<=b['started_unix'],'overlapping captures'
 corpus={}
 for row in rows:
  shape=row['case']['shape'];prior=corpus.setdefault(shape,row['corpus_fingerprint']);assert prior==row['corpus_fingerprint']
 return rows
if __name__=='__main__':
 rows=all_rows();label=sys.argv[1]
 d.write(d.P/f'qualification-replay-{label}.json',dict(status='pass',processes=len(rows),samples=sum(x['samples'] for x in rows),reader=d.desc(d.P/'readers.py'),rows=rows))
 print('independent qualification PASS',len(rows),'reports')
