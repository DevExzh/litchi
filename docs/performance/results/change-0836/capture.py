"""Serial native controls, inherited CPU profiles and separate syscall traces."""
import gzip,shutil,sys
import driver as d
import runner

def pack(path):
 dest=d.P/(path.name+'.gz')
 with path.open('rb') as src,dest.open('xb') as out:
  with gzip.GzipFile(filename='',mode='wb',fileobj=out,mtime=0) as z:shutil.copyfileobj(src,z)
 return d.desc(dest)
def execute(kind,index,spec):
 stage=spec['stage'];label=f'{kind}-{index:02}';plan=d.read(d.P/'plan.json');case=plan['case']
 binary=d.read(d.P/f'build-{stage}.json')['binary'];assert d.desc(binary['path'])==binary
 report=d.P/f'{label}.json';assert not report.exists()
 args=['taskset','-c','12',binary['path'],'--case',case,'--samples',str(spec['samples']),'--warmup',str(spec['warmup']),'--filesystem-cache',spec['state'],'--filesystem-root',str(d.SCRATCH),'--json',str(report)]
 raw=d.TARGET/f'{label}.data' if kind=='profile' else d.TARGET/f'{label}.trace'
 if kind=='profile':args=['perf','record','--no-buildid-cache','-e','cycles:u','-F','997','--call-graph','fp','-o',str(raw),'--',*args]
 if kind=='trace':args=['strace','-f','-qq','-ttt','-T','-yy','-e','trace=fsync,fdatasync,rename,renameat,renameat2','-o',str(raw),'--',*args]
 receipt=d.run(stage,label,args)
 row=dict(**spec,label=label,case=case,exit_code=receipt['exit_code'],report=d.desc(report) if report.exists() else None,receipt=d.desc(d.P/f'commands/{label}/receipt.json'))
 if raw.exists():row.update(raw=d.desc(raw),raw_gzip=pack(raw))
 if receipt['exit_code']!=0:
  d.write(d.P/f'{label}-failed.json',row);raise RuntimeError(label+' failed; no retry')
 if kind=='profile':
  decoded=d.TARGET/f'{label}.script';decode_label=f'decode-{label}'
  r=runner.run(stage,decode_label,['perf','script','--no-inline','--ns','--show-lost-events','-i',str(raw)],decoded);assert r['exit_code']==0
  row.update(decoded=pack(decoded),decoded_plain=d.desc(decoded),decode_receipt=d.desc(d.P/f'commands/{decode_label}/receipt.json'))
  id_label=f'buildid-{label}';r=d.run(stage,id_label,['perf','buildid-list','-i',str(raw)]);assert r['exit_code']==0
  row.update(buildid=d.desc(d.P/f'commands/{id_label}/output.log'),buildid_receipt=d.desc(d.P/f'commands/{id_label}/receipt.json'))
 return row

def main(kind):
 mapping={'qualification':'qualification','native':'native','profile':'perf','trace':'trace'}
 assert kind in mapping
 if kind!='qualification':
  admission=d.read(d.P/'admission.json');assert admission['status']=='pass'
  for n,b in admission['inputs'].items():assert d.desc(d.P/n)==b,n
 plan=d.read(d.P/'plan.json');rows=[]
 for i,spec in enumerate(plan[mapping[kind]]):rows.append(execute(kind,i,spec))
 name={'profile':'profiles','trace':'traces'}.get(kind,kind)
 d.write(d.P/f'{name}.json',dict(status='commands_pass',rows=rows,report_count=len(rows),sample_count=sum(r['samples']*len(r['state'].split(',')) for r in rows)))
 print(kind,'commands PASS',flush=True)
if __name__=='__main__':main(sys.argv[1])
