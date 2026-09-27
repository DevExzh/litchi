"""Decode while exact executables remain available, retain compressed raw data."""
import gzip,shutil,subprocess,sys,time
import custody as c
lane=sys.argv[1];assert lane in ['perf','perf-fp']
out=c.P/lane;assert not (out/'decode-receipts.json').exists();rows=[]
for repeat in range(2):
 raw=out/f'{repeat}.data';dest=out/f'{repeat}.decoded';log=out/f'{repeat}.final-decode.log'
 assert raw.is_file() and not dest.exists()
 cmd=['perf','script','--ns','-i',str(raw)];start=time.time()
 with dest.open('w') as f,log.open('w') as e:r=subprocess.run(cmd,stdout=f,stderr=e)
 row={'repeat':repeat,'command':cmd,'started':start,'ended':time.time(),'exit_code':r.returncode,'raw':c.artifact(raw),'decoded':c.artifact(dest),'log':c.artifact(log)}
 rows.append(row);c.write(out/'decode-receipts.json',rows);assert r.returncode==0
compressed=[]
for source in sorted(out.iterdir()):
 if source.suffix not in ['.data','.decoded','.script']:continue
 receipt=c.artifact(source);dest=source.with_name(source.name+'.gz');assert not dest.exists()
 with source.open('rb') as f,dest.open('wb') as raw:
  with gzip.GzipFile(filename='',mode='wb',fileobj=raw,mtime=0) as z:shutil.copyfileobj(f,z)
 with gzip.open(dest,'rb') as f:assert __import__('hashlib').sha256(f.read()).hexdigest()==receipt['sha256']
 compressed.append({'original':receipt,'compressed':c.artifact(dest)})
 source.unlink()
c.write(out/'compression.json',compressed)
print(lane,'decoded and compressed',flush=True)
