"""Canonical non-inline frame decode preserves full Rust symbol identities."""
import gzip,shutil,subprocess
import custody as c
out=c.P/'perf-fp';assert not (out/'frame-receipts.json').exists();rows=[]
for repeat in range(2):
 raw=out/f'{repeat}.data';assert not raw.exists()
 with gzip.open(str(raw)+'.gz','rb') as f,raw.open('wb') as d:shutil.copyfileobj(f,d)
 expected=c.read(out/'receipts.json')[repeat]['raw'];assert c.artifact(raw)==expected
 dest=out/f'{repeat}.frames';log=out/f'{repeat}.frames.log'
 cmd=['perf','script','--no-inline','--ns','-i',str(raw)]
 with dest.open('w') as d,log.open('w') as e:r=subprocess.run(cmd,stdout=d,stderr=e)
 row={'repeat':repeat,'command':cmd,'exit_code':r.returncode,'raw':expected,'frames':c.artifact(dest),'log':c.artifact(log)}
 assert r.returncode==0
 raw.unlink()
 compressed=dest.with_name(dest.name+'.gz')
 with dest.open('rb') as f,compressed.open('wb') as raw:
  with gzip.GzipFile(filename='',mode='wb',fileobj=raw,mtime=0) as z:shutil.copyfileobj(f,z)
 row['compressed']=c.artifact(compressed);dest.unlink();rows.append(row)
 c.write(out/'frame-receipts.json',rows)
print('Full native frame symbols decoded',flush=True)
