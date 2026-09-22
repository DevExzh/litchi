"""Losslessly archive raw perf bytes offline, including the rejected corrupt capture.

Perf's own compression is not used. Original receipt hashes remain raw-byte hashes.
"""
import gzip,hashlib,json,pathlib,shutil
P=pathlib.Path(__file__).resolve().parent
def digest_stream(f):
    h=hashlib.sha256()
    while block:=f.read(1024*1024):h.update(block)
    return h.hexdigest()
def digest(p):
    with p.open('rb') as f:return digest_stream(f)
if __name__=='__main__':
    assert not (P/'raw-archives.json').exists()
    rows=[]
    for raw in sorted(P.rglob('*.perf.data')):
        packed=raw.with_name(raw.name+'.gz');assert not packed.exists()
        original=digest(raw)
        with raw.open('rb') as src,packed.open('wb') as dst:
            with gzip.GzipFile(filename='',mode='wb',compresslevel=6,fileobj=dst,mtime=0) as gz:shutil.copyfileobj(src,gz)
        with gzip.open(packed,'rb') as f:assert digest_stream(f)==original
        rows.append({'raw':str(raw.relative_to(P)),'raw_bytes':raw.stat().st_size,'raw_sha256':original,'archive':str(packed.relative_to(P)),'archive_bytes':packed.stat().st_size,'archive_sha256':digest(packed)})
        (P/'raw-archives.json').write_text(json.dumps(rows,indent=2)+'\n');raw.unlink()
    print(f'PASS {len(rows)} raw captures archived with exact decoded-byte verification')
