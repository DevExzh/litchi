"""Archive explicit 0948 post-seal review/replay/cleanup supplements, without resealing v1."""
from pathlib import Path
import argparse, gzip, hashlib, io, json, tarfile

def sha(p):
    h=hashlib.sha256()
    with p.open('rb') as f:
        for b in iter(lambda:f.read(1048576),b''): h.update(b)
    return h.hexdigest()

ap=argparse.ArgumentParser(description=__doc__)
ap.add_argument('--packet',type=Path,required=True)
ap.add_argument('--archive-root',type=Path,required=True)
ap.add_argument('--output',type=Path,required=True)
a=ap.parse_args()
assert not a.output.exists()
files={}
packet_patterns=['tools/replay_0948_v*.py','receipts/root-host-v1.json',
 'review/protocol/review_protocol_0948_v9*','review/protocol/replay-adapter-independent-review-v*.json',
 'review/protocol/independent-leaf-controls-inline-v1.py','report-drafts/0948-rejection-report-v*.md']
for pat in packet_patterns:
    for p in a.packet.glob(pat): files['packet/'+p.relative_to(a.packet).as_posix()]=p
for p in a.archive_root.glob('*'):
    if p.is_file() and p.suffix in ('.json','.py','.stdout','.stderr') and not p.name.startswith('postseal-'):
        if p.name=='manifest-v1.json': continue
        files['archive/'+p.name]=p
for p in (a.archive_root/'postseal-review').rglob('*'):
    if p.is_file(): files['archive/'+p.relative_to(a.archive_root).as_posix()]=p
for p in files.values(): assert p.is_file() and not p.is_symlink(),p
entries=[{'path':n,'bytes':p.stat().st_size,'mode':p.stat().st_mode&0o777,'sha256':sha(p)} for n,p in sorted(files.items())]
manifest={'schema':'litchi.performance.0948.postseal-supplement.v1','entries':entries,'note':'These artifacts postdate the authoritative v1 archive; no original archive or measured artifact is replaced.'}
raw=(json.dumps(manifest,indent=2,sort_keys=True)+'\n').encode()
with a.output.open('xb') as f, gzip.GzipFile(filename='',mode='wb',fileobj=f,mtime=0,compresslevel=6) as gz, tarfile.open(fileobj=gz,mode='w|') as tar:
    m=tarfile.TarInfo('supplement-manifest.json');m.size=len(raw);m.mode=0o644;tar.addfile(m,io.BytesIO(raw))
    for e in entries:
        p=files[e['path']]
        assert sha(p)==e['sha256']
        t=tarfile.TarInfo(e['path']);t.size=e['bytes'];t.mode=e['mode']
        with p.open('rb') as inp:tar.addfile(t,inp)
with tarfile.open(a.output,'r:gz') as tar:
    members=tar.getmembers();assert {m.name for m in members}=={'supplement-manifest.json'}|set(files)
    assert tar.extractfile('supplement-manifest.json').read()==raw
    for e in entries:
        m=tar.getmember(e['path']);assert m.isfile() and m.size==e['bytes'] and m.mode==e['mode']
        h=hashlib.sha256()
        with tar.extractfile(m) as f:
            for b in iter(lambda:f.read(1048576),b''):h.update(b)
        assert h.hexdigest()==e['sha256']
receipt={'schema':'litchi.performance.0948.postseal-archive-receipt.v1','status':'verified',
 'archive':{'path':str(a.output),'bytes':a.output.stat().st_size,'sha256':sha(a.output)},
 'file_count':len(entries),'source_bytes':sum(e['bytes'] for e in entries),'manifest_sha256':hashlib.sha256(raw).hexdigest(),
 'helper':{'path':str(Path(__file__)),'sha256':sha(Path(__file__))},
 'command':['python3','-B',str(Path(__file__)),'--packet',str(a.packet),'--archive-root',str(a.archive_root),'--output',str(a.output)]}
with (a.archive_root/'postseal-archive-receipt-v1.json').open('x') as f:json.dump(receipt,f,indent=2,sort_keys=True);f.write('\n')
print(json.dumps(receipt,indent=2))
