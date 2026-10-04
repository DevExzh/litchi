"""Capture exact oracle stdout as deterministic gzip with logical-byte custody."""
import argparse,gzip,hashlib,json,pathlib,subprocess
P=pathlib.Path(__file__).resolve().parent
p=argparse.ArgumentParser();p.add_argument('role',choices=['baseline','candidate']);p.add_argument('binary');a=p.parse_args()
q=P/(a.role+'-parity.jsonl.gz');err=P/(a.role+'-parity.stderr');receipt=P/(a.role+'-parity.json')
assert not q.exists() and not err.exists() and not receipt.exists()
exe=pathlib.Path(a.binary);b=exe.read_bytes();binary={'path':str(exe.resolve()),'bytes':len(b),'sha256':hashlib.sha256(b).hexdigest()}
h=hashlib.sha256();total=0;lines=0;tail=b''
with q.open('xb') as raw,err.open('xb') as stderr,gzip.GzipFile(filename='',mode='wb',fileobj=raw,mtime=0) as z:
 child=subprocess.Popen([str(exe.resolve())],stdout=subprocess.PIPE,stderr=stderr)
 for block in iter(lambda:child.stdout.read(1<<20),b''):
  h.update(block);total+=len(block);lines+=block.count(b'\n');tail=(tail+block)[-4096:];z.write(block)
 code=child.wait()
last=json.loads(tail.rstrip().split(b'\n')[-1]);assert code==0 and last['complete'] and last['records']+2==lines,(code,last,lines)
r={'status':'pass','role':a.role,'binary':binary,'exit_code':code,'logical_bytes':total,'logical_sha256':h.hexdigest(),'lines':lines,'summary':last,'gzip':{'path':str(q),'bytes':q.stat().st_size,'sha256':hashlib.sha256(q.read_bytes()).hexdigest()}}
with receipt.open('x') as f:json.dump(r,f,indent=2,sort_keys=True);f.write('\n')
print(json.dumps(r,sort_keys=True))
