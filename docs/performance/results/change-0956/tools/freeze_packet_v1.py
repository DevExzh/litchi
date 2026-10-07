"""Create or verify a completed packet inventory; the manifest stays outside packet."""
import argparse,hashlib,json,pathlib
ap=argparse.ArgumentParser();ap.add_argument('packet',type=pathlib.Path);ap.add_argument('manifest',type=pathlib.Path);ap.add_argument('--verify',action='store_true');a=ap.parse_args()
r=a.packet.resolve();files=[]
assert not a.packet.is_symlink() and r.is_dir()
for p in sorted(r.rglob('*')):
 assert not p.is_symlink(),str(p)
 if p.is_file():
  h=hashlib.sha256()
  with p.open('rb') as f:
   for b in iter(lambda:f.read(1048576),b''):h.update(b)
  files.append({'path':str(p.relative_to(r)),'bytes':p.stat().st_size,'sha256':h.hexdigest()})
terminal=json.loads((r/'capture-terminal.json').read_text())
assert terminal['status']=='completed' and terminal['children_started']==terminal['children_expected']==144
record={'schema':'litchi.performance.0956.completed-packet-inventory.v1','packet':str(r),'file_count':len(files),'bytes':sum(f['bytes'] for f in files),'files':files}
data=(json.dumps(record,indent=2,sort_keys=True)+'\n').encode()
if a.verify:assert a.manifest.read_bytes()==data,'packet inventory changed'
else:
 with a.manifest.open('xb') as f:f.write(data)
print(json.dumps({'status':'verified' if a.verify else 'frozen','manifest':str(a.manifest),'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest(),'files':len(files),'packet_bytes':record['bytes']}))
