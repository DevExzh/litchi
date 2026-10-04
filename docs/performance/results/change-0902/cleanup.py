"""Remove only this batch's marked source and build roots; never follow fixture links."""
import pathlib,json,os,shutil,datetime,subprocess,hashlib
p=pathlib.Path(__file__).resolve().parent
roots=[pathlib.Path('/tmp/litchi-0902-src'),pathlib.Path('/home/zhuhe/code/litchi/target/litchi-0902-build')]
records=[]
assert json.loads((p/'decision.json').read_text())['status']=='rejected'
assert json.loads((p/'raw-replay.json').read_text())['status']=='pass'
for root in roots:
 assert root.is_dir() and not root.is_symlink() and root.resolve()==root
 marker=root/'.litchi-0902-owned.json';assert json.loads(marker.read_text())['owner']=='change-0902'
 files=0;size=0;links=[]
 for current,dirs,names in os.walk(root,followlinks=False):
  for name in dirs+names:
   f=pathlib.Path(current)/name
   if f.is_symlink():links.append({'path':str(f),'target':os.readlink(f)})
  for name in names:
   f=pathlib.Path(current)/name
   if not f.is_symlink():files+=1;size+=f.stat().st_size
 records.append({'path':str(root),'files':files,'logical_bytes':size,'symlinks_unlinked_without_following':links,'marker_sha256':hashlib.sha256(marker.read_bytes()).hexdigest()})
for root in roots:shutil.rmtree(root);assert not root.exists()
for row in records:row['removed']=True
with (p/'cleanup.json').open('x') as f:json.dump({'status':'pass','roots':records,'worktrees_created':0,'worktrees_removed':0,'retained_results_binaries':'intentional hash-bound experiment artifacts, not temporary build directories','finished_utc':datetime.datetime.now(datetime.timezone.utc).isoformat()},f,indent=2);f.write('\n')
print('removed',sum(r['logical_bytes'] for r in records),'logical bytes from two owned roots; no worktrees created or removed')
