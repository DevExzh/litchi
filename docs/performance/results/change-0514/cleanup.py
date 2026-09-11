"""Remove only this completed batch's owned scratch after process checks."""
import datetime,json,os,shutil,subprocess
from pathlib import Path
here=Path(__file__).resolve().parent
scratch=Path('/tmp/litchi-goal-0514')
assert scratch.is_dir() and not scratch.is_symlink()
assert scratch.resolve()==scratch
refs=[]
for proc in Path('/proc').iterdir():
    if not proc.name.isdigit() or int(proc.name)==os.getpid():continue
    try:
        paths=[proc/'cwd',proc/'exe',*list((proc/'fd').iterdir())]
        for path in paths:
            try:
                value=os.readlink(path)
                if value==str(scratch) or value.startswith(str(scratch)+'/'):refs.append([proc.name,str(path),value])
            except (FileNotFoundError,PermissionError,OSError):pass
        maps=(proc/'maps').read_text()
        if str(scratch)+'/' in maps:refs.append([proc.name,'maps','owned scratch mapped'])
    except (FileNotFoundError,PermissionError,ProcessLookupError):pass
assert not refs,refs
count=0;allocated=0;inodes=set()
for path in scratch.rglob('*'):
    if path.is_file():
        count+=1;info=path.stat();identity=(info.st_dev,info.st_ino)
        if identity not in inodes:allocated+=info.st_blocks*512;inodes.add(identity)
owned_worktrees=[]
for line in subprocess.check_output(['git','worktree','list','--porcelain']).decode().splitlines():
    if not line.startswith('worktree '):continue
    candidate=Path(line[len('worktree '):]).resolve()
    if candidate.is_relative_to(scratch):
        owned_worktrees.append(str(candidate))
for worktree in owned_worktrees:
    subprocess.run(['git','worktree','remove','--force',worktree],check=True)
shutil.rmtree(scratch)
assert not scratch.exists()
receipt={'scope':'owned /tmp/litchi-goal-0514 only','removed':True,'removed_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'file_count':count,'unique_inode_allocated_bytes':allocated,'process_references_before_removal':refs,'removed_owned_registered_worktrees':owned_worktrees,'retained_evidence':'docs/performance/results/change-0514; source/binary hashes, raw durations/profiles/logs retained; executable and target files removed'}
(here/'cleanup.json').write_text(json.dumps(receipt,indent=2)+'\n')
print(json.dumps(receipt))
