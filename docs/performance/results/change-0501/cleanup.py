#!/usr/bin/env python3
"""Remove only the dedicated 0501 scratch after retaining replay executables."""
import hashlib
import json
import os
from pathlib import Path
import shutil

HERE=Path(__file__).resolve().parent
CACHE=Path('/tmp/litchi-goal-0501')
assert CACHE.is_dir() and not CACHE.is_symlink()
retained=[]
for receipt_name in ['before-freeze.json','after-freeze.json']:
    receipt=json.loads((HERE/receipt_name).read_text())
    binary=receipt['binary']
    p=Path(binary['path'])
    assert p.parent == CACHE/'retained'
    assert hashlib.sha256(p.read_bytes()).hexdigest()==binary['sha256']
    retained.append({'path':str(p),'sha256':binary['sha256'],'bytes':p.stat().st_size})
for p in sorted((CACHE/'profiles').rglob('perf.data')):
    assert p.is_file() and not p.is_symlink()
    retained.append({'path':str(p),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size})
ancestors=set()
pid=os.getpid()
while pid>0 and pid not in ancestors:
    ancestors.add(pid)
    try:pid=int((Path('/proc')/str(pid)/'stat').read_text().rsplit(')',1)[1].split()[1])
    except (OSError,ValueError,IndexError):break
active=[]
inaccessible=0
scanned=0
for proc in Path('/proc').iterdir():
    if not proc.name.isdigit() or int(proc.name) in ancestors:continue
    try:
        if proc.stat().st_uid != os.getuid():continue
        scanned+=1
        cmd=(proc/'cmdline').read_bytes()
        env=(proc/'environ').read_bytes()
        if any(marker in cmd or marker in env for marker in [str(CACHE).encode(), b'/tmp/litchi-goal-0501-target']):active.append(int(proc.name))
    except (PermissionError,FileNotFoundError,ProcessLookupError):inaccessible+=1
assert not active, f'0501 processes still active: {active}'
physical_target = Path('/tmp/litchi-goal-0501-target')
assert physical_target.is_dir() and not physical_target.is_symlink()
paths=[p for p in CACHE.iterdir() if p.name not in {'retained', 'profiles'}] + [physical_target]
seen_inodes = set()
removed=[]
for p in paths:
    if not p.exists() and not p.is_symlink():continue
    assert p.stat().st_uid==os.getuid(),str(p)
    allocated=0;files=0
    entries=[p]
    if p.is_dir() and not p.is_symlink():entries+=list(p.rglob('*'))
    for item in entries:
        stat=item.lstat()
        inode=(stat.st_dev, stat.st_ino)
        if inode not in seen_inodes:
            allocated+=stat.st_blocks*512
            seen_inodes.add(inode)
        if item.is_file() or item.is_symlink():files+=1
    if p.is_dir() and not p.is_symlink():shutil.rmtree(p)
    else:p.unlink()
    assert not p.exists() and not p.is_symlink()
    removed.append({'path':str(p),'allocated_bytes':allocated,'files':files})
result={'scope':'dedicated 0501 scratch only; shared target and unrelated work untouched',
        'allocation_accounting':'unique device/inode block counts across removed paths',
        'retention_scope':'frozen binaries and all successful/failed raw perf recordings remain in local tmpfs for replay; committed source archives and text exports are portable',
        'same_user_processes_scanned':scanned,'inaccessible_or_exited_processes':inaccessible,
        'active_matching_processes':active,'retained':retained,'removed':removed,
        'removed_allocated_bytes':sum(p['allocated_bytes'] for p in removed)}
(HERE/'cleanup.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps({'removed_allocated_bytes':result['removed_allocated_bytes'],'retained_replay_files':len(retained)}))
