#!/usr/bin/env python3
"""Remove only the owned 0508 scratch after all gates and process references close."""
import datetime
import json
from pathlib import Path
import shutil

HERE = Path(__file__).resolve().parent
SCRATCH = Path('/tmp/litchi-goal-0508')
assert SCRATCH.is_dir() and not SCRATCH.is_symlink()
assert not (HERE / 'cleanup.json').exists()
for lane in ['build', 'preflight', 'r1', 'r2', 'tests', 'clippy', 'rustdoc', 'doctests', 'check-features']:
    record = json.loads((HERE / f'{lane}-receipt.json').read_text())
    assert record['exit_code'] == 0 and record['source_unchanged'], lane
references = []
for process in Path('/proc').iterdir():
    if not process.name.isdigit():
        continue
    candidates = [process / 'exe', process / 'cwd']
    try:
        candidates.extend((process / 'fd').iterdir())
    except (OSError, PermissionError):
        pass
    for candidate in candidates:
        try:
            target = str(candidate.readlink())
        except OSError:
            continue
        if target == str(SCRATCH) or target.startswith(str(SCRATCH) + '/'):
            references.append({'process_link': str(candidate), 'target': target})
assert not references, references
files = 0
allocated = 0
seen = set()
for path in SCRATCH.rglob('*'):
    info = path.lstat()
    if not path.is_file() or path.is_symlink():
        continue
    files += 1
    key = (info.st_dev, info.st_ino)
    if key not in seen:
        allocated += info.st_blocks * 512
        seen.add(key)
shutil.rmtree(SCRATCH)
cache = HERE / '__pycache__'
if cache.exists():
    shutil.rmtree(cache)
record = {'status': 'pass', 'path': str(SCRATCH), 'removed': not SCRATCH.exists(),
          'regular_files': files, 'unique_file_inodes': len(seen),
          'unique_inode_allocated_bytes': allocated, 'live_references': references,
          'completed_utc': datetime.datetime.now(datetime.timezone.utc).isoformat()}
(HERE / 'cleanup.json').write_text(json.dumps(record, indent=2) + '\n')
print(json.dumps(record))
