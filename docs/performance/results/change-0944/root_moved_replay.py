#!/usr/bin/env python3
"""Copy the replay closure, hide its original root, and require identical output."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import time

origin = Path('/home/zhuhe/litchi-performance-0944')
work = Path(__file__).resolve().parent
moved = work / 'moved-replay'
hidden = work / 'original-hidden'
assert origin.is_dir() and not moved.exists() and not hidden.exists()
started = time.time()
files = 0
size = 0

def ignored(directory, names):
    if Path(directory) == origin:
        return set(names) & {'source', 'target-native', 'tmp'}
    # Only actual tooling bytecode caches, not arbitrary source components.
    relative = Path(directory).relative_to(origin)
    if relative.parts and relative.parts[0] in {'tools', 'docs', 'audit-adapter-v1'}:
        return set(names) & {'__pycache__'}
    return set()

shutil.copytree(origin, moved, ignore=ignored)
for path in moved.rglob('*'):
    if path.is_file():
        assert not path.is_symlink()
        files += 1
        size += path.stat().st_size
expected = (moved / 'root-reduction-v1.json').read_bytes()
output = moved / 'root-moved-reduction.json'
origin.rename(hidden)
try:
    command = ['/usr/bin/python3', str(moved / 'tools/root_replay.py'), '--output', str(output)]
    process = subprocess.run(command, cwd=work, capture_output=True, timeout=180,
                             env={'PATH': '/usr/bin:/bin', 'LANG': 'C', 'LC_ALL': 'C', 'TZ': 'UTC'})
    assert process.returncode == 0, process.stderr.decode()
    actual = output.read_bytes()
    assert actual == expected, 'relocated reduction differs'
finally:
    hidden.rename(origin)
receipt = {'schema': 'litchi.xlsx-publication-0944-moved-replay-v1', 'status': 'pass',
           'original_hidden': True, 'original_restored': True, 'byte_identical': True,
           'copied_files': files, 'copied_bytes': size, 'excluded_root_directories': ['source', 'target-native', 'tmp'],
           'argv': command, 'returncode': process.returncode,
           'stdout': process.stdout.decode(), 'stderr': process.stderr.decode(),
           'reduction_bytes': len(actual), 'reduction_sha256': hashlib.sha256(actual).hexdigest(),
           'elapsed_seconds': time.time() - started}
with (origin / 'root-moved-replay.json').open('x') as stream:
    json.dump(receipt, stream, indent=2, sort_keys=True)
    stream.write('\n')
shutil.rmtree(moved)
print(json.dumps(receipt, sort_keys=True))
