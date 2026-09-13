#!/usr/bin/env python3
"""Remove only this campaign's idle, owned target, retaining binary custody."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat

HERE = Path(__file__).resolve().parent
TARGET = Path('/home/zhuhe/litchi-goal-0557-target')
assert TARGET.is_dir() and not TARGET.is_symlink()
assert TARGET.resolve() == TARGET
assert not (HERE / 'cleanup.json').exists()
descriptor = json.loads((HERE / 'baseline/binary-normal.json').read_text())
binary = Path(descriptor['path'])
assert binary.is_relative_to(TARGET) and not binary.is_symlink()
binary_hash = hashlib.sha256(binary.read_bytes()).hexdigest()
assert binary_hash == descriptor['sha256']
assert binary.stat().st_size == descriptor['bytes']
references = []
unreadable = 0
for proc in Path('/proc').iterdir():
    if not proc.name.isdigit():
        continue
    paths = [proc / 'cwd', proc / 'exe']
    try:
        paths += list((proc / 'fd').iterdir())
    except OSError:
        unreadable += 1
    for path in paths:
        try:
            value = os.readlink(path)
            if value == str(TARGET) or value.startswith(str(TARGET) + '/'):
                references.append({'pid': int(proc.name), 'kind': path.name, 'path': value})
        except OSError:
            pass
    try:
        if str(TARGET) in (proc / 'maps').read_text():
            references.append({'pid': int(proc.name), 'kind': 'maps'})
    except OSError:
        unreadable += 1
assert not references, references
files = 0
logical_bytes = 0
for root, _dirs, names in os.walk(TARGET, followlinks=False):
    for name in names:
        info = (Path(root) / name).lstat()
        if stat.S_ISREG(info.st_mode):
            files += 1
            logical_bytes += info.st_size
shutil.rmtree(TARGET)
assert not TARGET.exists() and not TARGET.is_symlink()
record = {
    'utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
    'removed_target': str(TARGET), 'removed': True,
    'regular_file_count': files, 'logical_regular_file_bytes': logical_bytes,
    'binary_sha256_by_kind': {'baseline-normal': binary_hash},
    'binary_bytes_by_kind': {'baseline-normal': descriptor['bytes']},
    'binary_descriptor_sha256': hashlib.sha256(
        (HERE / 'baseline/binary-normal.json').read_bytes()).hexdigest(),
    'accessible_process_references': references,
    'unreadable_or_raced_process_observations': unreadable,
    'reference_scope': 'Accessible /proc cwd, executable, file descriptors and maps; not host-wide visibility.',
    'cleanup_script_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
}
with (HERE / 'cleanup.json').open('x') as stream:
    json.dump(record, stream, indent=2)
    stream.write('\n')
print(json.dumps(record, sort_keys=True))
