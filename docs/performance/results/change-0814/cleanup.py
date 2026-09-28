"""Remove only the completed diagnostic's owned target after evidence review."""
import shutil
import time

import custody as c

assert not (c.P / 'cleanup.json').exists()
assert c.TARGET == c.ROOT.parent / 'litchi-target-0814'
assert c.TARGET.is_dir() and not c.TARGET.is_symlink()
for name in ('analysis.json', 'root-audit.json'):
    assert (c.P / name).is_file(), name
source_path = c.P / 'build/source.json'
source = c.read(source_path)
assert c.source()['files'] == source['files']
build = c.read(c.P / 'build/build.json')
assert set(build['binaries']) == {'control', 'profile', 'fp'}
removed = []
for binary in build['binaries'].values():
    assert c.artifact(binary['path']) == binary
    removed.append(binary)
paths = [p for p in c.TARGET.rglob('*') if p.is_file()]
size = sum(p.stat().st_size for p in paths)
started = time.time()
shutil.rmtree(c.TARGET)
assert not c.TARGET.exists()
assert c.source()['files'] == source['files']
c.write(c.P / 'cleanup.json', {
    'schema': 'litchi.performance.0814.cleanup.v1',
    'target': str(c.TARGET), 'target_removed': True,
    'removed_files': len(paths), 'removed_logical_bytes': size,
    'removed_binaries': removed, 'binaries_verified_before_removal': True,
    'source_manifest': c.artifact(source_path),
    'started': started, 'ended': time.time(),
})
print('0814 owned target removed after exact three-binary and source checks')
