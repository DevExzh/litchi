"""Root-only removal of the owned target and scratch after terminal captures."""
import shutil
import time

import custody as c

assert not (c.P / 'cleanup.json').exists()
assert c.TARGET == c.ROOT.parent / 'litchi-target-0821'
assert c.SCRATCH == c.ROOT.parent / 'litchi-fs-0821'
assert (c.P / 'results-review.md').is_file()
build = c.read(c.P / 'build.json')
frozen = c.read(build['source']['path'])
c.unchanged(frozen)
c.assert_unrelated()
removed_binaries = [x['artifact'] for x in build['binaries'].values()]
assert len(removed_binaries) == 3
for descriptor in removed_binaries:
    assert c.artifact(descriptor['path']) == descriptor
started = time.time()
removed = []
for path in (c.TARGET, c.SCRATCH):
    assert path.is_dir() and not path.is_symlink()
    if path == c.SCRATCH:
        marker = path / '.litchi-performance-0821-owned'
        assert marker.read_text() == 'litchi-performance-0821-owned-scratch-v1\n'
    files = [x for x in path.rglob('*') if x.is_file()]
    removed.append({'path': str(path), 'files': len(files), 'logical_bytes': sum(x.stat().st_size for x in files)})
    shutil.rmtree(path)
    assert not path.exists()
c.unchanged(frozen)
c.assert_unrelated()
c.write(c.P / 'cleanup.json', {'schema': 'litchi.performance.0821.cleanup.v1', 'target_removed': True, 'scratch_removed': True, 'binaries_verified_before_removal': True, 'removed_binaries': removed_binaries, 'source': build['source'], 'removed': removed, 'started': started, 'ended': time.time()})
print('0821 owned build/scratch removed; three binaries and source custody verified')
