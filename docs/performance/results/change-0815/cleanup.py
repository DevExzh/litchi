"""Root-only cleanup after completed analysis and source disposition."""
import shutil
import time

import custody as c

assert not (c.P / 'cleanup.json').exists()
assert c.TARGET == c.ROOT.parent / 'litchi-target-0815'
assert c.TARGET.is_dir() and not c.TARGET.is_symlink()
for name in ('analysis.json', 'profile-analysis.json', 'root-audit.json',
             'quality-summary.json', 'decision.json', 'disposition.json'):
    assert (c.P / name).is_file(), name
disposition = c.read(c.P / 'disposition.json')
leg = 'after' if disposition['production_change_retained'] else 'before'
expected = c.read(c.P / f'build-{leg}/source.json')['files']
assert c.source()['files'] == expected
removed = []
for build_leg in ('before', 'after'):
    for name, artifact in c.read(c.P / f'build-{build_leg}/build.json')['binaries'].items():
        assert c.artifact(artifact['path']) == artifact
        removed.append(artifact)
assert len(removed) == 6
paths = [p for p in c.TARGET.rglob('*') if p.is_file()]
size = sum(p.stat().st_size for p in paths)
started = time.time()
shutil.rmtree(c.TARGET)
assert not c.TARGET.exists()
assert c.source()['files'] == expected
c.write(c.P / 'cleanup.json', {
    'schema': 'litchi.performance.0815.cleanup.v1',
    'target': str(c.TARGET), 'target_removed': True,
    'removed_files': len(paths), 'removed_logical_bytes': size,
    'removed_binaries': removed, 'binaries_verified_before_removal': True,
    'source_leg': leg, 'source_manifest': c.artifact(c.P / f'build-{leg}/source.json'),
    'started': started, 'ended': time.time(),
})
print('0815 owned target removed; six binary identities and final source verified')
