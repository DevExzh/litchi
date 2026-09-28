"""Root-only removal of the owned target after all captures and decoding."""
import shutil
import time

import custody as c

assert not (c.P / 'cleanup.json').exists()
assert c.TARGET == c.ROOT.parent / 'litchi-target-0822'
assert c.SCRATCH is None
assert (c.P / 'results-review.md').is_file()
assert (c.P / 'analysis.json').is_file()
assert (c.P / 'perf/decode-complete.json').is_file()
build = c.read(c.P / 'build.json')
frozen = c.read(build['frozen_inputs']['path'])
c.stable(frozen)
removed_binaries = [x['artifact'] for x in build['binaries'].values()]
assert set(build['binaries']) == {'ordinary', 'fp'}
for descriptor in removed_binaries:
    assert c.artifact(descriptor['path']) == descriptor
started = time.time()
assert c.TARGET.is_dir() and not c.TARGET.is_symlink()
files = [x for x in c.TARGET.rglob('*') if x.is_file()]
removed = [{'path': str(c.TARGET), 'files': len(files),
            'logical_bytes': sum(x.stat().st_size for x in files)}]
shutil.rmtree(c.TARGET)
assert not c.TARGET.exists()
c.stable(frozen)
c.write(c.P / 'cleanup.json', {
    'schema': 'litchi.performance.0822.cleanup.v1',
    'target': str(c.TARGET),
    'target_removed': True, 'scratch': None,
    'binaries_verified_before_removal': True,
    'removed_binaries': removed_binaries,
    'source': build['source'], 'frozen_inputs': build['frozen_inputs'],
    'removed': removed, 'started': started, 'ended': time.time(),
})
print('0822 owned target removed; two binaries and source custody verified')
