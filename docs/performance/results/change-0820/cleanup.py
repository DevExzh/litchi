"""Root-only cleanup after the allocator test repair; no release binaries ran."""
import shutil
import time

import custody as c

P = c.P
assert not (P / 'cleanup.json').exists()
assert (P / 'results-review.md').is_file()
assert c.read(P / 'repair/quality.json')['status'] == 'pass'
source_path = P / 'repair/source.json'
frozen = c.read(source_path)
c.unchanged(frozen)
c.assert_unrelated()
assert c.TARGET == c.ROOT.parent / 'litchi-target-0820'
assert c.SCRATCH == c.ROOT.parent / 'litchi-fs-0820'
assert not (P / 'build.json').exists()
assert not c.SCRATCH.exists(), 'no workload was authorized to create scratch'
started = time.time()
removed = []
assert c.TARGET.is_dir() and not c.TARGET.is_symlink()
assert (c.TARGET / 'quality-1').is_dir()
files = [path for path in c.TARGET.rglob('*') if path.is_file()]
removed.append({'path': str(c.TARGET), 'files': len(files),
                'logical_bytes': sum(path.stat().st_size for path in files)})
shutil.rmtree(c.TARGET)
assert not c.TARGET.exists() and not c.SCRATCH.exists()
c.unchanged(frozen)
c.assert_unrelated()
c.write(P / 'cleanup.json', {'schema': 'litchi.performance.0820.cleanup.v1',
                             'target_removed': True, 'scratch_removed': True,
                             'release_binaries_built': 0, 'source': c.artifact(source_path),
                             'removed': removed, 'started': started, 'ended': time.time()})
print('0820 repair target removed; scratch was never created')
