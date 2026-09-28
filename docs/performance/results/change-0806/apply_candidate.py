"""Root-only temporary candidate application after baseline qualification."""
import subprocess
import custody as c

p = c.P
assert not (p / 'application.json').exists()
assert not (p / 'build-after').exists()
assert (p / 'qualification-review.md').is_file()
oracle = c.read(p / 'qualification-four-attr.json')
assert oracle['schema'] == 'litchi.performance.0806.qualification-four-attr.v1'
assert len(oracle['oracles']) == 3
before = c.read(p / 'source.json')
assert c.source() == before
manifest = c.read(p / 'candidate/manifest.json')
plan = c.read(p / 'plan.json')
files = manifest['files']
assert sorted(row['production_path'] for row in files.values()) == sorted(plan['source_allowlist'])
expected = dict(before['files'])
for row in files.values():
    assert c.artifact(row['before']['path']) == row['before']
    assert c.artifact(row['after']['path']) == row['after']
    assert c.sha(c.ROOT / row['production_path']) == row['before']['sha256']
    expected[row['production_path']] = row['after']['sha256']
patch = p / 'candidate/candidate.patch'
subprocess.run(['git', 'apply', '--check', str(patch)], cwd=c.ROOT, check=True)
subprocess.run(['git', 'apply', str(patch)], cwd=c.ROOT, check=True)
source = c.source()
assert source['revision'] == before['revision'] and source['files'] == expected
c.write(p / 'application.json', {
    'manifest': c.artifact(p / 'candidate/manifest.json'),
    'patch': c.artifact(patch),
    'source': source,
})
print('Exact six-file candidate applied temporarily; adoption remains pending')
