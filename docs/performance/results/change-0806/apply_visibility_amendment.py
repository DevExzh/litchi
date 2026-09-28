"""Restore the OLE helper's baseline public API before full workflow gates."""
import subprocess
from pathlib import Path
import custody as c

p = c.P
assert not (p / 'visibility-amendment-application.json').exists()
assert not (p / 'build-after').exists()
assert (p / 'visibility-amendment-review.md').is_file()
parent_path = p / 'quality-amendment-application.json'
parent = c.read(parent_path)
assert c.source() == parent['source']
manifest_path = p / 'candidate-visibility-amendment/manifest.json'
manifest = c.read(manifest_path)
assert manifest['schema'] == 'litchi.performance.0806.visibility-amendment.v1'
assert manifest['parent_application'] == c.artifact(parent_path)
name = 'litchi-ole-common-xml_attributes.rs'
production_path = 'crates/litchi-ole-common/src/xml_attributes.rs'
assert set(manifest['files']) == {name}
row = manifest['files'][name]
assert row['production_path'] == production_path
for leg in ('before', 'after'):
    assert c.artifact(row[leg]['path']) == row[leg]
before = Path(row['before']['path']).read_text()
after = Path(row['after']['path']).read_text()
assert before.count('pub(crate) trait BytesStartExt') == 1
assert before.count('pub(crate) struct CheckedAttributes') == 1
assert after == before.replace('pub(crate) trait BytesStartExt', 'pub trait BytesStartExt', 1).replace(
    'pub(crate) struct CheckedAttributes', 'pub struct CheckedAttributes', 1)
assert c.sha(c.ROOT / production_path) == row['before']['sha256']
patch = p / 'candidate-visibility-amendment/candidate-visibility-amendment.patch'
assert manifest['patch'] == c.artifact(patch)
subprocess.run(['git', 'apply', '--check', str(patch)], cwd=c.ROOT, check=True)
subprocess.run(['git', 'apply', str(patch)], cwd=c.ROOT, check=True)
expected = dict(parent['source']['files'])
expected[production_path] = row['after']['sha256']
source = c.source()
assert source['revision'] == parent['source']['revision'] and source['files'] == expected
c.write(p / 'visibility-amendment-application.json', {
    'schema': 'litchi.performance.0806.visibility-amendment-application.v1',
    'original_application': c.artifact(parent_path),
    'manifest': c.artifact(manifest_path),
    'patch': c.artifact(patch),
    'source': source,
})
print('Baseline OLE public visibility restored; production adoption remains pending')
