"""Root-only constructor repair application after its protected preflight."""
import subprocess
from pathlib import Path
import custody as c

p = c.P
assert not (p / 'quality-amendment-application.json').exists()
assert not (p / 'build-after').exists()
assert (p / 'source-amendment-review.md').is_file()
original = c.read(p / 'application.json')
assert c.source() == original['source']
decision_path = p / 'amendment-preflight/decision.json'
decision = c.read(decision_path)
assert decision['advance_to_workflow_trials'] is True
assert decision['production_adoption'] is False
assert decision['protected_consume_regressions'] == []
assert decision['dominant_class_benefits'] == {'distinct-1': True, 'distinct-2': True}
audit_path = p / 'amendment-preflight/root-native-audit.json'
audit_descriptor = decision['independent_audit']
recorded_path = Path(audit_descriptor['path'])
if not recorded_path.is_absolute():
    recorded_path = p / 'amendment-preflight' / recorded_path
assert recorded_path.resolve() == audit_path.resolve()
assert audit_descriptor['bytes'] == audit_path.stat().st_size
assert audit_descriptor['sha256'] == c.sha(audit_path)
audit = c.read(audit_path)
assert audit['passed'] is True and audit['matches_primary_analysis'] is True
assert audit['native_reports'] == 936 and audit['native_samples'] == 28080
assert audit['advance_to_workflow_trials'] is True and audit['production_adoption'] is False
manifest_path = p / 'candidate-quality-amendment/manifest.json'
manifest = c.read(manifest_path)
assert manifest['schema'] == 'litchi.performance.0806.quality-amendment.v1'
assert manifest['parent_candidate']['application'] == c.artifact(p / 'application.json')
files = manifest['files']
allowed = set(c.read(p / 'plan.json')['source_allowlist'])
shared_tests = 'crates/litchi-opc/src/xml_attributes/tests.rs'
assert {row['production_path'] for row in files.values()} == allowed - {shared_tests}
expected = dict(original['source']['files'])
for row in files.values():
    assert c.artifact(row['before']['path']) == row['before']
    assert c.artifact(row['after']['path']) == row['after']
    assert c.sha(c.ROOT / row['production_path']) == row['before']['sha256']
    expected[row['production_path']] = row['after']['sha256']
patch = p / 'candidate-quality-amendment/candidate-quality-amendment.patch'
subprocess.run(['git', 'apply', '--check', str(patch)], cwd=c.ROOT, check=True)
subprocess.run(['git', 'apply', str(patch)], cwd=c.ROOT, check=True)
source = c.source()
assert source['revision'] == original['source']['revision'] and source['files'] == expected
assert c.sha(c.ROOT / shared_tests) == original['source']['files'][shared_tests]
c.write(p / 'quality-amendment-application.json', {
    'schema': 'litchi.performance.0806.quality-amendment-application.v1',
    'original_application': c.artifact(p / 'application.json'),
    'manifest': c.artifact(manifest_path),
    'patch': c.artifact(patch),
    'preflight': c.artifact(decision_path),
    'source': source,
})
print('Preflight-qualified constructor repair applied; production adoption remains pending')
