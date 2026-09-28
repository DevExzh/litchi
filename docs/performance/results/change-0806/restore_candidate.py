"""Restore the six baseline files after the frozen benefit gate rejects 0806."""
from pathlib import Path
import custody as c

p = c.P
assert not (p / 'disposition.json').exists()
assert not (p / 'restored-source.json').exists()
before = c.read(p / 'build-before/source.json')
after = c.read(p / 'build-after/source.json')
assert c.source() == after
analysis = c.read(p / 'analysis.json')
guard = analysis['decision_guards']
assert guard['adoption_eligible'] is False
assert guard['benefit_satisfied'] is False and guard['eligible_benefits'] == []
audit = c.read(p / 'root-audit.json')
assert audit['reports'] == 306 and audit['samples'] == 6714
assert audit['benefits'] == [] and audit['production_adoption'] is False
assert c.read(p / 'profile-analysis.json')['qualified'] is True
files = c.read(p / 'candidate/manifest.json')['files']
allowed = set(c.read(p / 'plan.json')['source_allowlist'])
assert {row['production_path'] for row in files.values()} == allowed
assert {name for name in before['files'] if before['files'][name] != after['files'][name]} == allowed
for row in files.values():
    baseline = Path(row['before']['path'])
    assert c.artifact(baseline) == row['before']
    assert row['before']['sha256'] == before['files'][row['production_path']]
for row in files.values():
    (c.ROOT / row['production_path']).write_bytes(Path(row['before']['path']).read_bytes())
restored = c.source()
assert restored == before
c.write(p / 'restored-source.json', restored)
c.write(p / 'disposition.json', {
    'status': 'rejected',
    'production_change_retained': False,
    'reason': 'No capture or lifecycle row satisfies the frozen 3% latency benefit gate; allocation reductions alone are insufficient.',
    'restored_source': c.artifact(p / 'restored-source.json'),
})
print('Rejected for missing required workflow latency benefit; all six baseline files restored exactly')
