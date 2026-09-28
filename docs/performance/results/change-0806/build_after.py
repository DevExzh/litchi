"""Root-only after-build handoff, binding qualified fixtures before compilation."""
from datetime import datetime, timezone
import subprocess
import sys
import custody as c

p = c.P
assert not (p / 'build-after').exists()
assert not (p / 'after-build-inputs.json').exists()
assert (p / 'qualification-review.md').is_file()
contract = c.read(p / 'qualification-four-attr.json')
assert contract['schema'] == 'litchi.performance.0806.qualification-four-attr.v1'
assert {row['mode'] for row in contract['oracles']} == {'capture', 'commit', 'lifecycle'}
application_path = p / 'visibility-amendment-application.json'
application = c.read(application_path)
assert application['schema'] == 'litchi.performance.0806.visibility-amendment-application.v1'
assert application['original_application'] == c.artifact(p / 'quality-amendment-application.json')
constructor_application = c.read(p / 'quality-amendment-application.json')
assert constructor_application['original_application'] == c.artifact(p / 'application.json')
preflight = c.read(p / 'amendment-preflight/decision.json')
assert constructor_application['preflight'] == c.artifact(p / 'amendment-preflight/decision.json')
assert preflight['advance_to_workflow_trials'] is True
assert preflight['production_adoption'] is False
assert preflight['protected_consume_regressions'] == []
assert preflight['dominant_class_benefits'] == {'distinct-1': True, 'distinct-2': True}
source = c.source()
assert source == application['source']
quality = c.read(p / 'quality.json')
assert c.read(quality['source']['path']) == source
assert len(quality['rows']) == 6 and all(row['exit_code'] == 0 for row in quality['rows'])
frozen_at = datetime.now(timezone.utc)
qualification_ended = max(row['ended'] for row in c.read(p / 'qualification/receipts.json'))
assert frozen_at.timestamp() > qualification_ended
assert frozen_at > datetime.fromisoformat(contract['freeze']['frozen_at_utc'].replace('Z', '+00:00'))
c.write(p / 'after-build-inputs.json', {
    'schema': 'litchi.performance.0806.after-build-inputs.v1',
    'qualification': c.artifact(p / 'qualification-four-attr.json'),
    'application': c.artifact(application_path),
    'source': source,
    'frozen_at_utc': frozen_at.isoformat().replace('+00:00', 'Z'),
})
subprocess.run([sys.executable, '-B', str(p / 'build.py'), 'after'],
               cwd=c.ROOT, check=True)
