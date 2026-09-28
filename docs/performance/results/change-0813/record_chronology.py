"""Retain observed artifact timestamps so replay does not depend on checkout mtimes.

These are filesystem observations recorded after the decision, not timestamps
emitted by the earlier drivers. Timestamped workload receipts remain primary
for stage ordering. No artifact time is changed by this script.
"""
import time

import custody as c

assert not (c.P / 'chronology.json').exists()
names = ('application', 'qualification-audit', 'analysis', 'root-audit',
         'profile-analysis', 'decision', 'disposition')
files = {}
for name in names:
    path = c.P / (name + '.json')
    files[name] = {**c.artifact(path), 'observed_mtime_ns': path.stat().st_mtime_ns}
c.write(c.P / 'chronology.json', {
    'schema': 'litchi.performance.0813.chronology.v1',
    'captured_at': time.time(),
    'evidence_kind': 'post-decision observation of unchanged artifact filesystem mtimes',
    'files': files,
})
print('0813 chronology observations retained without changing artifact timestamps')
