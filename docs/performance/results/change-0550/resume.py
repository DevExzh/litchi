"""Resume after pre-measurement gdbserver setup failure; retain prior captures."""
import run as R
from capture_amended import capture
capture('profile', 1)
capture('profile', 2)
capture('native', 2)
R.build('alloc')
capture('alloc', 1)
capture('alloc', 2)
for name, command in [
    ('fmt', ['cargo','fmt','--all','--check']),
    ('harness-fmt', ['cargo','fmt','--manifest-path','tools/perf-baseline/Cargo.toml','--all','--check']),
    ('boundaries', ['python3','-B','tools/check_crate_boundaries.py']),
    ('claims', ['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--evidence-root','.','--mode','strict'])]:
    R.run('check-' + name, command)
