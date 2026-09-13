"""Restore the measured baseline after rejected-candidate quality terminates."""
import json
from pathlib import Path
import subprocess
import sys

import run as R


subprocess.run([
    sys.executable, '-B', str(R.HERE / 'summarize_quality.py'),
    '--stage', 'candidate', '--output', str(R.HERE / 'candidate-quality-summary.json'),
], cwd=R.REPO, check=True)
production = R.REPO / 'crates/litchi-cfb/src/file.rs'
assert R.sha(production) == R.sha(R.HERE / 'candidate-sources/file.rs')
baseline = json.loads((R.HERE / 'baseline/source-manifest.json').read_text())
restored = subprocess.check_output([
    'git', 'show', json.loads((R.HERE / 'plan.json').read_text())['revision']
    + ':crates/litchi-cfb/src/file.rs',
], cwd=R.REPO)
production.write_bytes(restored)
assert R.sha(production) == baseline['crates/litchi-cfb/src/file.rs']
R.configure('final')
R.freeze()
assert (R.HERE / 'final/source-manifest.json').read_bytes() == (R.HERE / 'baseline/source-manifest.json').read_bytes()
assert not (R.HERE / 'final/source.patch').read_bytes()
identities = {}
for kind in ['normal', 'alloc']:
    descriptor = R.HERE / 'baseline' / ('binary-' + kind + '.json')
    identity = json.loads(descriptor.read_text())
    assert R.sha(Path(identity['path'])) == identity['sha256']
    identities[kind] = {'descriptor_sha256': R.sha(descriptor), 'binary_sha256': identity['sha256']}
R.write(R.HERE / 'restoration.json', {
    'observed_utc': R.now(), 'disposition': 'rejected',
    'final_source_manifest_sha256': R.sha(R.HERE / 'final/source-manifest.json'),
    'baseline_source_manifest_sha256': R.sha(R.HERE / 'baseline/source-manifest.json'),
    'candidate_quality_summary_sha256': R.sha(R.HERE / 'candidate-quality-summary.json'),
    'retained_baseline_binaries': identities,
    'scope': 'Exact final runtime-plus-test source equals the measured baseline, including the common guard. No changed final rebuild or new performance claim. Fresh final quality follows.',
})
print('Exact measured baseline restored; running final quality.', flush=True)
subprocess.run([sys.executable, '-B', str(R.HERE / 'checks.py'), '--stage', 'final'], cwd=R.REPO, check=True)
subprocess.run([sys.executable, '-B', str(R.HERE / 'summarize_quality.py'), '--stage', 'final'], cwd=R.REPO, check=True)
