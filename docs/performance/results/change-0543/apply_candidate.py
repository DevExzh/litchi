"""Apply the reviewed candidate only after the complete baseline phase exits."""
import json
import subprocess
import run

run.check_source('baseline')
# The baseline's last standalone capture proves that the full phase completed.
last = run.HERE / 'baseline/guard-alloc-r1-dense-sparse-late-raw.receipt.json'
assert json.loads(last.read_text())['exit_code'] == 0
patch = run.HERE / 'candidate.patch'
subprocess.run(['git', 'apply', '--check', str(patch)], cwd=run.REPO, check=True)
subprocess.run(['git', 'apply', str(patch)], cwd=run.REPO, check=True)
files = run.plan_data()['candidate_files']
subprocess.run(['rustfmt', '--edition', '2024', '--check', *files], cwd=run.REPO, check=True)
(run.HERE / 'applied-candidate.patch').write_bytes(patch.read_bytes())
run.write(run.HERE / 'candidate-frozen-inputs.json', {
    'created_utc': run.now(),
    'stage': 'before candidate freeze/build/capture',
    'files': {name: run.sha(run.HERE / name) for name in [
        'candidate.patch', 'applied-candidate.patch', 'candidate-design.md',
        'source-review.md', 'test-review.md', 'apply_candidate.py',
        'differential-tests-final.patch', 'baseline-frozen-inputs.json',
    ]},
    'candidate_sources': {name: run.sha(run.REPO / name) for name in files},
    'baseline_manifest_sha256': run.sha(run.HERE / 'baseline/source-manifest.json'),
})
print('Reviewed candidate applied and frozen.', flush=True)
