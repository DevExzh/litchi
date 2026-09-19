#!/usr/bin/env python3
"""Check retained baseline/source/probe/corpus/raw-result bindings."""
import hashlib
import json
import subprocess
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
def read(name):
    return json.loads((P / name).read_text())
base = read('baseline.json')
for section in ['constraints_sha256','source_sha256']:
    for name, digest in base[section].items():
        assert sha(ROOT / name) == digest, name
for name, digest in base['build_inputs_sha256'].items():
    if name != 'Cargo.lock':
        assert sha(ROOT / name) == digest, name
builds = read('builds.json')
assert {b['label'] for b in builds} == {'native','allocations'}
for build in builds:
    assert build['exit_code'] == 0
    assert build['source_sha256'] == base['source_sha256']
    for name, digest in build['probe_sha256'].items():
        assert sha(P / name) == digest, name
native = next(b for b in builds if b['label']=='native')
runs = read('native-runs.json')
assert len(runs) == 20
assert len({(r['case'],r['leg']) for r in runs}) == 20
for run in runs:
    assert run['exit_code'] == 0
    assert run['binary_sha256'] == native['binary_sha256']
    assert sha(P / run['output']) == run['output_sha256']
    assert sha((P / run['output']).with_suffix('.stderr')) == run['stderr_sha256']
for case in read('corpus.json').values():
    assert sha(ROOT / case['path']) == case['sha256']
control = read('control-manifest.json')
assert sha(ROOT / control['source']) == control['source_sha256']
assert sum(row['replacements'] for row in control['members']) == 43
assert len(control['members']) == 103
for name in ['native-summary.json','semantic-bindings.json']:
    prior = (P / name).read_bytes()
    result = subprocess.run(['python3',str(P / 'summarize.py')],cwd=ROOT,capture_output=True)
    assert result.returncode == 0, result.stderr.decode()
    assert (P / name).read_bytes() == prior, name
profile = read('profile/binding.json')
assert profile['binary_sha256'] == native['binary_sha256']
for name, digest in profile['files'].items():
    assert sha(P / 'profile' / name) == digest, name
print('PASS: unchanged production and constraints; frozen builds, corpus, raw samples, semantic repeatability, summaries and native profile.')
