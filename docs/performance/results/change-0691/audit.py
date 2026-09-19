#!/usr/bin/env python3
"""Check retained baseline/source/probe/corpus/raw-result bindings."""
import hashlib
import json
import subprocess
import sys
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
if '--initial' in sys.argv:
    P = P / 'initial'
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
allocation_build = next(b for b in builds if b['label']=='allocations')
allocation_runs = read('allocation-runs.json')
assert len(allocation_runs)==5
for run in allocation_runs:
    assert run['exit_code']==0
    assert run['binary_sha256']==allocation_build['binary_sha256']
    assert sha(P / run['output'])==run['output_sha256']
for case in read('corpus.json').values():
    assert sha(ROOT / case['path']) == case['sha256']
control = read('control-manifest.json')
assert sha(ROOT / control['source']) == control['source_sha256']
assert sum(row['replacements'] for row in control['members']) == 43
assert len(control['members']) == 103
prior = {name:(P / name).read_bytes() for name in ['native-summary.json','semantic-bindings.json']}
result = subprocess.run(['python3',str(P / 'summarize.py')],cwd=ROOT,capture_output=True)
assert result.returncode == 0, result.stderr.decode()
for name, contents in prior.items():
    assert (P / name).read_bytes() == contents, name
profile = read('profile/binding.json')
assert profile['binary_sha256'] == native['binary_sha256']
for name, digest in profile['files'].items():
    assert sha(P / 'profile' / name) == digest, name
quality = read('quality.json')
assert quality['passed']==587 and quality['failed']==quality['ignored']==0
assert quality['pptx_lib_exit_code']==quality['format_exit_code']==0
assert quality['source_sha256']==base['source_sha256']
for name,digest in quality['logs_sha256'].items():
    assert sha(P / name)==digest,name
packet = ROOT / 'docs/performance/results/change-0691'
trace_output = packet / ('trace-audit-initial.json' if '--initial' in sys.argv else 'trace-audit.json')
result = subprocess.run(['python3',str(packet/'trace-summary.py'),
                         *map(str, sorted((P/'trace-runs').iterdir())), '--output',str(trace_output)],
                        cwd=ROOT,capture_output=True)
assert result.returncode==0,result.stderr.decode()
assert json.loads(trace_output.read_text())['status']=='pass'
print('PASS: unchanged production and constraints; frozen builds, corpus, raw samples, semantic repeatability, summaries and native profile.')
