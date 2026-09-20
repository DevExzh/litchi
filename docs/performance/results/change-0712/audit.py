#!/usr/bin/env python3
"""Verify current-source diagnostic custody, historical attribution and cleanup."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def read(path):
    return json.loads(path.read_text())

def main():
    spec = importlib.util.spec_from_file_location('custody0712', P/'custody.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    source = read(P/'source-final.json')
    assert custody.census() == source == read(P.parent/'change-0711/source-baseline.json')
    for name, digest in read(P/'constraints.json').items():
        assert sha(ROOT/name) == digest
    freeze = read(P/'oracle-freeze.json')
    for name, digest in freeze.items():
        assert sha(P/'oracle'/name) == digest
    folder = P/'oracle/current'
    result, build = read(folder/'result.json'), read(folder/'build.json')
    assert result['exit_code'] == build['exit_code'] == 0
    assert read(folder/'source.json') == source
    assert read(folder/'probe.json') == freeze
    assert result['source_sha256'] == build['source_sha256'] == sha(folder/'source.json') == sha(P/'source-final.json')
    assert result['probe_sha256'] == build['probe_sha256'] == sha(folder/'probe.json')
    assert sha(folder/'build.log') == build['log_sha256']
    for name, digest in result['artifacts'].items():
        assert sha(folder/name) == digest
    assert '--locked' in build['command']
    replay = read(P/'oracle-replay.json')
    assert replay['exit_code'] == 0 and replay['byte_identical'] and replay['temporary_output_removed']
    assert replay['binary'] == result['binary'] and replay['report_sha256'] == sha(folder/'report.json')
    fmt = read(P/'probe-format.json')
    assert fmt['exit_code'] == 0 and fmt['source_sha256'] == sha(P/'oracle/src/main.rs')
    evidence = read(P/'evidence/results.json')
    assert len(evidence) == 6
    for row in evidence:
        assert row['exit_code'] == 0 and row['source_manifest_sha256'] == sha(P/'source-final.json')
        assert row['log_sha256'] == sha(P/'evidence'/(row['name']+'.log'))
    final = read(P/'final-report-gate.json')
    assert final['exit_code'] == 0 and final['log_sha256'] == sha(P/'final-report-gate.log')
    for name, digest in final['docs'].items():
        assert sha(ROOT/name) == digest
    cleanup = read(P/'cleanup.json')
    assert cleanup['owned_paths'] == ['/home/zhuhe/code/litchi-target-0712', '/home/zhuhe/code/litchi-0712-bin']
    assert cleanup['owned_paths_absent'] and all(not Path(n).exists() for n in cleanup['owned_paths'])
    assert cleanup['binaries'] == [result['binary']]
    for script, args in [('analyze.py', []), ('mce-attribution.py', ['--replay'])]:
        subprocess.run(['python3', '-B', str(P/script), *args], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
    print('PASS: unchanged production, 11 current-source cases, exact replay, four historical profiles, six gates, final docs and cleanup')

if __name__ == '__main__':
    main()
