#!/usr/bin/env python3
"""Audit unchanged source, reused checks, fresh reports, cleanup and replay."""
import importlib.util
import json
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('capture0714audit', P/'capture.py')
C = importlib.util.module_from_spec(spec)
spec.loader.exec_module(C)

def main():
    source = C.read(P/'source.json')
    assert C.C.census() == source == C.read(P.parent/'change-0713/source-final.json')
    assert C.read(P/'revision.json')['previous_goal_turn'].startswith('progress:')
    for name, sha in C.read(P/'helper-freeze.json').items():
        assert C.C.sha(C.C.ROOT/name) == sha
    for name, sha in C.read(P/'analysis-freeze.json').items():
        assert C.C.sha(P/name) == sha
    reused = C.read(P/'reused-verification.json')
    prior = C.C.ROOT/reused['packet']
    for name, sha in reused['files'].items():
        assert C.C.sha(prior/name) == sha
    assert C.read(prior/'quality-source.json') == source
    for row in C.read(prior/'quality.json'):
        assert row['exit_code'] == 0 and row['source_manifest_sha256'] == C.C.sha(prior/'quality-source.json')
        assert row['log_sha256'] == C.C.sha(prior/row['log'])
    evidence = C.read(prior/'evidence/results.json')
    assert len(evidence) == 6
    for row in evidence:
        assert row['exit_code'] == 0 and row['source_manifest_sha256'] == C.C.sha(prior/'source-final.json')
        assert row['log_sha256'] == C.C.sha(prior/'evidence'/(row['name']+'.log'))
    cleanup = C.read(P/'cleanup.json')
    assert cleanup['owned_paths'] == ['/home/zhuhe/code/litchi-target-0714', '/home/zhuhe/code/litchi-0714-bin', '/home/zhuhe/code/litchi-0714-fs']
    assert cleanup['owned_paths_absent'] and all(not Path(n).exists() for n in cleanup['owned_paths'])
    assert cleanup['binaries'] == [C.read(P/'build.json')['binary']]
    final = C.read(P/'final-report-gate.json')
    assert final['exit_code'] == 0 and final['log_sha256'] == C.C.sha(P/'final-report-gate.log')
    for name, sha in final['docs'].items():
        assert C.C.sha(C.C.ROOT/name) == sha
    negative = C.read(P/'negative-checks.json')
    assert negative['status'] == 'pass' and negative['retained_inputs_unchanged']
    assert negative['analysis_sha256'] == C.C.sha(P/'analysis.json')
    assert negative['analyzer_sha256'] == C.C.sha(P/'analyze.py')
    assert negative['trace_parser_sha256'] == C.C.sha(P/'trace_analysis.py')
    assert negative['exact_positive_replay'] and negative['temporary_trace_removed']
    assert all(row['rejected'] for row in negative['checks']) and len(negative['checks']) >= 4
    subprocess.run(['python3', '-B', str(P/'analyze.py'), '--check'], cwd=C.C.ROOT, check=True)
    print('PASS: unchanged source, bound reused checks, fresh phase/trace evidence, docs and cleanup')

if __name__ == '__main__':
    main()
