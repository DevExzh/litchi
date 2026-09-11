#!/usr/bin/env python3
"""Replay descriptive 0508 evidence checks without requiring removed binaries."""
import copy
import hashlib
import json
from pathlib import Path
import sys

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
sys.path.insert(0, str(REPO))
from tools import generate_corpus_manifest_v2 as generator
from tools import perf_compare
from tools.summarize_crud_baseline import _validate_elapsed
from tools.validate_crud_coverage_index import validate_paths, validate_index, ValidationError
from tools.validate_perf_corpus_binding import validate_binding

def load(path):
    return json.loads(path.read_text())

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    previous = load(HERE / 'inputs/perf-regression-default-manifest-v1.json')
    identity = load(REPO / 'docs/performance/results/perf-regression-default-manifest-v1.json')
    checked = load(REPO / 'docs/performance/results/perf-corpus-manifest-v2.json')
    index = load(REPO / 'docs/performance/crud-coverage-index-v1.json')
    new = set(load(HERE / 'protocol.json')['new_cases'])
    assert identity['default_cases'][:37] == previous['default_cases']
    assert set(identity['default_cases'][37:]) == new
    assert identity['case_count'] == 41 and identity['result_count'] == 213
    assert len(identity['corpora']) == 43
    for name, value in previous['corpora'].items():
        assert identity['corpora'][name] == value
    for case, value in previous['case_corpora'].items():
        assert identity['case_corpora'][case] == value
    old_catalog = load(HERE / 'inputs/perf-corpus-manifest-v2.json')
    old_by_id = {c['id']: c for c in old_catalog['corpora']}
    for corpus in checked['corpora']:
        if corpus['id'] in old_by_id:
            assert corpus == old_by_id[corpus['id']], 'old corpus metadata changed'
    source = load(HERE / 'source-manifest.json')
    for path, digest in source.items():
        assert sha(REPO / path) == digest, path
    for path, digest in load(HERE / 'verifier-sources.json').items():
        assert sha(REPO / path) == digest, path
    for path, digest in load(HERE / 'adr-manifest.json')['files'].items():
        assert sha(REPO / path) == digest, path
    policy = load(REPO / 'docs/performance/perf-regression-policy-v1.json')
    perf_compare.validate_policy(policy)
    assert policy['required_cases'] == identity['default_cases']
    assert policy['expected_result_count'] == 213
    assert policy['expected_result_keys_sha256'] == identity['result_keys_sha256']
    summary = {'kind': 'descriptive baseline; no before/after speedup claim', 'preserved_prior_rows': 201, 'cases': 41, 'corpora': 43, 'rows_per_run': 213, 'full_run_samples': 6390, 'new_export_samples': 360, 'repeats': [], 'negative_probes': []}
    references = {}
    for lane in ['preflight', 'r1', 'r2']:
        report_path = HERE / f'{lane}-report.json'
        catalog_path = HERE / f'{lane}-catalog.json'
        report, catalog = load(report_path), load(catalog_path)
        receipt = load(HERE / f'{lane}-receipt.json')
        assert receipt['exit_code'] == 0 and receipt['source_unchanged']
        assert receipt['binary_sha256'] == load(HERE / 'build-receipt.json')['binary_sha256']
        assert receipt['source_manifest_sha256'] == sha(HERE / 'source-manifest.json')
        assert receipt['log_sha256'] == sha(HERE / f'{lane}.log')
        assert report['binary_identity']['binary_sha256'] == receipt['binary_sha256']
        assert report['binary_identity']['path'] == receipt['command'][0]
        assert report['binary_identity']['executable'] is True
        assert report['tool']['profile'] == 'release' and report['tool']['instrumentation'] == 'none'
        assert report['environment']['git_revision'] == load(HERE / 'protocol.json')['revision']
        assert report['environment']['git_worktree_dirty'] is True
        assert report['environment']['rustc_version'] == load(HERE / 'host.json')['commands']['rustc -Vv']['stdout'].splitlines()[0]
        validate_binding(report, catalog)
        assert catalog == generator.generate(identity, catalog['build']['git_revision'], worktree_dirty=catalog['build']['git_worktree_dirty'])
        assert catalog == checked
        config = report['configuration']
        assert config['cases'] == identity['default_cases']
        samples, warmup = (1, 0) if lane == 'preflight' else (15, 3)
        assert config['samples_per_case'] == samples and config['warmup_iterations_per_case'] == warmup
        assert len(report['results']) == 213
        for key, value in identity['identity_configuration'].items():
            assert config[key] == value
        keys = []
        exports = []
        for row in report['results']:
            corpus = row['corpus']
            assert corpus == identity['corpora'][corpus['name']]
            key = (row['case'], json.dumps(corpus, sort_keys=True, separators=(',', ':')))
            keys.append(key)
            _validate_elapsed(row, samples, f'{lane}:{key[0]}:{corpus["name"]}')
            if row['case'] in new:
                sink = row['sink']
                assert sink['accepted_bytes'] > 0 and sink['write_calls'] > 0
                assert 0 < sink['largest_write'] <= sink['accepted_bytes']
                if row['case'] != 'rtf_semantic_text_to_sink':
                    assert sink['retained_output_bytes'] == 0
                    assert len(row['output_sha256']) == 64
                else:
                    assert 'retained_output_bytes' not in sink
                    assert row.get('output_sha256') is None
                fixed = {'sink': sink, 'output_sha256': row.get('output_sha256')}
                if key in references:
                    assert fixed == references[key]
                references[key] = fixed
                exports.append({'case': row['case'], 'corpus': corpus['name'], 'p50_ns': row['elapsed_ns']['p50'], 'p95_ns': row['elapsed_ns']['p95'], 'p99_ns': row['elapsed_ns']['p99'], **fixed})
        assert len(set(keys)) == 213
        assert perf_compare.result_key_manifest_sha256(keys) == identity['result_keys_sha256']
        if lane != 'preflight':
            validate_paths(REPO / 'docs/performance/crud-coverage-index-v1.json', catalog_path, REPO / 'tools/perf-baseline/src/lib.rs', REPO / 'docs/CRUD_Scenario_Checklist.md', repo_root=REPO, report_path=report_path)
            summary['repeats'].append({'lane': lane, 'elapsed_seconds': receipt['elapsed_seconds'], 'exports': exports})
    # A newly promoted binding must have real timing vectors and a row in the report.
    report = load(HERE / 'r1-report.json')
    for kind in ['missing-export-row', 'short-export-sample-vector']:
        bad = copy.deepcopy(report)
        position = next(i for i, r in enumerate(bad['results']) if r['case'] in new)
        if kind == 'missing-export-row':
            bad['results'].pop(position)
        else:
            bad['results'][position]['elapsed_ns']['samples'].pop()
        try:
            validate_index(index, checked, (REPO / 'tools/perf-baseline/src/lib.rs').read_text(), (REPO / 'docs/CRUD_Scenario_Checklist.md').read_text(), repo_root=REPO, checked_catalog=checked, report=bad)
        except ValidationError as error:
            summary['negative_probes'].append({'kind': kind, 'rejected': True, 'reason': str(error)})
        else:
            raise AssertionError(f'accepted {kind}')
    gates = load(HERE / 'gates.json')
    assert gates['rust_tests_passed'] == 483 and gates['rust_tests_ignored'] == 1
    assert gates['python_tests_passed'] == 199
    for gate in gates['gates']:
        assert gate['exit_code'] == 0
        log = gate['lane'] if gate['lane'].endswith('.log') else gate['lane'] + '.log'
        assert sha(HERE / log) == gate['log_sha256'], log
    for lane in ['build', 'tests', 'clippy', 'rustdoc', 'doctests', 'check-features']:
        receipt = load(HERE / f'{lane}-receipt.json')
        assert receipt['exit_code'] == 0 and receipt['source_unchanged']
        assert receipt['source_manifest_sha256'] == sha(HERE / 'source-manifest.json')
        assert receipt['log_sha256'] == sha(HERE / f'{lane}.log')
    summary['rust_tests_passed'] = 483
    summary['rust_tests_ignored'] = 1
    summary['python_tests_passed'] = 199
    print(json.dumps(summary, indent=2))

if __name__ == '__main__':
    main()
