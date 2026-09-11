#!/usr/bin/env python3
"""Derive a provisional checked contract from the one-sample preflight.

The pass result concerns identity derivation only. Do not publish the coverage
promotion until both full-run reports pass verify.py.
"""
import copy
import hashlib
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
sys.path.insert(0, str(REPO))
from tools import generate_corpus_manifest_v2 as catalog_tool
from tools import perf_compare


def load(path):
    return json.loads(path.read_text())


def write(path, value):
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + '\n')


def main():
    report = load(ROOT / 'preflight-report.json')
    catalog = load(ROOT / 'preflight-catalog.json')
    previous = load(ROOT / 'inputs/perf-regression-default-manifest-v1.json')
    config = report['configuration']
    assert config['samples_per_case'] == 1 and config['warmup_iterations_per_case'] == 0
    rows = report['results']
    assert len(rows) == 213 and len(config['cases']) == 41
    identity = copy.deepcopy(previous)
    identity['default_cases'] = config['cases']
    identity['case_count'] = 41
    identity['result_count'] = 213
    identity['case_corpora'] = {case: [] for case in config['cases']}
    identity['corpora'] = {}
    keys = []
    for row in rows:
        assert len(row['elapsed_ns']['samples']) == 1
        corpus = row['corpus']
        name = corpus['name']
        assert name not in identity['corpora'] or identity['corpora'][name] == corpus
        identity['corpora'][name] = corpus
        identity['case_corpora'][row['case']].append(name)
        keys.append((row['case'], json.dumps(corpus, sort_keys=True, separators=(',', ':'))))
    for case in previous['default_cases']:
        assert identity['case_corpora'][case] == previous['case_corpora'][case], case
    for name, corpus in previous['corpora'].items():
        assert identity['corpora'][name] == corpus, name
    assert set(identity['default_cases']) - set(previous['default_cases']) == {'rtf_semantic_text_to_sink', 'odt_semantic_text_to_sink', 'ods_semantic_text_to_sink', 'odp_semantic_text_to_sink'}
    identity['result_keys_sha256'] = perf_compare.result_key_manifest_sha256(keys)
    for name, expected in identity['identity_configuration'].items():
        assert config[name] == expected, name
    generated = catalog_tool.generate(identity, catalog['build']['git_revision'],
                                      worktree_dirty=catalog['build']['git_worktree_dirty'])
    assert generated == catalog, 'Python/Rust catalog derivation differs'
    checked = ROOT / 'checked'
    checked.mkdir(exist_ok=True)
    for path, value in (
        ('docs/performance/results/perf-regression-default-manifest-v1.json', identity),
        ('docs/performance/results/perf-corpus-manifest-v2.json', catalog),
    ):
        write(REPO / path, value)
        write(checked / Path(path).name, value)
    policy_path = REPO / 'docs/performance/perf-regression-policy-v1.json'
    policy = load(ROOT / 'inputs' / policy_path.name)
    policy['policy_id'] = 'litchi-hosted-default-matrix-v4'
    policy['expected_result_count'] = 213
    policy['expected_result_keys_sha256'] = identity['result_keys_sha256']
    policy['required_cases'] = identity['default_cases']
    write(policy_path, policy)
    write(checked / policy_path.name, policy)
    index_path = REPO / 'docs/performance/crud-coverage-index-v1.json'
    index = load(ROOT / 'inputs' / index_path.name)
    index['checked_catalog'].update({k: catalog[k] for k in ('catalog_id', 'catalog_sha256', 'content_set_sha256')})
    category = index['categories'][2]
    category['checklist_refs'] = ['Scope is explicit', 'Performance claims are measured']
    category['status'] = 'measured'
    category['measurement'] = copy.deepcopy(index['categories'][0]['measurement'])
    for row in category['scenarios']:
        assert row['selector'] in set(identity['default_cases']) - set(previous['default_cases'])
        assert row['status'] == 'correctness-only'
        row['status'] = 'measured'
        corpora = [c for c in catalog['corpora'] if row['selector'] in c['coverage']['timed_cases']]
        row['corpus'] = {'kind': 'checked-catalog', 'case': row['selector'],
                         'ids': sorted(c['id'] for c in corpora),
                         'shapes': sorted(c['legacy_v1']['shape'] for c in corpora)}
    write(index_path, index)
    write(checked / index_path.name, index)
    print(json.dumps({'status': 'pass', 'cases': 41, 'rows': 213,
                      'result_keys_sha256': identity['result_keys_sha256'],
                      'catalog_sha256': catalog['catalog_sha256'],
                      'content_set_sha256': catalog['content_set_sha256']}))


if __name__ == '__main__':
    main()
