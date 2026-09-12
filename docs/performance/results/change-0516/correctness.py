"""Replay report, corpus, sink, and public-oracle checks from retained vectors."""
import copy
import importlib.util
import json
import sys

from run import HERE, REPO
from audit import read


def prior_verifier():
    path = HERE.parent / 'change-0514/verify.py'
    spec = importlib.util.spec_from_file_location('verify_0514_for_0516', path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    module.BASE_REVISION = read(HERE / 'plan.json')['base_revision']
    return module


def check():
    v = prior_verifier()
    identities = {}
    captures = []
    for path in sorted(HERE.glob('*/*-report.json')):
        stage = path.parent.name
        if stage not in ('before', 'after'):
            continue
        name = path.name.removesuffix('-report.json')
        family = 'fallback' if name.startswith('fallback-') else 'guard' if name.startswith('guard-') else 'main'
        lane = name.removeprefix(family + '-') if family != 'main' else name
        allocator = lane.startswith('allocator-')
        samples, warmups = (10, 1) if allocator else (1, 0) if lane == 'preflight' else (20, 2) if lane == 'pilot' else (3, 0) if lane == 'profile' else (500, 5) if family == 'main' else (100, 3)
        report = read(path)
        rows = []
        if family == 'main':
            cases = ['xlsx_one_percent_commit_save'] if lane == 'profile' else list(v.CASES)
            shapes = ['dense-wide'] if lane == 'profile' else list(v.SHAPES)
            build = read(path.parent / ('allocator-build-receipt.json' if allocator else 'build-receipt.json'))
            v.verify_report_identity(report, build, samples, warmups, cases, shapes, str(path), allocator)
            v.validate_binding(report, read(path.parent / (name + '-catalog.json')))
            expected = [(shape, case) for shape in shapes for case in cases]
            assert len(report['results']) == len(expected), path
            for row, (shape, case) in zip(report['results'], expected):
                v.verify_row(row, case, shape, samples, str(path), allocator)
                identity = {'corpus': row['corpus'], 'output_sha256': row.get('output_sha256'), 'sink': row.get('sink')}
                key = (family, shape, case)
                assert identities.setdefault(key, identity) == identity, (path, key)
                rows.append(key)
        else:
            assert report['samples'] == samples and report['warmups'] == warmups, path
            assert [shape['shape'] for shape in report['shapes']] == list(v.SHAPES), path
            for shape in report['shapes']:
                expected_scenarios = ['warm-changed-one-cell', 'warm-changed-one-percent'] if family == 'fallback' else list(v.GUARD_SCENARIOS)
                assert [row['scenario'] for row in shape['scenarios']] == expected_scenarios, path
                identity = {key: value for key, value in shape.items() if key != 'scenarios'}
                key = (family, shape['shape'])
                assert identities.setdefault(key, identity) == identity, (path, key)
                if family == 'fallback':
                    assert shape['corpus_variant'] == 'numeric-two-sheet-public-default-descent-x14ac-v1'
                    assert shape['source_markers'] == dict.fromkeys(['worksheet_parts_checked', 'x14ac_namespace_bindings', 'mce_namespace_bindings', 'mce_ignorable_attributes', 'dy_descent_attributes'], 2)
                    assert shape['descent_readback'] == {'expected': 0.2, 'sheets_checked': 2, 'all_sheets_match': True}
                for row in shape['scenarios']:
                    checked = copy.deepcopy(row)
                    if family == 'fallback':
                        for field in ['output_serialized', 'untouched_data_equal', 'defaults_descent_readback', 'source_markers']:
                            assert checked['oracle'].pop(field) is True, (path, field)
                    v._guard_scenario(checked, shape['shape'], samples, warmups, str(path), allocator)
                    rows.append((family, shape['shape'], row['scenario']))
        captures.append({'path': str(path.relative_to(HERE)), 'rows': len(rows), 'samples_per_row': samples, 'warmups_per_row': warmups})
    return {'captures': captures, 'corpus_and_output_identity_groups': len(identities),
            'scope': 'Existing reports only; required lane completeness is a separate final check.'}


if __name__ == '__main__':
    print(json.dumps(check(), indent=2))
