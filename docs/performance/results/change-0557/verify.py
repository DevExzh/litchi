#!/usr/bin/env python3
"""Verify the retained, stopped 0557 pilot and its source/artifact custody."""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
TARGET = Path('/home/zhuhe/litchi-goal-0557-target')
EXTERNAL = [REPO / 'docs/performance' / name for name in (
    '0557-xlsx-allocation-noise.md', 'BASELINE.md', 'HOTSPOTS.md')]


def sha(path):
    assert path.is_file() and not path.is_symlink(), path
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path):
    return json.loads(path.read_text())


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


def seal_paths():
    return sorted([p for p in HERE.rglob('*') if p.is_file()
                   and p != HERE / 'seal.json'] + EXTERNAL)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cleaned', action='store_true')
    parser.add_argument('--sealed', action='store_true')
    args = parser.parse_args()
    assert not (HERE / '__pycache__').exists()
    for name, expected in read(HERE / 'frozen-inputs.json')['sha256'].items():
        assert sha(HERE / name) == expected, name
    driver = module('driver0557', HERE / 'run.py')
    manifest = read(HERE / 'baseline/source-manifest.json')
    assert set(manifest) == set(driver.source_paths())
    for name, expected in manifest.items():
        assert sha(REPO / name) == expected, name
    assert sha(REPO / 'crates/litchi-xlsx/src/cell.rs') == (
        '7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7')
    assert not (HERE / 'candidate').exists()
    adrs = read(HERE.parent / 'change-0555/adr-manifest.json')
    for name, expected in adrs.items():
        assert sha(REPO / name) == expected, name
    plan = read(HERE / 'plan.json')
    expected_names = {x['name'] for x in plan['quality']['harness_commands']}
    expected_names.update(('workspace-check', 'boundaries', 'claims', 'build-normal'))
    expected_names.update(
        f'noise-r{repeat}-{shape}-{entry["case"]}'
        for repeat in (1, 2) for shape in plan['corpus']['shapes']
        for entry in plan['workloads']['primary'])
    receipts = sorted((HERE / 'baseline').glob('*.receipt.json'))
    assert {p.name.removesuffix('.receipt.json') for p in receipts} == expected_names
    for path in receipts:
        receipt = read(path)
        assert receipt['success'] and receipt['exit_code'] == 0, path
        assert receipt['child_started'] and receipt['spawn_error'] is None
        assert receipt['stage'] == receipt['execution_stage'] == 'baseline'
        assert receipt['plan_sha256'] == sha(HERE / 'plan.json')
        assert receipt['script_sha256'] == sha(HERE / 'run.py')
        assert receipt['output_manifest_unchanged']
        assert receipt['execution_manifest_unchanged']
        for guard in ('output_stage_before', 'output_stage_after',
                      'execution_stage_before', 'execution_stage_after'):
            assert receipt[guard]['ok'], (path, guard)
            assert receipt[guard]['identity']['manifest_sha256'] == sha(
                HERE / 'baseline/source-manifest.json')
        for name, expected in receipt['artifacts'].items():
            assert Path(name).name == name
            assert sha(path.parent / name) == expected, (path, name)
    if args.cleaned:
        cleanup = read(HERE / 'cleanup.json')
        assert cleanup['removed_target'] == str(TARGET)
        assert cleanup['removed'] is True and not TARGET.exists()
        assert not TARGET.is_symlink()
        descriptor = read(HERE / 'baseline/binary-normal.json')
        assert cleanup['binary_sha256_by_kind']['baseline-normal'] == descriptor['sha256']
    analyzer = module('analysis0557', HERE / 'analyze.py')
    result = analyzer.analyze(noise_only=True)
    assert result == read(HERE / 'noise-analysis.json')
    assert result['status'] == 'too_unstable'
    assert result['noise']['N_percent'] > 5
    assert len(result['noise']['rows']) == 8
    subprocess.run([sys.executable, '-B', str(HERE / 'review_noise.py')], check=True)
    review = read(HERE / 'noise-review.json')
    assert review['comparison_count'] == 464 and review['flag_count'] == 135
    if args.sealed:
        seal = read(HERE / 'seal.json')
        actual = {str(p.relative_to(REPO)): sha(p) for p in seal_paths()}
        assert actual == seal['sha256']
    print(json.dumps({'status': 'pass', 'source_files': len(manifest),
                      'adr_files': len(adrs), 'successful_receipts': len(receipts),
                      'noise_children': 16, 'measured_samples': 16000,
                      'candidate_measured': False, 'noise_N_percent': result['noise']['N_percent'],
                      'cleaned': args.cleaned, 'sealed': args.sealed}, sort_keys=True))


if __name__ == '__main__':
    main()
