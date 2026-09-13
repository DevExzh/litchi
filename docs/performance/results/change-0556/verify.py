"""Verify the retained preparation evidence, without claiming performance."""
import difflib
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def read(name):
    return json.loads((HERE / name).read_text())


def main():
    preparation = read('preparation.json')
    baseline = read('baseline-source.json')
    candidate = read('candidate-source.json')
    source = preparation['production_source']
    assert baseline.keys() == candidate.keys()
    assert [p for p in baseline if baseline[p] != candidate[p]] == [source]
    assert sha(HERE / 'baseline-source.json') == preparation['baseline_source_sha256']
    assert sha(HERE / 'baseline-cell.rs') == baseline[source] == preparation['production_sha256']
    assert sha(HERE / 'candidate/cell.rs') == candidate[source]
    original = subprocess.check_output(
        ['git', 'show', preparation['base_revision'] + ':' + source], cwd=REPO)
    assert original == (HERE / 'baseline-cell.rs').read_bytes()
    patch = ''.join(difflib.unified_diff(
        original.decode().splitlines(keepends=True),
        (HERE / 'candidate/cell.rs').read_text().splitlines(keepends=True),
        fromfile='a/' + source, tofile='b/' + source))
    assert patch == (HERE / 'candidate/candidate.patch').read_text()
    binding = read('candidate/candidate.json')
    assert binding['production_sha256'] == baseline[source]
    assert binding['candidate_source_sha256'] == candidate[source]
    assert binding['candidate_patch_sha256'] == sha(HERE / 'candidate/candidate.patch')
    assert binding['source_review_sha256'] == sha(REPO / binding['source_review'])
    for name, digest in baseline.items():
        assert sha(REPO / name) == digest, name
    assert sha(REPO / 'Cargo.lock') == preparation['cargo_lock_sha256']
    adr = REPO / 'docs/performance/results/change-0555/adr-manifest.json'
    assert sha(adr) == preparation['adr_manifest_sha256']
    for name, digest in json.loads(adr.read_text()).items():
        assert sha(REPO / name) == digest, name

    plan = read('quality-plan.json')
    checked = []
    for result_path in sorted((HERE / 'quality-attempts').glob('*/result.json')):
        result = json.loads(result_path.read_text())
        inputs_path = result_path.parent / 'inputs.json'
        inputs = json.loads(inputs_path.read_text())
        assert sha(inputs_path) == result['inputs_sha256']
        assert inputs['plan_sha256'] == sha(HERE / 'quality-plan.json')
        assert inputs['script_sha256'] == sha(HERE / 'quality.py')
        manifest = HERE / (result['stage'] + '-source.json')
        if result_path.parent.name == 'candidate-01':
            manifest = HERE / 'candidate-attempts/before-format/candidate-source.json'
        assert inputs['source_manifest_sha256'] == sha(manifest)
        assert result['source_manifest_sha256'] == sha(manifest)
        commands = plan[result['mode']]
        assert inputs['commands'] == commands
        format_failure = result_path.parent.name == 'candidate-01'
        assert result['status'] == ('failed' if format_failure else 'pass')
        assert [r['name'] for r in result['rows']] == list(commands)
        for row in result['rows']:
            folder = HERE / row['path']
            receipt = json.loads((folder / 'receipt.json').read_text())
            assert sha(folder / 'receipt.json') == row['receipt_sha256']
            expected_code = 1 if format_failure and row['name'] == 'format' else 0
            assert receipt['exit_code'] == row['exit_code'] == expected_code
            assert receipt['command'] == commands[row['name']]
            assert receipt['source_manifest_sha256'] == sha(manifest)
            assert receipt['inputs_sha256'] == sha(inputs_path)
            assert receipt['source_stable'] is True
            for name, digest in receipt['artifacts'].items():
                assert sha(folder / name) == digest
        checked.append(result_path.parent.name)

    aborted_root = HERE / 'quality-attempts/candidate-02'
    aborted = json.loads((aborted_root / 'aborted.json').read_text())
    assert aborted['status'] == 'inadmissible-source-changed'
    assert aborted['admissible_quality_evidence'] is False
    assert aborted['observed_exit_code'] == 1
    assert not (aborted_root / 'result.json').exists()
    for name, digest in aborted['artifacts'].items():
        assert sha(aborted_root / name) == digest
    assert sorted(p.name for p in (HERE / 'quality-attempts').iterdir()) == sorted(checked + ['candidate-02'])
    correction = read('format-correction.json')
    old_source = HERE / 'candidate-attempts/before-format/candidate/cell.rs'
    assert sha(old_source) == correction['before_source_sha256']
    assert sha(HERE / 'candidate/cell.rs') == correction['after_source_sha256']
    assert sha(HERE / 'candidate/candidate.patch') == correction['after_patch_sha256']
    for path_key, hash_key in [('receipt', 'receipt_sha256'), ('failed_check', 'failed_check_sha256')]:
        assert sha(HERE / correction[path_key]) == correction[hash_key]
    assert old_source.read_text().split('#[cfg(test)]\nmod tests {')[0] == (
        HERE / 'candidate/cell.rs').read_text().split('#[cfg(test)]\nmod tests {')[0]
    for path in [old_source, HERE / 'candidate/cell.rs']:
        formatted = subprocess.run(['rustfmt', '--edition', '2024', '--emit', 'stdout'],
                                   cwd=REPO, input=path.read_bytes(), capture_output=True,
                                   check=True)
        assert formatted.stdout == (HERE / 'candidate/cell.rs').read_bytes()

    final_inputs = read('final-checks/inputs.json')
    final_result = read('final-checks/result.json')
    assert final_inputs['driver_sha256'] == sha(HERE / 'final_checks.py')
    assert final_inputs['source_manifest_sha256'] == sha(HERE / 'baseline-source.json')
    assert final_result['status'] == 'pass'
    assert [r['name'] for r in final_result['rows']] == list(final_inputs['commands'])
    for row in final_result['rows']:
        name = row['name']
        path = HERE / 'final-checks' / (name + '.json')
        receipt = json.loads(path.read_text())
        assert sha(path) == row['receipt_sha256']
        assert receipt['command'] == final_inputs['commands'][name]
        assert receipt['exit_code'] == row['exit_code'] == 0
        assert receipt['source_stable'] is True
        assert receipt['inputs_sha256'] == sha(HERE / 'final-checks/inputs.json')
        for stream in ['stdout', 'stderr']:
            assert sha(HERE / 'final-checks' / (name + '.' + stream)) == receipt[stream + '_sha256']

    restoration = read('restoration.json')
    assert restoration['before_sha256'] == candidate[source]
    assert restoration['after_sha256'] == baseline[source]
    assert restoration['all_manifest_files_match'] is True
    for label, digest in restoration['candidate_quality'].items():
        assert sha(HERE / 'quality-attempts' / label / 'result.json') == digest
        assert read('quality-attempts/' + label + '/result.json')['status'] == 'pass'

    decision = read('decision.json')
    assert decision['status'] == 'prepared-awaiting-performance-measurement'
    assert decision['adoption_allowed'] is False
    assert sorted(decision['quality_attempts']) == checked
    for name, digest in decision['evidence'].items():
        assert sha(HERE / name) == digest, name
    for name, digest in read('documentation.json').items():
        assert sha(REPO / name) == digest, name
    for label, expected in decision['test_counts'].items():
        output = (HERE / 'quality-attempts' / label / 'xlsx-tests/stdout').read_text()
        rows = re.findall(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored', output)
        actual = dict(groups=len(rows), passed=sum(int(r[1]) for r in rows),
                      failed=sum(int(r[2]) for r in rows), ignored=sum(int(r[3]) for r in rows))
        assert actual == expected, (label, actual)
    if '--cleaned' in sys.argv:
        cleanup = read('cleanup.json')
        target = Path(preparation['target'])
        assert cleanup['target'] == str(target)
        assert cleanup['removed'] is True and not target.exists()
        assert not list(HERE.rglob('__pycache__'))
    if '--sealed' in sys.argv:
        seal = {}
        for line in (HERE / 'SHA256SUMS').read_text().splitlines():
            digest, name = line.split('  ', 1)
            assert name not in seal
            seal[name] = digest
        files = sorted(p.relative_to(HERE).as_posix() for p in HERE.rglob('*')
                       if p.is_file() and p.name != 'SHA256SUMS')
        assert list(seal) == files
        for name, digest in seal.items():
            assert sha(HERE / name) == digest, name
    print(json.dumps(dict(status='pass', quality_attempts=checked,
                          adoption_allowed=False, production_restored=True)))


if __name__ == '__main__':
    main()
