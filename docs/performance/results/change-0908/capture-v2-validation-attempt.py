"""Root-only sequential instruction capture on the retained 0907 native binary."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PRIOR = P.parent / 'change-0907'
CWD = Path('/tmp/litchi-0908-run')
OWNER = 'litchi_perf_baseline::publish_pptx_multi_slide_batch_edit'
SCHEDULE = [('eager-a', 'eager'), ('source-a', 'source_backed'), ('source-b', 'source_backed'), ('eager-b', 'eager')]
ENV = {'PATH': '/usr/bin:/bin', 'LC_ALL': 'C', 'LANG': 'C', 'TZ': 'UTC'}

def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()

def read(path):
    return json.loads(Path(path).read_text())

def desc(path):
    path = Path(path).resolve()
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}

def write(path, value):
    with Path(path).open('x') as out:
        json.dump(value, out, indent=2, sort_keys=True)
        out.write('\n')

def freeze():
    assert not (P / 'plan.json').exists() and not CWD.exists()
    binary = desc(PRIOR / 'binaries/native')
    assert binary['sha256'] == read(PRIOR / 'binary.json')['sha256']
    assert read(PRIOR / 'replay-after-cleanup.json')['status'] == 'pass'
    symbols = subprocess.check_output(['/usr/bin/nm', '-C', binary['path']], text=True)
    matched = [line for line in symbols.splitlines() if line.endswith(' ' + OWNER)]
    assert len(matched) == 1, matched
    (P / 'owner-symbol.txt').write_text('\n'.join(matched) + '\n')
    CWD.mkdir()
    (CWD / '.owned-0908').write_text(binary['sha256'] + '\n')
    refs = ['source.json', 'binary.json', 'native-build-v2-run.json', 'analysis.json', 'replay-after-cleanup.json', 'closure.json', *[label + '.json' for label, _ in SCHEDULE]]
    plan = {'schema': 'litchi.pptx-instruction-profile-0908.v1', 'created_utc': now(), 'cpu': 12, 'environment': ENV, 'cwd': str(CWD), 'binary': binary, 'source_commit': read(PRIOR / 'source.json')['commit'], 'samples': 1, 'warmup': 0, 'owner': OWNER, 'parent': 'litchi_perf_baseline::run_pptx_multi_slide_batch_edit_save', 'options': ['--tool=callgrind', '--collect-atstart=no', '--toggle-collect=' + OWNER, '--zero-before=' + OWNER, '--dump-after=' + OWNER], 'schedule': [{'index': i, 'label': label, 'kind': kind, 'case': 'pptx_' + kind + '_multi_slide_batch_edit_save'} for i, (label, kind) in enumerate(SCHEDULE)], 'inputs': [desc(PRIOR / name) for name in refs] + [desc(P / 'owner-symbol.txt'), desc(P.parent / 'change-0845/reader_v2.py')], 'tools': [desc(name) for name in ['/usr/bin/taskset', '/usr/bin/valgrind', '/usr/libexec/valgrind/callgrind-amd64-linux', '/usr/bin/nm']], 'valgrind_version': subprocess.check_output(['/usr/bin/valgrind', '--version'], text=True).strip(), 'helper': desc(__file__), 'scope': 'Collected Ir inside exact existing guarded helper only. Instrumented elapsed samples excluded from native timing conclusions; first child requires raw scope admission before continuing.'}
    write(P / 'plan.json', plan)
    print('plan frozen', flush=True)

def validate_semantics(plan, row, folder):
    report = read(folder / 'report.json')
    reference = read(PRIOR / (row['label'] + '.json'))
    assert report['schema_version'] == 1
    assert report['binary_identity']['binary_sha256'] == plan['binary']['sha256']
    assert report['environment']['cpu_affinity'] == str(plan['cpu'])
    assert report['tool']['profile'] == 'release' and report['tool']['instrumentation'] == 'none'
    for key in ['git_revision', 'git_worktree_dirty', 'rustc_version']:
        assert report['environment'][key] is None, key
    config = report['configuration']
    assert config['samples_per_case'] == 1 and config['warmup_iterations_per_case'] == 0 and config['cases'] == [row['case']]
    assert len(report['results']) == 1
    result = report['results'][0]
    expected = reference['results'][0]
    assert result['case'] == row['case']
    for key in ['corpus', 'sink', 'output_sha256']:
        assert result[key] == expected[key], key
    metrics = result['operation_metrics']
    assert metrics['sample_count'] == 1 and metrics['sample_indices'] == [0]
    reference_metrics = expected['operation_metrics']
    assert {key: value for key, value in metrics.items() if key not in ['sample_count', 'sample_indices']} == {key: value for key, value in reference_metrics.items() if key not in ['sample_count', 'sample_indices']}
    assert len(result['elapsed_ns']['samples']) == 1 and result['elapsed_ns']['samples'][0] > 0 and result['elapsed_ns']['sample_order'] == [0]
    assert result['source'].keys() == expected['source'].keys()
    for key, values in result['source'].items():
        assert len(values) == 1 and values[0] > 0
        assert all(value == values[0] for value in expected['source'][key]), key
    return {'status': 'pass', 'reference': desc(PRIOR / (row['label'] + '.json')), 'case': row['case'], 'corpus': result['corpus'], 'sink': result['sink'], 'output_sha256': result['output_sha256'], 'source': result['source'], 'scope': 'Exact corpus/output/sink and per-sample logical source-counter equality to current native reference. Instrumented duration excluded.'}

def capture(index):
    plan = read(P / 'plan-v2.json')
    row = plan['schedule'][index]
    assert row['index'] == index
    assert (CWD / '.owned-0908').read_text() == plan['binary']['sha256'] + '\n'
    for item in [plan['binary'], plan['helper'], *plan['inputs'], *plan['tools']]:
        assert desc(item['path']) == item, item['path']
    if index:
        assert read(P / 'first-child-admission.json')['status'] == 'pass'
        previous = plan['schedule'][index - 1]['label']
        assert read(P / 'capture' / previous / 'semantics.json')['status'] == 'pass'
    folder = P / 'capture' / row['label']
    folder.mkdir(parents=True)
    argv = ['/usr/bin/taskset', '-c', str(plan['cpu']), '/usr/bin/valgrind', *plan['options'], '--callgrind-out-file=' + str(folder / 'callgrind.out'), plan['binary']['path'], '--workers', '1', '--warmup', '0', '--samples', '1', '--case', row['case'], '--json', str(folder / 'report.json')]
    receipt = {'argv': argv, 'cwd': plan['cwd'], 'environment': plan['environment'], 'started_utc': now(), 'load_before': os.getloadavg(), 'plan': desc(P / 'plan-v2.json'), 'binary': plan['binary'], 'helper': plan['helper'], 'row': row}
    write(folder / 'start.json', receipt)
    print('start', row['label'], receipt['started_utc'], flush=True)
    with (folder / 'stdout.txt').open('x') as out, (folder / 'stderr.txt').open('x') as err:
        result = subprocess.run(argv, cwd=CWD, env=plan['environment'], stdout=out, stderr=err)
    receipt.update(exit_code=result.returncode, finished_utc=now(), load_after=os.getloadavg(), artifacts=[desc(f) for f in sorted(folder.iterdir()) if f.is_file()])
    write(folder / 'command.json', receipt)
    assert result.returncode == 0, receipt
    try:
        semantics = validate_semantics(plan, row, folder)
    except Exception as error:
        write(folder / 'semantic-failure.json', {'status': 'fail', 'error': repr(error)})
        raise
    write(folder / 'semantics.json', semantics)
    assert desc(plan['binary']['path']) == plan['binary']
    print('finish', row['label'], 'semantic pass', flush=True)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=['freeze', 'capture'])
    parser.add_argument('--index', type=int)
    args = parser.parse_args()
    if args.action == 'freeze':
        freeze()
    else:
        assert args.index is not None and 0 <= args.index < len(SCHEDULE)
        capture(args.index)
