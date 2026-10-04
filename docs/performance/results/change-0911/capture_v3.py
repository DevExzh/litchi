"""Sequential same-binary snapshot instruction capture; first child requires raw admission."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PRIOR = P.parent / 'change-0910'
CWD = Path('/tmp/litchi-0911-run-v3')
OWNER = 'litchi_pptx::package::model::Package::opened_presentation_transaction'
PARENT = 'litchi_perf_baseline::pptx_owned_benchmark::timed_iteration'
SCHEDULE = [('materialized-a', 'materialized'), ('owned-a', 'owned'), ('owned-b', 'owned'), ('materialized-b', 'materialized')]
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
    assert not (P / 'plan-v3.json').exists() and not CWD.exists()
    binary = desc(PRIOR / 'binaries/native')
    assert binary['sha256'] == read(PRIOR / 'binary.json')['retained']['sha256'] == 'f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d'
    assert read(PRIOR / 'replay-after-cleanup.json')['status'] == 'pass'
    symbols = subprocess.check_output(['/usr/bin/nm', '-C', binary['path']], text=True)
    matches = []
    for function in (OWNER, PARENT):
        matched = [line for line in symbols.splitlines() if line.endswith(' ' + function)]
        assert len(matched) == 1, (function, matched)
        matches.extend(matched)
    with (P / 'owner-symbols-v3.txt').open('x') as stream:
        stream.write('\n'.join(matches) + '\n')
    CWD.mkdir()
    (CWD / '.owned-0911').write_text(binary['sha256'] + '\n')
    refs = ['source-v1.json', 'source-base.json', 'source-overlay-v1.patch', 'binary.json', 'build-v1/command.json', 'quality-v1/fmt/command.json', 'quality-v1/clippy/command.json', 'analysis.json', 'replay-after-cleanup.json', 'closure.json', 'capture/block-00-materialized/artifacts/report.json', 'capture/block-00-owned/artifacts/report.json']
    inputs = [desc(PRIOR / name) for name in refs] + [desc(P / 'owner-symbols-v3.txt'), desc(P.parent / 'change-0845/reader_v2.py'), desc(P / 'protocol-review-v3.md'), desc(P / 'protocol-clarification.md'), desc(P / 'first-child-failure.json'), desc(P / 'first-child-failure.md'), desc(P / 'first-child-failure-v2.json'), desc(P / 'first-child-failure-v2.md'), desc(P / 'plan.json'), desc(P / 'capture.py'), desc(P / 'plan-v2.json'), desc(P / 'capture_v2.py')]
    plan = {'schema': 'litchi.pptx-snapshot-instruction-profile-0911.v3', 'created_utc': now(), 'evidence_parent': '52a27130589915884be5fd10bb1ade2f835ed32e', 'compiled_source_parent': read(PRIOR / 'source-v1.json')['parent'], 'cpu': 12, 'environment': ENV, 'cwd': str(CWD), 'binary': binary, 'samples': 1, 'warmup': 0, 'owner': OWNER, 'parent': PARENT, 'options': ['--tool=callgrind', '--collect-atstart=no', '--toggle-collect=' + OWNER, '--zero-before=' + OWNER, '--dump-after=' + OWNER], 'schedule': [{'index': i, 'label': label, 'kind': kind} for i, (label, kind) in enumerate(SCHEDULE)], 'inputs': inputs, 'tools': [desc(name) for name in ['/usr/bin/taskset', '/usr/bin/valgrind', '/usr/libexec/valgrind/callgrind-amd64-linux', '/usr/bin/nm']], 'valgrind_version': subprocess.check_output(['/usr/bin/valgrind', '--version'], text=True).strip(), 'helper': desc(__file__), 'expected_positive_callers': {'setup': 'litchi_perf_baseline::build_pptx_source_edit_corpus', 'measured': PARENT}, 'scope': 'Exactly two independently conserved positive transaction-owner dumps, each with one exact incoming caller: setup corpus builder and measured timed_iteration. Retain both, classify by caller, never dump order. Only measured context enters route comparison. Require zero terminal dump. Instrumented elapsed durations are not native timing. No allocation, CPU-cycle or codec-byte inference.'}
    write(P / 'plan-v3.json', plan)
    print('0911 plan frozen', flush=True)


def validate_semantics(plan, row, folder):
    report = read(folder / 'artifacts/report.json')
    reference_path = PRIOR / f"capture/block-00-{row['kind']}/artifacts/report.json"
    reference = read(reference_path)
    assert report['schema'] == reference['schema'] == 'pptx-owned-input-benchmark-v1'
    assert report['status'] == 'completed' and report['route'] == row['kind']
    assert report['warmup_iterations'] == 0 and report['measured_samples'] == 1
    for key in ['implementation', 'corpus', 'source', 'output', 'sink_capacity_bytes', 'sink_max_write_bytes', 'selected_slide_positions', 'replacement_count', 'timing_scope', 'phase_scope', 'teardown_scope', 'verification_scope', 'performance_claim', 'gates']:
        assert report[key] == reference[key], key
    assert all(report['gates'].values())
    assert len(report['samples']) == 1
    sample = report['samples'][0]
    assert sample['index'] == 0
    phases = ['clone_ns', 'ingress_ns', 'snapshot_ns', 'edit_commit_ns', 'publish_ns']
    assert all(type(sample[k]) is int and sample[k] > 0 for k in phases + ['total_ns'])
    assert sum(sample[k] for k in phases) <= sample['total_ns']
    for key in ['output_sha256', 'output_bytes', 'sink_write_calls', 'sink_largest_write']:
        assert all(native[key] == sample[key] for native in reference['samples']), key
    for name in ('source', 'output'):
        actual = desc(folder / 'artifacts' / report[name]['file'])
        assert actual['bytes'] == report[name]['bytes'] and actual['sha256'] == report[name]['sha256']
    return {'status': 'pass', 'reference': desc(reference_path), 'report': desc(folder / 'artifacts/report.json'), 'route': row['kind'], 'corpus': report['corpus'], 'source': report['source'], 'output': report['output'], 'sink_write_calls': sample['sink_write_calls'], 'sink_largest_write': sample['sink_largest_write'], 'scope': 'Exact source/output/corpus and guard/sink equality to the 0910 native control. Instrumented durations excluded from native conclusions.'}


def capture(index):
    plan = read(P / 'plan-v3.json')
    row = plan['schedule'][index]
    assert row['index'] == index
    assert (CWD / '.owned-0911').read_text() == plan['binary']['sha256'] + '\n'
    for item in [plan['binary'], plan['helper'], *plan['inputs'], *plan['tools']]:
        assert desc(item['path']) == item, item['path']
    if index:
        assert read(P / 'first-child-admission-v3.json')['status'] == 'pass'
        previous = plan['schedule'][index - 1]['label']
        assert read(P / 'capture-v3' / previous / 'semantics.json')['status'] == 'pass'
    folder = P / 'capture-v3' / row['label']
    folder.mkdir(parents=True)
    argv = ['/usr/bin/taskset', '-c', str(plan['cpu']), '/usr/bin/valgrind', *plan['options'], '--callgrind-out-file=' + str(folder / 'callgrind.out'), plan['binary']['path'], row['kind'], '0', '1', str(folder / 'artifacts')]
    receipt = {'argv': argv, 'cwd': plan['cwd'], 'environment': plan['environment'], 'started_utc': now(), 'load_before': os.getloadavg(), 'plan': desc(P / 'plan-v3.json'), 'binary': plan['binary'], 'helper': plan['helper'], 'row': row}
    write(folder / 'start.json', receipt)
    print('start', row['label'], receipt['started_utc'], flush=True)
    with (folder / 'stdout.txt').open('x') as out, (folder / 'stderr.txt').open('x') as err:
        result = subprocess.run(argv, cwd=CWD, env=plan['environment'], stdout=out, stderr=err)
    receipt.update(exit_code=result.returncode, finished_utc=now(), load_after=os.getloadavg(), artifacts=[desc(f) for f in sorted(folder.rglob('*')) if f.is_file()])
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
