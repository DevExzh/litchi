"""Same-binary native cycle sampling with exact frozen commands and semantic gates."""
import argparse
import datetime
import gzip
import hashlib
import json
import os
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PRIOR = P.parent / 'change-0910'
CWD = Path('/tmp/litchi-0912-run')
ENV = {'PATH': '/usr/bin:/bin', 'LC_ALL': 'C', 'LANG': 'C', 'TZ': 'UTC', 'DEBUGINFOD_URLS': ''}
SCHEDULE = [('materialized-a', 'materialized'), ('owned-a', 'owned'), ('owned-b', 'owned'), ('materialized-b', 'materialized')]
OWNER = 'litchi_pptx::package::model::Package::opened_presentation_transaction'
PARENT = 'litchi_perf_baseline::pptx_owned_benchmark::timed_iteration'


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def read(path):
    return json.loads(Path(path).read_text())


def sha(data):
    return hashlib.sha256(data).hexdigest()


def desc(path):
    path = Path(path).resolve()
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': sha(path.read_bytes())}


def write(path, value):
    with Path(path).open('x') as out:
        json.dump(value, out, indent=2, sort_keys=True)
        out.write('\n')


def freeze():
    assert not (P / 'plan.json').exists() and not CWD.exists()
    binary = desc(PRIOR / 'binaries/native')
    assert binary['sha256'] == read(PRIOR / 'binary.json')['retained']['sha256'] == 'f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d'
    assert read(P / 'perf-preflight.json')['exit_code'] == 0
    symbols = subprocess.check_output(['/usr/bin/nm', '-S', '-C', '--defined-only', binary['path']], text=True)
    required = [OWNER, PARENT, 'litchi_perf_baseline::build_pptx_source_edit_corpus', 'sha2::sha256::compress256', 'sha2::sha256::x86_sha::compress', 'sha2::sha256::soft::unroll::compress']
    selected = []
    for function in required:
        matches = [line for line in symbols.splitlines() if line.endswith(' ' + function)]
        assert len(matches) == 1, (function, matches)
        selected += matches
    with (P / 'symbols.txt').open('x') as out:
        out.write('\n'.join(selected) + '\n')
    ranges = []
    for line in selected:
        address, size, kind, name = line.split(maxsplit=3)
        ranges.append({'name': name, 'address': int(address, 16), 'size': int(size, 16), 'end': int(address, 16) + int(size, 16), 'kind': kind})
    write(P / 'symbol-ranges.json', {'binary': binary, 'ranges': ranges})
    elf = subprocess.check_output(['/usr/bin/readelf', '-lW', binary['path']])
    (P / 'elf-segments.txt').write_bytes(elf)
    CWD.mkdir()
    (CWD / '.owned-0912').write_text(binary['sha256'] + '\n')
    refs = [PRIOR / n for n in ('source-v1.json', 'binary.json', 'build-v1/command.json', 'quality-v1/fmt/command.json', 'quality-v1/clippy/command.json', 'analysis.json', 'closure.json', 'capture/block-00-owned/artifacts/report.json', 'capture/block-00-materialized/artifacts/report.json')]
    refs += [P.parent / 'change-0911' / n for n in ('analysis.json', 'replay-after-cleanup.json', 'closure.json')]
    refs += [P / n for n in ('perf-preflight.json', 'perf-preflight.stdout.txt', 'perf-preflight.stderr.txt', 'protocol-review.md', 'symbols.txt', 'symbol-ranges.json', 'elf-segments.txt')]
    options = ['record', '--no-buildid-cache', '--compression-level=3', '-e', 'cycles:u', '-F', '997', '--strict-freq', '--call-graph', 'dwarf,16384', '--sample-cpu']
    write(P / 'plan.json', {'schema': 'litchi.native-snapshot-profile-0912.v1', 'utc': now(), 'evidence_parent': '9a108730aa5533678b376ae097cd526c26e92b93', 'binary': binary, 'helper': desc(__file__), 'inputs': [desc(f) for f in refs], 'tools': [desc(f) for f in ('/usr/bin/perf', '/usr/bin/taskset', '/usr/bin/nm', '/usr/bin/readelf', '/usr/lib/x86_64-linux-gnu/libc.so.6')], 'perf_version': subprocess.check_output(['/usr/bin/perf', '--version'], text=True).strip(), 'cpu': 12, 'environment': ENV, 'cwd': str(CWD), 'samples': 50, 'warmup': 0, 'event': 'cycles:u', 'frequency_hz': 997, 'stack_bytes': 16384, 'script_max_stack': 256, 'options': options, 'decode_options': ['script', '--no-inline', '--ns', '--show-lost-events', '--demangle', '--max-stack', '256', '-F', 'comm,pid,tid,cpu,time,period,event,ip,sym,dso'], 'owner': OWNER, 'parent': PARENT, 'setup_parent': 'litchi_perf_baseline::build_pptx_source_edit_corpus', 'schedule': [{'index': i, 'label': label, 'kind': kind} for i, (label, kind) in enumerate(SCHEDULE)], 'scope': 'Whole-process native user-cycle samples retained. Measured snapshot requires exact owner and timed_iteration frames in the retained binary DSO. Setup is separate. Counts and sampled-period weights are descriptive, not exact elapsed-time or cycle totals. No debugger or modified executable.'})
    print('0912 native sampling plan frozen', flush=True)


def validate_semantics(plan, row, folder):
    report = read(folder / 'artifacts/report.json')
    reference_path = PRIOR / f"capture/block-00-{row['kind']}/artifacts/report.json"
    reference = read(reference_path)
    assert report['schema'] == reference['schema'] == 'pptx-owned-input-benchmark-v1'
    assert report['status'] == 'completed' and report['route'] == row['kind']
    assert report['warmup_iterations'] == 0 and report['measured_samples'] == plan['samples']
    for key in ['implementation', 'corpus', 'source', 'output', 'sink_capacity_bytes', 'sink_max_write_bytes', 'selected_slide_positions', 'replacement_count', 'timing_scope', 'phase_scope', 'teardown_scope', 'verification_scope', 'performance_claim', 'gates']:
        assert report[key] == reference[key], key
    assert all(report['gates'].values()) and len(report['samples']) == plan['samples']
    for i, sample in enumerate(report['samples']):
        assert sample['index'] == i
        phases = ['clone_ns', 'ingress_ns', 'snapshot_ns', 'edit_commit_ns', 'publish_ns']
        assert all(type(sample[k]) is int and sample[k] > 0 for k in phases + ['total_ns'])
        assert sum(sample[k] for k in phases) <= sample['total_ns']
        for key in ['output_sha256', 'output_bytes', 'sink_write_calls', 'sink_largest_write']:
            assert all(native[key] == sample[key] for native in reference['samples']), key
    for name in ('source', 'output'):
        actual = desc(folder / 'artifacts' / report[name]['file'])
        assert actual['bytes'] == report[name]['bytes'] and actual['sha256'] == report[name]['sha256']
    return {'status': 'pass', 'reference': desc(reference_path), 'report': desc(folder / 'artifacts/report.json'), 'route': row['kind'], 'operations': plan['samples'], 'source': report['source'], 'output': report['output'], 'scope': 'All native-control corpus/output/guards/sink and phase arithmetic pass. Profiled timings are not used for latency comparison.'}


def capture(index):
    plan = read(P / 'plan.json')
    row = plan['schedule'][index]
    for item in [plan['binary'], plan['helper'], *plan['inputs'], *plan['tools']]:
        assert desc(item['path']) == item, item['path']
    assert (CWD / '.owned-0912').read_text() == plan['binary']['sha256'] + '\n'
    if index:
        assert read(P / 'first-child-admission.json')['status'] == 'pass'
        assert read(P / 'capture' / plan['schedule'][index - 1]['label'] / 'semantics.json')['status'] == 'pass'
    folder = P / 'capture' / row['label']
    folder.mkdir(parents=True)
    raw = folder / 'perf.data'
    argv = ['/usr/bin/taskset', '-c', '12', '/usr/bin/perf', *plan['options'], '-o', str(raw), '--', plan['binary']['path'], row['kind'], '0', str(plan['samples']), str(folder / 'artifacts')]
    receipt = {'argv': argv, 'cwd': str(CWD), 'environment': ENV, 'started_utc': now(), 'load_before': os.getloadavg(), 'plan': desc(P / 'plan.json'), 'binary': plan['binary'], 'helper': plan['helper'], 'row': row}
    write(folder / 'start.json', receipt)
    print('start', row['label'], flush=True)
    with (folder / 'stdout.txt').open('x') as out, (folder / 'stderr.txt').open('x') as err:
        result = subprocess.run(argv, cwd=CWD, env=ENV, stdout=out, stderr=err)
    receipt.update(exit_code=result.returncode, finished_utc=now(), load_after=os.getloadavg(), artifacts=[desc(f) for f in sorted(folder.rglob('*')) if f.is_file()])
    write(folder / 'command.json', receipt)
    assert result.returncode == 0
    write(folder / 'semantics.json', validate_semantics(plan, row, folder))
    decode = ['/usr/bin/perf', *plan['decode_options'], '-i', str(raw)]
    result = subprocess.run(decode, cwd=CWD, env=ENV, capture_output=True)
    with (folder / 'script.txt.gz').open('xb') as out:
        out.write(gzip.compress(result.stdout, compresslevel=6, mtime=0))
    with (folder / 'decode.stderr.txt').open('xb') as out:
        out.write(result.stderr)
    write(folder / 'decode.json', {'argv': decode, 'exit_code': result.returncode, 'environment': ENV, 'cwd': str(CWD), 'raw': desc(raw), 'script': desc(folder / 'script.txt.gz'), 'stderr': desc(folder / 'decode.stderr.txt'), 'uncompressed': {'bytes': len(result.stdout), 'sha256': sha(result.stdout)}, 'helper': plan['helper'], 'perf': plan['tools'][0]})
    assert result.returncode == 0
    for name, options in [('mappings', ['script', '--no-inline', '--ns', '--show-mmap-events', '--show-task-events']), ('header', ['script', '--header-only']), ('buildids', ['buildid-list']), ('sample-debug', ['script', '-d', '--no-inline', '--ns', '--show-lost-events'])]:
        command = ['/usr/bin/perf', *options, '-i', str(raw)]
        result = subprocess.run(command, cwd=CWD, env=ENV, capture_output=True)
        # debug mode output is retained compressed if perf emits any samples.
        (folder / (name + '.stdout.gz')).write_bytes(gzip.compress(result.stdout, mtime=0))
        (folder / (name + '.stderr.txt')).write_bytes(result.stderr)
        write(folder / (name + '.json'), {'argv': command, 'exit_code': result.returncode, 'stdout': desc(folder / (name + '.stdout.gz')), 'stderr': desc(folder / (name + '.stderr.txt')), 'raw': desc(raw)})
        assert result.returncode == 0
    print('finish', row['label'], 'semantic and export PASS', flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('freeze', 'capture'))
    parser.add_argument('--index', type=int, choices=range(4))
    args = parser.parse_args()
    if args.action == 'freeze':
        freeze()
    else:
        assert args.index is not None
        capture(args.index)
