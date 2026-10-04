"""Freeze and execute owner-gated GDB dispatch witnesses on the retained binary."""
import argparse
import importlib.util
import json
import os
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('capture_0912', P / 'capture.py')
cap = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cap)
CWD = Path('/tmp/litchi-0912-gdb-run')


def freeze():
    assert not (P / 'plan-gdb.json').exists() and not CWD.exists()
    prior = cap.read(P / 'plan.json')
    assert cap.desc(prior['binary']['path']) == prior['binary']
    CWD.mkdir()
    (CWD / '.owned-0912').write_text(prior['binary']['sha256'] + '\n')
    refs = [P / n for n in ('plan.json', 'capture.py', 'symbol-ranges.json', 'elf-segments.txt', 'protocol-review-gdb.md', 'capture/materialized-a/command.json', 'capture/materialized-a/semantics.json', 'capture/materialized-a/decode.json', 'capture/materialized-a/sample-debug.json')]
    cap.write(P / 'plan-gdb.json', {'schema': 'litchi.native-sha-dispatch-0912.v1', 'utc': cap.now(), 'binary': prior['binary'], 'parent': prior['evidence_parent'], 'helper': cap.desc(__file__), 'gdb_script': cap.desc(P / 'dispatch_gdb.py'), 'inputs': [cap.desc(f) for f in refs], 'tools': [cap.desc('/usr/bin/gdb'), cap.desc('/usr/bin/taskset')], 'gdb_version': subprocess.check_output(['/usr/bin/gdb', '--version'], text=True).splitlines()[0], 'cwd': str(CWD), 'environment': cap.ENV, 'cpu': 12, 'samples': 1, 'warmup': 0, 'schedule': [{'index': 0, 'kind': 'materialized'}, {'index': 1, 'kind': 'owned'}], 'scope': 'ABI-scoped native SHA backend reachability only. Every owner entry and return checked; backend PC/backtrace bound to active measured owner. No perf attribution or timing result.'})
    print('0912 GDB plan frozen', flush=True)


def capture(index):
    plan = cap.read(P / 'plan-gdb.json')
    for item in [plan['binary'], plan['helper'], plan['gdb_script'], *plan['inputs'], *plan['tools']]:
        assert cap.desc(item['path']) == item, item['path']
    assert (CWD / '.owned-0912').read_text() == plan['binary']['sha256'] + '\n'
    row = plan['schedule'][index]
    if index:
        assert cap.read(P / 'dispatch/materialized/semantics.json')['status'] == 'pass'
        assert cap.read(P / 'dispatch/materialized/dispatch.json')['status'] == 'pass'
    folder = P / 'dispatch' / row['kind']
    folder.mkdir(parents=True)
    env = dict(plan['environment'], LITCHI_0912_DISPATCH_LOG=str(folder / 'dispatch.json'), LITCHI_0912_SYMBOL_RANGES=str(P / 'symbol-ranges.json'))
    argv = ['/usr/bin/taskset', '-c', '12', '/usr/bin/gdb', '--batch', '-nx', '-nh', '-iex', 'set auto-load off', '-iex', 'set debuginfod enabled off', '-x', plan['gdb_script']['path'], '--args', plan['binary']['path'], row['kind'], '0', '1', str(folder / 'artifacts')]
    receipt = {'argv': argv, 'environment': env, 'cwd': str(CWD), 'started_utc': cap.now(), 'plan': cap.desc(P / 'plan-gdb.json'), 'binary': plan['binary'], 'helper': plan['helper'], 'gdb_script': plan['gdb_script'], 'row': row}
    cap.write(folder / 'start.json', receipt)
    print('start GDB', row['kind'], flush=True)
    with (folder / 'stdout.txt').open('x') as out, (folder / 'stderr.txt').open('x') as err:
        result = subprocess.run(argv, cwd=CWD, env=env, stdout=out, stderr=err)
    receipt.update(exit_code=result.returncode, finished_utc=cap.now(), artifacts=[cap.desc(f) for f in sorted(folder.rglob('*')) if f.is_file()])
    cap.write(folder / 'command.json', receipt)
    assert result.returncode == 0
    assert cap.read(folder / 'dispatch.json')['status'] == 'pass'
    cap.write(folder / 'semantics.json', cap.validate_semantics(plan, row, folder))
    print('finish GDB', row['kind'], 'dispatch and semantic PASS', flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('freeze', 'capture'))
    parser.add_argument('--index', type=int, choices=(0, 1))
    args = parser.parse_args()
    if args.action == 'freeze':
        freeze()
    else:
        assert args.index is not None
        capture(args.index)
