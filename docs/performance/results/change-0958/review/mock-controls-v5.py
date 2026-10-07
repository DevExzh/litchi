#!/usr/bin/env python3
"""Offline custody controls for capture_0958_v5.py; never launches a workload."""
import contextlib
import importlib.util
import io
import json
import shutil
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True
RUNNER = Path('/home/zhuhe/litchi-performance-0958/tooling/capture_0958_v5.py')
PLAN = Path('/home/zhuhe/litchi-performance-0958/plan-v5.json')
spec = importlib.util.spec_from_file_location('capture0958_v5_controls', RUNNER)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

plan, validation = mod.validate_plan()
assert validation['workload_not_started'] is True

jobs = [
    {'name': f'job-{index}', 'shape': 'medium', 'workload': 'one-edit',
     'workers': 1, 'warmup': 0, 'samples': 3}
    for index in range(6)
]
failstop = []
for fail_at in range(6):
    root = Path(tempfile.mkdtemp(prefix=f'litchi-0958-v4-failstop-{fail_at}-'))
    calls = []
    old_run_child = mod.run_child
    old_persist = mod.persist_report_oracle
    old_admission = mod.admission_print
    def fake_run_child(plan, validation, job, capture, *, _fail_at=fail_at, _calls=calls):
        _calls.append(job['name'])
        return {
            'id': job['name'],
            'status': 'fail' if len(_calls) - 1 == _fail_at else 'pass',
            'artifacts': {'profile_actual': {'path': '/tmp/mock-profile'}},
        }
    try:
        mod.run_child = fake_run_child
        mod.persist_report_oracle = lambda *args, **kwargs: {'status': 'pass'}
        mod.admission_print = lambda *args, **kwargs: {'status': 'pass'}
        mock_plan = {
            'capture_root': str(root / 'capture'),
            'admission_root': str(root / 'admission'),
            'matrix': {'jobs': jobs},
        }
        mock_validation = {'plan': {'path': 'mock-plan'}, 'runner': {'path': 'mock-runner'}}
        with contextlib.redirect_stdout(io.StringIO()):
            rc = mod.run_capture(mock_plan, mock_validation)
        expected = [job['name'] for job in jobs[:fail_at + 1]]
        assert rc == 1 and calls == expected, (fail_at, rc, calls, expected)
        outer = json.loads((root / 'capture' / 'runner-terminal.json').read_text())
        assert outer['child_terminals_written'] == fail_at + 1
        failstop.append({'failed_index': fail_at, 'launched': calls})
    finally:
        mod.run_child = old_run_child
        mod.persist_report_oracle = old_persist
        mod.admission_print = old_admission
        shutil.rmtree(root)

# A KeyboardInterrupt from wait() must route through stop_group.
root = Path(tempfile.mkdtemp(prefix='litchi-0958-v4-interrupt-'))
stop_calls = []
class InterruptProc:
    pid = 424242
    returncode = -2
    def wait(self, timeout=None):
        raise KeyboardInterrupt()
old_popen = mod.subprocess.Popen
old_stop = mod.stop_group
try:
    mod.subprocess.Popen = lambda *args, **kwargs: InterruptProc()
    def stop_spy(pgid, proc):
        stop_calls.append((pgid, proc.pid))
        return {'actions': ['mock-stop'], 'process_group_clear': True, 'reaped': True}
    mod.stop_group = stop_spy
    result = mod.run_process(['mock-child'], root, mod.FIXED_ENV,
                             root / 'stdout.log', root / 'stderr.log', 1)
    assert stop_calls == [(424242, 424242)]
    assert result['cleanup']['reaped'] is True
    assert 'KeyboardInterrupt' in result['launch_error']
finally:
    mod.subprocess.Popen = old_popen
    mod.stop_group = old_stop
    shutil.rmtree(root)

# A zero return code must still be rejected when launch_error is present.
root = Path(tempfile.mkdtemp(prefix='litchi-0958-v4-error-zero-'))
old_run_process = mod.run_process
old_child_artifacts = mod.child_artifacts
try:
    source = root / 'source'
    source.mkdir()
    profile = root / 'profile.data'
    profile.write_bytes(b'profile')
    error_run = {
        'returncode': 0,
        'timed_out': False,
        'launch_error': 'KeyboardInterrupt: interrupted wait',
        'cleanup': {'process_group_clear': True, 'reaped': True},
        'pid': 123,
        'process_group_id': 123,
    }
    def error_process(argv, cwd, env, stdout_path, stderr_path, deadline):
        stdout_path.write_text('litchi_perf_baseline::xlsx_managed_compression::run_case\\n')
        stderr_path.write_text('')
        return dict(error_run)
    mod.run_process = error_process
    mod.child_artifacts = lambda folder: {
        'start': 'start.json', 'terminal': 'terminal.json', 'stdout': 'stdout.log',
        'stderr': 'stderr.log', 'time': 'time.txt', 'workload_dir': 'workload',
        'report': 'workload/report.json', 'profile_requested': 'heaptrack.data.gz',
        'profile_candidates': ['heaptrack.data.gz'],
        'profile_actual': {'path': str(profile), 'bytes': profile.stat().st_size,
                           'mode': 420, 'executable': False, 'sha256': 'x'},
        'report_descriptor': {'path': str(root / 'report.json'), 'bytes': 1,
                              'mode': 420, 'executable': False, 'sha256': 'y'},
    }
    tools = {name: {'path': f'/tmp/{name}'}
             for name in ('time', 'taskset', 'heaptrack', 'heaptrack_print')}
    mock_plan = {
        'capture_root': str(root / 'capture'), 'admission_root': str(root / 'admission'),
        'tools': tools, 'binary': {'binary': {'path': '/tmp/fake-binary'}},
        'first_child_admission': {'deadline_seconds': 1},
    }
    mock_validation = {
        'plan': {'path': 'plan'}, 'runner': {'path': 'runner'},
        'source': {'root': str(source)},
        'binary': {'binary': {'path': '/tmp/fake-binary'}},
    }
    job = {'name': 'job-0', 'shape': 'medium', 'workload': 'one-percent'}
    child = mod.run_child(mock_plan, mock_validation, job, root / 'capture')
    admission = mod.admission_print(
        mock_plan, mock_validation, job,
        root / 'capture' / 'children' / 'job-0', profile, {'status': 'pass'},
    )
    assert child['status'] == 'fail', child
    assert admission['status'] == 'fail', admission
    error_zero = {
        'child_status': child['status'],
        'admission_status': admission['status'],
        'launch_error': error_run['launch_error'],
    }
finally:
    mod.run_process = old_run_process
    mod.child_artifacts = old_child_artifacts
    shutil.rmtree(root)

# The actual plan must reject a mismatched executing __file__.
original_file = mod.__file__
try:
    mod.__file__ = '/tmp/litchi-0958-wrong-runner.py'
    try:
        mod.validate_plan()
    except mod.CaptureError as exc:
        wrong_runner = str(exc)
        assert 'executing runner path differs from plan.runner' in wrong_runner
    else:
        raise AssertionError('wrong executing runner was accepted')
finally:
    mod.__file__ = original_file

print(json.dumps({
    'status': 'pass',
    'plan': str(PLAN),
    'plan_sha256': '5bd826e1f93913b5be2fc5878be0c5270e323789222fffcf0e97121dcbc37695',
    'runner': str(RUNNER),
    'runner_sha256': '15c09dffefed8bf6084ec0c54ddf7d19f2f69a3c70f4e2bf41a7442a40000f6c',
    'plan_validation': 'pass_without_workload',
    'failstop': failstop,
    'interrupt_cleanup': {'status': 'pass', 'stop_group_calls': stop_calls},
    'zero_return_launch_error': error_zero,
    'wrong_runner': {'status': 'pass', 'error': wrong_runner},
}, sort_keys=True))
