#!/usr/bin/env python3
"""Separate live-mapping diagnostic; elapsed times are not adoption evidence."""
import datetime
import hashlib
import gzip
import json
import os
from pathlib import Path
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PREVIOUS = P.parent / 'change-0887'
sys.path.insert(0, str(PREVIOUS))
import experiment


def digest(path):
    h = hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def desc(path):
    return {'path': str(path.resolve()), 'bytes': path.stat().st_size, 'sha256': digest(path)}


def write(path, value):
    with path.open('x') as f:
        json.dump(value, f, indent=2, sort_keys=True)
        f.write('\n')


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    binaries = {role: PREVIOUS.parent / packet / 'binaries/native' for role, packet in [('before', 'change-0885'), ('after', 'change-0887')]}
    old_plan = json.loads((PREVIOUS / 'plan.json').read_text())
    for role, path in binaries.items():
        assert digest(path) == old_plan['binaries'][role]['native']['sha256']
    rows = [{'ordinal': block * 2 + offset, 'block': block, 'role': role} for block in range(4) for offset, role in enumerate(('before', 'after') if block % 2 == 0 else ('after', 'before'))]
    settings = ['/sys/kernel/mm/transparent_hugepage/enabled', '/sys/kernel/mm/transparent_hugepage/shmem_enabled']
    plan = {
        'schema': 'litchi.live-mapping-diagnostic-0888.v1', 'frozen_utc': now(),
        'source_head': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip(),
        'binaries': {role: desc(path) for role, path in binaries.items()},
        'helpers': [desc(Path(__file__)), desc(PREVIOUS / 'experiment.py'), desc(PREVIOUS.parent / 'change-0885/verify.py'), desc(PREVIOUS.parent / 'change-0885/capture.py')],
        'tools': [desc(Path('/usr/bin/time')), desc(Path('/usr/bin/taskset')), desc(Path(sys.executable).resolve())],
        'previous_plan': desc(PREVIOUS / 'plan.json'), 'rows': rows,
        'case': experiment.PRIMARY_CASE, 'warmup': 3, 'samples': 200,
        'observer_affinity': sorted(os.sched_getaffinity(0)), 'cpu': '12', 'child_cwd': str(PREVIOUS), 'environment': experiment.CHILD_ENV, 'poll_interval_seconds': 0.01,
        'kernel': {'uname': list(os.uname()), 'thp': {name: Path(name).read_text() for name in settings}},
        'scope': 'External GNU time process peak RSS plus sampled /proc status/smaps of its taskset/executable child. Polling is non-atomic and perturbs execution; sampled maxima can miss exit or transient peaks. No native latency or adoption claim.',
        'no_retries_or_replacements': True,
    }
    write(P / 'plan.json', plan)
    expected = experiment._load_expected_corpora()[experiment.PRIMARY_CASE]
    for row in rows:
        folder = P / 'runs' / f"{row['ordinal']:02d}-{row['role']}"
        folder.mkdir(parents=True, exist_ok=False)
        binary = binaries[row['role']].resolve()
        argv = experiment._command(str(binary), experiment.PRIMARY_CASE, 3, 200, folder / 'report.json', folder / 'producer.json', folder / 'peak-rss-kib.txt')
        started = now()
        observations = []
        misses = []
        with (folder / 'stdout.txt').open('xb') as out, (folder / 'stderr.txt').open('xb') as err, gzip.open(folder / 'snapshots.jsonl.gz', 'xt') as snapshots:
            child = subprocess.Popen(argv, cwd=PREVIOUS, env=experiment.CHILD_ENV, stdout=out, stderr=err)
            while child.poll() is None:
                try:
                    ids = Path(f'/proc/{child.pid}/task/{child.pid}/children').read_text().split()
                    if len(ids) > 1:
                        raise ValueError('unexpected multiple GNU time children')
                    if ids:
                        pid = int(ids[0])
                        base = Path('/proc') / str(pid)
                        executable = os.readlink(base / 'exe')
                        # Taskset execs the benchmark in the same PID. Keep
                        # only snapshots whose executable identity is native.
                        if executable == str(binary):
                            status = (base / 'status').read_text()
                            smaps = (base / 'smaps').read_text()
                            index = len(observations)
                            metadata = {'index': index, 'monotonic_ns': time.monotonic_ns(), 'pid': pid, 'exe': executable}
                            snapshots.write(json.dumps({**metadata, 'status': status, 'smaps': smaps}) + '\n')
                            observations.append(metadata)
                except (FileNotFoundError, ProcessLookupError) as error:
                    misses.append({'monotonic_ns': time.monotonic_ns(), 'reason': type(error).__name__})
                time.sleep(0.01)
            code = child.wait()
        write(folder / 'command.json', {'row': row, 'argv': argv, 'cwd': str(PREVIOUS), 'environment': experiment.CHILD_ENV, 'started_utc': started, 'finished_utc': now(), 'exit_code': code, 'launcher_pid': child.pid, 'plan_sha256': digest(P / 'plan.json')})
        write(folder / 'observations.json', {'snapshots': observations, 'proc_races': misses})
        if code:
            raise RuntimeError(f'child {row} failed; no retry')
        report = json.loads((folder / 'report.json').read_text())
        producer = json.loads((folder / 'producer.json').read_text())
        identity = experiment._binary(old_plan, row['role'], 'native')
        experiment._validate_child(report, producer, experiment.PRIMARY_CASE, 'native', 3, 200, expected, identity)
        assert observations, 'no native snapshots captured'
        write(folder / 'receipt.json', {'status': 'pass', 'row': row, 'peak_rss_kib': experiment._read_peak(folder / 'peak-rss-kib.txt'), 'snapshots': len(observations), 'files': {f.name: desc(f) for f in sorted(folder.iterdir()) if f.is_file()}})
        print(row, 'snapshots', len(observations), flush=True)
    write(P / 'capture.json', {'status': 'pass', 'children': len(rows), 'measured_iterations': 1600, 'warmups': 24, 'plan': desc(P / 'plan.json'), 'latency_claim': None})


if __name__ == '__main__':
    main()
