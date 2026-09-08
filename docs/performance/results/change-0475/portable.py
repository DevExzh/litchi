#!/usr/bin/env python3
"""Verify and replay a fresh standalone copy after owned runtime cleanup."""
import datetime
import fcntl
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parent


def main():
    assert not Path('/tmp/litchi-goal-0474').exists()
    assert not Path('/tmp/litchi-goal-0475').exists()
    subprocess.run(['python3', '-B', 'seal.py'], cwd=ROOT, check=True)
    seal_hash = hashlib.sha256((ROOT / 'SHA256SUMS').read_bytes()).hexdigest()
    temporary = Path(tempfile.mkdtemp(prefix='litchi-goal-0475-portable-', dir='/tmp'))
    results = []
    completed = False
    try:
        with (temporary / 'cpu.lock').open('w') as lock:
            fcntl.flock(lock, fcntl.LOCK_EX)
            clone = temporary / 'bundle'
            shutil.copytree(ROOT, clone)

            def run(argv, expected=0):
                started = datetime.datetime.now(datetime.timezone.utc).isoformat()
                process = subprocess.run(argv, cwd=clone, text=True, capture_output=True,
                    env=dict(os.environ, PYTHONDONTWRITEBYTECODE='1'))
                results.append(dict(argv=argv, cwd=str(clone), started_utc=started,
                    finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                    exit_code=process.returncode, expected_exit_code=expected,
                    stdout=process.stdout, stderr=process.stderr))
                print(argv[2:], process.returncode, flush=True)
                assert process.returncode == expected, process.stderr

            run(['python3', '-B', 'verify.py', '--sealed'])
            run(['python3', '-B', 'replay.py'])
            run(['python3', '-B', '-m', 'unittest', 'discover', '-s', '.', '-p', 'test_*.py', '-v'])
            target = clone / 'owners.json'
            original = target.read_bytes()
            mutated = json.loads(original)
            mutated['lanes'][0]['combined_requested_bytes'] += 1
            target.write_text(json.dumps(mutated))
            run(['python3', '-B', 'replay.py'], expected=1)
            assert 'attribution does not replay exactly' in results[-1]['stderr']
            target.write_bytes(original)
            assert hashlib.sha256((clone / 'SHA256SUMS').read_bytes()).hexdigest() == seal_hash
            completed = True
    finally:
        shutil.rmtree(temporary)
        (ROOT / 'portable-replay.json').write_text(json.dumps(dict(
            schema='litchi-0475-portable-replay-v1', verified_seal_sha256=seal_hash,
            original_runtime_roots_absent=True, copied_root=str(temporary),
            copied_root_removed=not temporary.exists(), commands=results,
            completed=completed), indent=2, sort_keys=True)+'\n')


if __name__ == '__main__':
    main()
