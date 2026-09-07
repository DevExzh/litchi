#!/usr/bin/env python3
"""Capture file-backed ODP appends and independent native-fixture oracles."""
import datetime
import gzip
import hashlib
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[4]
TASK = Path('/tmp/litchi-goal-0457/native')
TITLE = 'Native title & <字> café'
BODY = 'Native body Ω 😀'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path, value):
    with path.open('x') as output:
        json.dump(value, output, indent=2, ensure_ascii=False)
        output.write('\n')


def retain_archive(source, target):
    raw = source.read_bytes()
    with target.open('xb') as output:
        output.write(gzip.compress(raw, mtime=0))
    return {'path': target.name, 'bytes': target.stat().st_size, 'sha256': sha(target),
            'decoded_bytes': len(raw), 'decoded_sha256': hashlib.sha256(raw).hexdigest()}


def main():
    binding_path = ROOT / 'binary-binding.json'
    binding = json.loads(binding_path.read_text())
    binary = Path(binding['binary']['path'])
    assert sha(binary) == binding['binary']['sha256']
    inventory_path = ROOT / 'staticinventory.json'
    assert sha(inventory_path) == binding['inventory_sha256']
    assert sha(ROOT / 'verify-output.py') == binding['oracle_sha256']
    assert sha(Path(__file__)) == binding['runner_sha256']
    inventory = json.loads(inventory_path.read_text())
    TASK.mkdir(parents=True, exist_ok=False)
    destination = ROOT / 'runs'
    destination.mkdir(exist_ok=False)
    records = []
    for index, fixture in enumerate(inventory['fixtures']):
        source = REPO / fixture['path']
        assert sha(source) == fixture['archive_sha256']
        output = TASK / f'{index:02d}.odp'
        argv = [str(binary), str(source), str(output), TITLE, BODY]
        started = datetime.datetime.now(datetime.timezone.utc).isoformat()
        result = subprocess.run(argv, cwd=REPO, capture_output=True)
        log = destination / f'{index:02d}-probe.log'
        log.write_bytes(result.stdout + result.stderr)
        record = {'fixture': fixture['path'], 'source_sha256': sha(source),
                  'argv': argv, 'cwd': str(REPO), 'started_utc': started,
                  'exit_code': result.returncode, 'log': log.name,
                  'log_sha256': sha(log), 'title': TITLE, 'body': BODY}
        record['retained_source'] = retain_archive(source, destination / f'{index:02d}-source.odp.gz')
        if output.exists():
            record['output'] = {'path': str(output), 'bytes': output.stat().st_size,
                                'sha256': sha(output)}
            record['retained_output'] = retain_archive(output, destination / f'{index:02d}-output.odp.gz')
        if result.returncode == 0:
            oracle_argv = ['python3', '-B', str(ROOT / 'verify-output.py'),
                           '--source', str(source), '--output', str(output),
                           '--title', TITLE, '--body', BODY]
            oracle = subprocess.run(oracle_argv, cwd=REPO, capture_output=True)
            oracle_log = destination / f'{index:02d}-oracle.log'
            oracle_log.write_bytes(oracle.stdout + oracle.stderr)
            record['oracle'] = {'argv': oracle_argv, 'exit_code': oracle.returncode,
                                'log': oracle_log.name, 'sha256': sha(oracle_log)}
            record['status'] = 'validated' if oracle.returncode == 0 else 'oracle-failed'
        else:
            record['status'] = 'probe-failed'
        assert sha(source) == fixture['archive_sha256']
        write_json(destination / f'{index:02d}-receipt.json', record)
        records.append(record)
        print(json.dumps({'fixture': fixture['path'], 'status': record['status']}), flush=True)
    write_json(destination / 'index.json', {'binding_sha256': sha(binding_path), 'records': records})
    return 0 if all(record['status'] == 'validated' for record in records) else 1


if __name__ == '__main__':
    raise SystemExit(main())
