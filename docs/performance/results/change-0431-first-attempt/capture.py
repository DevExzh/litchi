#!/usr/bin/env python3
"""Capture one frozen sixteen-process role, serially, with source custody."""
import argparse
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--role', choices=['before', 'after'], required=True)
    args = parser.parse_args()
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    build = json.loads((ROOT / ('build-' + args.role + '.json')).read_text())
    binary = Path(build['capture_binary'])
    assert sha(binary) == build['binary_sha256']
    assert sha(ROOT / 'protocol.json') == build['protocol_sha256']
    assert sha(ROOT / 'verify-report.py') == build['verifier_sha256']
    assert protocol['cpu'] in os.sched_getaffinity(0)
    spec = importlib.util.spec_from_file_location('custody', ROOT / 'check.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    assert custody.sources() == build['source_manifest']
    target = ROOT / args.role
    target.mkdir(exist_ok=False)
    index = []
    for lane in protocol['order']:
        name = lane['selector'] + '-' + lane['provider_label'] + '-' + lane['repeat'].lower()
        report = target / (name + '.json')
        log = target / (name + '.log')
        resource = target / (name + '-resource.log')
        receipt_path = target / (name + '-receipt.json')
        argv = ['taskset', '-c', str(protocol['cpu']), '/usr/bin/time', '-v', '-o', str(resource),
                str(binary), 'provider-lifecycle', '--corpus', lane['selector'],
                '--provider', lane['provider'], '--samples', str(protocol['samples_per_process']),
                '--warmup', str(protocol['warmups_per_process']), '--source-revision', build['revision'],
                '--output', str(report)]
        if lane['provider'] == 'range':
            argv += ['--max-range', str(lane['max_range']), '--delay-us', str(lane['delay_us'])]
        row = {'role': args.role, 'name': name, 'lane': lane, 'argv': argv,
               'source_manifest': build['source_manifest'], 'revision': build['revision'],
               'binary_sha256': build['binary_sha256'], 'protocol_sha256': build['protocol_sha256'],
               'driver_sha256': sha(Path(__file__)), 'verifier_sha256': build['verifier_sha256'],
               'started_utc': now(), 'status': 'running'}
        receipt_path.write_text(json.dumps(row, indent=2) + '\n')
        print('START ' + args.role + ' ' + name, flush=True)
        try:
            with log.open('xb') as stream:
                result = subprocess.run(argv, cwd=REPO, stdout=stream, stderr=subprocess.STDOUT)
            row['exit_code'] = result.returncode
            assert result.returncode == 0
            subprocess.run([sys.executable, '-B', str(ROOT / 'verify-report.py'), str(report)], check=True)
            row['status'] = 'pass'
        finally:
            if row['status'] == 'running':
                row['status'] = 'failed'
            row['finished_utc'] = now()
            row['artifacts'] = {str(p.relative_to(ROOT)): {'sha256': sha(p), 'bytes': p.stat().st_size}
                                for p in [report, log, resource] if p.is_file()}
            receipt_path.write_text(json.dumps(row, indent=2) + '\n')
        index.append(str(receipt_path.relative_to(ROOT)))
        print('FINISH ' + name, flush=True)
    assert len(index) == protocol['expected_processes_per_role']
    assert custody.sources() == build['source_manifest']
    (ROOT / (args.role + '-index.json')).write_text(json.dumps(index, indent=2) + '\n')


if __name__ == '__main__':
    main()
