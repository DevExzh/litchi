#!/usr/bin/env python3
"""Record two frame-pointer CPU attribution controls on the unchanged binary."""
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import platform
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    p = json.loads((ROOT / 'profile-protocol.json').read_text())
    b = json.loads((ROOT / 'input-build-0429.json').read_text())
    spec = importlib.util.spec_from_file_location('custody', ROOT / 'check.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    source = custody.sources()
    assert source['sha256'] == b['source_manifest']['sha256']
    assert digest(Path(p['capture_binary'])) == p['binary_sha256'] == b['binary_sha256']
    assert p['cpu'] in os.sched_getaffinity(0)
    machine = {'captured_utc': now(), 'platform': platform.platform(), 'uname': list(os.uname()),
               'affinity': sorted(os.sched_getaffinity(0)), 'python': sys.version,
               'prior_machine_sha256': digest(ROOT / 'machine-prior.json'),
               'scope': 'Same workspace host; hardware/toolchain details retained from the previous machine record.'}
    (ROOT / 'machine-current.json').write_text(json.dumps(machine, indent=2) + '\n')
    target = ROOT / 'profiles'
    target.mkdir(exist_ok=False)
    receipts = []
    for provider in p['providers']:
        data, report, log = (target / (provider + suffix) for suffix in ['.data', '.json', '-record.log'])
        argv = ['taskset', '-c', str(p['cpu']), 'perf', 'record', '-e', p['event'], '-F', str(p['frequency_hz']),
                '--call-graph', p['call_graph'], '-o', str(data), '--', p['capture_binary'], 'provider-lifecycle',
                '--corpus', p['corpus'], '--provider', provider, '--samples', str(p['samples']),
                '--warmup', str(p['warmup']), '--source-revision', p['source_revision'], '--output', str(report)]
        receipt = {'provider': provider, 'argv': argv, 'source': source, 'binary_sha256': p['binary_sha256'],
                   'protocol_sha256': digest(ROOT / 'profile-protocol.json'), 'driver_sha256': digest(Path(__file__)),
                   'verifier_sha256': digest(ROOT / 'verify-report.py'), 'started_utc': now(), 'status': 'running'}
        path = target / (provider + '-receipt.json')
        path.write_text(json.dumps(receipt, indent=2) + '\n')
        print('PROFILE ' + provider, flush=True)
        try:
            with log.open('xb') as out:
                result = subprocess.run(argv, cwd=REPO, stdout=out, stderr=subprocess.STDOUT)
            receipt['record_exit_code'] = result.returncode
            assert result.returncode == 0
            subprocess.run([sys.executable, '-B', str(ROOT / 'verify-report.py'), str(report)], check=True)
            for suffix, command in [('script.log', ['perf', 'script', '-i', str(data)]),
                                    ('report.log', ['perf', 'report', '--stdio', '--no-children', '-i', str(data)])]:
                with (target / (provider + '-' + suffix)).open('xb') as out:
                    subprocess.run(command, cwd=REPO, stdout=out, stderr=subprocess.STDOUT, check=True)
            receipt['status'] = 'pass'
        finally:
            if receipt['status'] == 'running':
                receipt['status'] = 'failed'
            receipt['finished_utc'] = now()
            receipt['artifacts'] = {str(item.relative_to(ROOT)): {'sha256': digest(item), 'bytes': item.stat().st_size}
                                     for item in sorted(target.glob(provider + '*')) if item.is_file() and item != path}
            path.write_text(json.dumps(receipt, indent=2) + '\n')
        receipts.append(str(path.relative_to(ROOT)))
    assert custody.sources() == source
    (ROOT / 'profile-index.json').write_text(json.dumps(receipts, indent=2) + '\n')


if __name__ == '__main__':
    main()
