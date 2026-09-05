#!/usr/bin/env python3
"""Capture supplementary CPU stacks; retain setup and observer scope explicitly."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--binary', type=Path, required=True)
    args = parser.parse_args()
    build = json.loads((ROOT / 'build.json').read_text())
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert digest(args.binary) == build['binary_sha256']
    assert digest(ROOT / 'profiles.py') == build['additional_artifact_sha256']['profiles.py']
    assert digest(ROOT / 'protocol.json') == build['protocol_sha256']
    plan = protocol['supplementary_profiles']
    target = ROOT / 'profiles'
    target.mkdir(exist_ok=False)
    receipts = []
    for case in plan['cases']:
        name = case['corpus'] + '-' + case['provider']
        data = target / (name + '.data')
        report = target / (name + '.json')
        argv = ['taskset', '-c', str(protocol['cpu']), 'perf', 'record', '-e', plan['event'],
                '-F', str(plan['frequency_hz']), '--call-graph', plan['call_graph'],
                '-o', str(data), '--', str(args.binary), case['command'], '--corpus', case['corpus'],
                '--provider', case['provider'], '--samples', str(plan['samples']),
                '--warmup', str(plan['warmups']), '--source-revision', build['revision'],
                '--output', str(report)]
        receipt = {'name': name, 'argv': argv, 'revision': build['revision'],
                   'binary_sha256': build['binary_sha256'], 'scope': plan['scope'],
                   'started_utc': datetime.datetime.now(datetime.timezone.utc).isoformat()}
        print('PROFILE ' + name, flush=True)
        with (target / (name + '.log')).open('xb') as output:
            result = subprocess.run(argv, cwd=REPO, stdout=output, stderr=subprocess.STDOUT)
        receipt['exit_code'] = result.returncode
        receipt['status'] = 'pass' if result.returncode == 0 else 'failed'
        receipt['finished_utc'] = datetime.datetime.now(datetime.timezone.utc).isoformat()
        if result.returncode == 0:
            subprocess.run([sys.executable, '-B', str(ROOT / 'verify-report.py'), str(report)], check=True)
            for suffix, command in [('script.log', ['perf', 'script', '-i', str(data)]),
                                    ('report.log', ['perf', 'report', '--stdio', '--no-children', '-i', str(data)])]:
                with (target / (name + '-' + suffix)).open('xb') as output:
                    subprocess.run(command, cwd=REPO, stdout=output, stderr=subprocess.STDOUT, check=True)
        receipt['artifacts'] = {str(p.relative_to(ROOT)): {'sha256': digest(p), 'bytes': p.stat().st_size}
                                for p in sorted(target.glob(name + '*')) if p.is_file()}
        receipt_path = target / (name + '-receipt.json')
        receipt_path.write_text(json.dumps(receipt, indent=2) + '\n')
        receipts.append(str(receipt_path.relative_to(ROOT)))
        if result.returncode:
            return result.returncode
    (ROOT / 'profile-index.json').write_text(json.dumps(receipts, indent=2) + '\n')
    return 0


if __name__ == '__main__':
    sys.exit(main())
