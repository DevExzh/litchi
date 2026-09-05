#!/usr/bin/env python3
"""Run every declared CLI lane and preserve failure/create-new controls."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--binary', type=Path, required=True)
    parser.add_argument('--directory', required=True)
    args = parser.parse_args()
    target = ROOT / args.directory
    assert target.parent == ROOT
    target.mkdir(exist_ok=False)
    revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO, text=True).strip()
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    receipts = []
    for lane in protocol['lanes']:
        name = lane['command'] + '-' + lane['selector'] + '-' + lane['provider_label']
        report = target / (name + '.json')
        argv = [str(args.binary.resolve()), lane['command'], '--' + lane['selector_key'], lane['selector'],
                '--provider', lane['provider'], '--samples', '1', '--warmup', '0',
                '--source-revision', revision, '--output', str(report)]
        if lane['provider'] == 'range':
            argv += ['--max-range', str(lane['max_range']), '--delay-us', str(lane['delay_us'])]
        with (target / (name + '.log')).open('xb') as output:
            result = subprocess.run(argv, cwd=REPO, stdout=output, stderr=subprocess.STDOUT)
        receipt = {'name': name, 'argv': argv, 'exit_code': result.returncode}
        (target / (name + '-receipt.json')).write_text(json.dumps(receipt, indent=2) + '\n')
        assert result.returncode == 0, name
        subprocess.run([sys.executable, '-B', str(ROOT / 'verify-report.py'), str(report)], check=True)
        subprocess.run([sys.executable, '-B', str(ROOT / 'probe-report.py'), str(report)], check=True)
        before = hashlib.sha256(report.read_bytes()).hexdigest()
        with (target / (name + '-create-new.log')).open('xb') as output:
            duplicate = subprocess.run(argv, cwd=REPO, stdout=output, stderr=subprocess.STDOUT)
        assert duplicate.returncode != 0 and hashlib.sha256(report.read_bytes()).hexdigest() == before
        receipt.update(status='pass', output_sha256=before, existing_output_preserved=True,
                       existing_output_exit_code=duplicate.returncode)
        (target / (name + '-receipt.json')).write_text(json.dumps(receipt, indent=2) + '\n')
        receipts.append(receipt)
        print('SMOKE ' + name + ' pass', flush=True)
    (target / 'index.json').write_text(json.dumps(receipts, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'lanes': len(receipts)}))


if __name__ == '__main__':
    main()
