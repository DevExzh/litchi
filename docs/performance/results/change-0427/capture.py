#!/usr/bin/env python3
"""Capture the frozen eight-process retention matrix serially on one CPU."""
import argparse
import datetime
import hashlib
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
    parser.add_argument('--binary', type=Path, required=True)
    parser.add_argument('--revision', required=True)
    args = parser.parse_args()
    binary = args.binary.resolve()
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO).decode().strip() == args.revision
    subprocess.run(['git', 'diff', '--exit-code', args.revision, '--', '*.rs', '*.toml', '*.lock'], cwd=REPO, check=True)
    assert protocol['cpu'] in os.sched_getaffinity(0)
    build = json.loads((ROOT / 'build.json').read_text())
    assert build['revision'] == args.revision
    assert sha(binary) == build['binary_sha256']
    assert sha(ROOT / 'protocol.json') == build['protocol_sha256']
    output = ROOT / 'capture'
    output.mkdir(exist_ok=False)
    records = []
    for corpus in protocol['corpora']:
        for role in protocol['order_per_corpus']:
            name = corpus + '-' + role['api'] + '-' + role['repeat'].lower()
            report = output / (name + '.json')
            log = output / (name + '.log')
            resource = output / (name + '-resource.log')
            receipt = output / (name + '-receipt.json')
            argv = ['taskset', '-c', str(protocol['cpu']), '/usr/bin/time', '-v', '-o', str(resource),
                    str(binary), 'retention', '--api', role['api'], '--corpus', corpus,
                    '--samples', str(protocol['samples_per_process']),
                    '--warmup', str(protocol['warmups_per_process']),
                    '--source-revision', args.revision, '--output', str(report)]
            record = {'name': name, 'corpus': corpus, **role, 'argv': argv,
                      'revision': args.revision, 'binary_sha256': sha(binary),
                      'protocol_sha256': sha(ROOT / 'protocol.json'),
                      'started_utc': now(), 'status': 'running'}
            receipt.write_text(json.dumps(record, indent=2) + '\n')
            print('START ' + name, flush=True)
            with log.open('wb') as stream:
                result = subprocess.run(argv, cwd=REPO, stdout=stream, stderr=subprocess.STDOUT)
            record.update(exit_code=result.returncode, finished_utc=now(),
                          status='pass' if result.returncode == 0 else 'failed')
            record['artifacts'] = {str(p.relative_to(ROOT)): {'sha256': sha(p), 'bytes': p.stat().st_size}
                                   for p in [report, log, resource] if p.is_file()}
            receipt.write_text(json.dumps(record, indent=2) + '\n')
            print('FINISH ' + name + ' ' + record['status'], flush=True)
            if result.returncode:
                return result.returncode
            subprocess.run([sys.executable, '-B', str(ROOT / 'verify-report.py'), str(report)], check=True)
            records.append(str(receipt.relative_to(ROOT)))
    assert len(records) == protocol['expected_processes']
    (ROOT / 'capture-index.json').write_text(json.dumps(records, indent=2) + '\n')
    return 0


if __name__ == '__main__':
    sys.exit(main())
