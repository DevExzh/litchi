#!/usr/bin/env python3
"""Capture serial A1/B1/B2/A2 ODG open controls with the unchanged 0502 probe."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess

CORPORA = ('plain-small', 'plain-large', 'metadata-small', 'metadata-large')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--before', required=True, type=Path)
    parser.add_argument('--after', required=True, type=Path)
    parser.add_argument('--output', type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    binaries = {phase: path.resolve() for phase, path in [('before', args.before), ('after', args.after)]}
    identities = {phase: hashlib.sha256(path.read_bytes()).hexdigest() for phase, path in binaries.items()}
    schedule = [('before', 'r1'), ('after', 'r1'), ('after', 'r2'), ('before', 'r2')]
    receipts = []
    for phase, repeat in schedule:
        target = args.output / phase
        target.mkdir(parents=True, exist_ok=True)
        for corpus in CORPORA:
            output = target / f'{corpus}-{repeat}.json'
            argv = ['taskset', '-c', '2', '/usr/bin/time', '-v', '-o', str(output.with_suffix('.time.txt')),
                    str(binaries[phase]), '--corpus', corpus, '--warmups', '25', '--samples', '200', '--output', str(output)]
            subprocess.run(argv, check=True, env={**os.environ, 'LITCHI_GIT_REV': f'0504-{phase}-see-source-manifest'})
            report = json.loads(output.read_text())
            assert report['binary_sha256'] == identities[phase]
            assert report['samples'] == len(report['elapsed_ns']) == 200
            receipts.append({'phase': phase, 'repeat': repeat, 'corpus': corpus, 'argv': argv,
                             'report_sha256': hashlib.sha256(output.read_bytes()).hexdigest()})
    for corpus in CORPORA:
        reports = [json.loads((args.output / phase / f'{corpus}-{repeat}.json').read_text()) for phase, repeat in schedule]
        for field in ('input_sha256', 'input_bytes', 'semantic_checksum', 'pages', 'shapes_per_page', 'rich_metadata'):
            assert all(report[field] == reports[0][field] for report in reports), (corpus, field)
    for phase, path in binaries.items():
        assert hashlib.sha256(path.read_bytes()).hexdigest() == identities[phase]
    (args.output / 'capture.json').write_text(json.dumps({'schedule': schedule, 'binary_sha256': identities,
        'samples_per_child': 200, 'warmups_per_child': 25, 'cpu': 2, 'receipts': receipts,
        'scope': 'owned-byte open plus semantic traversal; whole-child RSS includes setup/report work'}, indent=2) + '\n')


if __name__ == '__main__':
    main()
