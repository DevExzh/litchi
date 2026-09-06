#!/usr/bin/env python3
"""Run one serialized command with immutable, deduplicated source custody."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def sources(workspace_only=False):
    names = set(subprocess.check_output(
        ['git', 'ls-files', '--cached', '--others', '--exclude-standard', '-z'],
        cwd=REPO).decode().split('\0'))
    names.add('Cargo.lock')
    rows = {name: sha((REPO / name).read_bytes()) for name in sorted(names)
            if name.endswith(('.rs', '.toml', '.lock')) and (REPO / name).is_file()
            and (not workspace_only or not name.startswith('tools/'))}
    raw = (json.dumps(rows, sort_keys=True, indent=2) + '\n').encode()
    digest = sha(raw)
    target = ROOT / 'sources' / (digest + '.json')
    target.parent.mkdir(exist_ok=True)
    if target.exists():
        assert target.read_bytes() == raw
    else:
        target.write_bytes(raw)
    return {'path': str(target.relative_to(ROOT)), 'sha256': digest, 'files': len(rows)}


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tag', required=True)
    parser.add_argument('--workspace-only', action='store_true',
                        help='Scope source custody to workspace sources, excluding standalone tools')
    parser.add_argument('command', nargs=argparse.REMAINDER)
    args = parser.parse_args()
    assert re.fullmatch(r'[a-z0-9_-]+', args.tag)
    argv = args.command[1:] if args.command[:1] == ['--'] else args.command
    assert argv
    directory = ROOT / 'checks'
    directory.mkdir(exist_ok=True)
    receipt = directory / (args.tag + '.json')
    log = directory / (args.tag + '.log')
    assert not receipt.exists() and not log.exists() and not log.with_suffix('.log.gz').exists()
    env = {'RUSTUP_TOOLCHAIN': '1.98.1', 'CARGO_BUILD_JOBS': '4',
           'CARGO_INCREMENTAL': '0', 'PYTHONDONTWRITEBYTECODE': '1'}
    if 'doc' in argv:
        env['RUSTDOCFLAGS'] = '-D warnings'
    record = {'change': 447, 'argv': argv, 'cwd': str(REPO),
              'revision': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO).decode().strip(),
              'driver_sha256': sha(Path(__file__).read_bytes()), 'environment': env,
              'source_scope': 'workspace excluding tools' if args.workspace_only else 'workspace and standalone tools',
              'started_utc': now(), 'source_before': sources(args.workspace_only), 'status': 'running'}
    receipt.write_text(json.dumps(record, indent=2) + '\n')
    with log.open('wb') as output:
        result = subprocess.run(argv, cwd=REPO, env=os.environ | env,
                                stdout=output, stderr=subprocess.STDOUT)
    record.update(exit_code=result.returncode, finished_utc=now(), source_after=sources(args.workspace_only))
    record['source_unchanged'] = record['source_before'] == record['source_after']
    raw = log.read_bytes()
    record['log'] = {'path': str(log.relative_to(ROOT)), 'bytes': len(raw), 'sha256': sha(raw)}
    if 'test' in argv:
        rows = re.findall(rb'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;', raw)
        for index, key in enumerate(['passed_tests', 'failed_tests', 'ignored_tests']):
            record[key] = sum(int(row[index]) for row in rows)
    passed = result.returncode == 0 and record['source_unchanged'] and record.get('passed_tests', 1) > 0
    record['status'] = 'pass' if passed else 'failed'
    receipt.write_text(json.dumps(record, indent=2) + '\n')
    print(json.dumps({key: record[key] for key in ['status', 'exit_code', 'source_unchanged']}))
    return 0 if passed else 1


if __name__ == '__main__':
    sys.exit(main())
