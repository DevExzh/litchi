#!/usr/bin/env python3
"""Run one serialized source-bound verification command and retain its output."""
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


def sources():
    names = set(subprocess.check_output(['git', 'ls-files', '-z'], cwd=REPO).decode().split('\0'))
    # The workspace lockfile is ignored; bind it as well as tracked sources.
    names.add('Cargo.lock')
    return {name: hashlib.sha256((REPO / name).read_bytes()).hexdigest()
            for name in names if name and (name.endswith(('.rs', '.toml', '.lock')))
            and (REPO / name).is_file()}


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    global REPO
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tag', required=True)
    parser.add_argument('--repo-root', type=Path, default=REPO)
    parser.add_argument('command', nargs=argparse.REMAINDER)
    args = parser.parse_args()
    REPO = args.repo_root.resolve()
    assert re.fullmatch(r'[a-z0-9_-]+', args.tag)
    argv = args.command[1:] if args.command[:1] == ['--'] else args.command
    if not argv:
        parser.error('command is required')
    receipt = ROOT / 'checks' / (args.tag + '.json')
    log = ROOT / 'checks' / (args.tag + '.log')
    assert not receipt.exists() and not log.exists() and not log.with_suffix('.log.gz').exists()
    env = {'RUSTUP_TOOLCHAIN': '1.98.1', 'CARGO_BUILD_JOBS': '4', 'CARGO_INCREMENTAL': '0', 'PYTHONDONTWRITEBYTECODE': '1'}
    if 'doc' in argv:
        env['RUSTDOCFLAGS'] = '-D warnings'
    record = {'change': 425, 'argv': argv, 'cwd': str(REPO), 'revision': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO).decode().strip(), 'environment': env, 'started_utc': now(), 'source_before': sources(), 'status': 'running'}
    receipt.write_text(json.dumps(record, indent=2, sort_keys=True) + '\n')
    with log.open('wb') as out:
        result = subprocess.run(argv, cwd=REPO, env=os.environ | env, stdout=out, stderr=subprocess.STDOUT)
    record.update(exit_code=result.returncode, finished_utc=now(), source_after=sources())
    record['source_unchanged'] = record['source_before'] == record['source_after']
    raw = log.read_bytes()
    record['log'] = {'path': str(log.relative_to(ROOT)), 'bytes': len(raw), 'sha256': hashlib.sha256(raw).hexdigest()}
    if 'test' in argv:
        counts = re.findall(rb'test result: ok\. (\d+) passed', raw)
        record['passed_tests'] = sum(map(int, counts))
    passed = result.returncode == 0 and record['source_unchanged'] and record.get('passed_tests', 1) > 0
    record['status'] = 'pass' if passed else 'failed'
    receipt.write_text(json.dumps(record, indent=2, sort_keys=True) + '\n')
    print(json.dumps({key: record[key] for key in ('status', 'exit_code', 'source_unchanged')}, sort_keys=True))
    return 0 if passed else 1


if __name__ == '__main__':
    sys.exit(main())
