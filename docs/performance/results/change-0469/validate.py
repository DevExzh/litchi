#!/usr/bin/env python3
"""Run the candidate's correctness gates serially, preserving each result."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
CHECKS = [
    ('xlsx-fmt', ['cargo', 'fmt', '-p', 'litchi-xlsx', '--', '--check']),
    ('xlsx-tests', ['cargo', 'test', '--locked', '-p', 'litchi-xlsx', '--all-features']),
    ('workspace-check', ['cargo', 'check', '--locked', '--workspace', '--all-features']),
    ('xlsx-clippy', ['cargo', 'clippy', '--locked', '-p', 'litchi-xlsx', '--all-features', '--lib', '--no-deps', '--', '-D', 'warnings']),
    ('xlsx-rustdoc', ['cargo', 'doc', '--locked', '-p', 'litchi-xlsx', '--all-features', '--no-deps']),
    ('boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py']),
]


def main():
    output = ROOT / 'validation'
    output.mkdir(exist_ok=True)
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4',
               CARGO_INCREMENTAL='0', RUSTDOCFLAGS='-D warnings', PYTHONDONTWRITEBYTECODE='1')
    receipt = output / 'receipt.json'
    checks = json.loads(receipt.read_text())['checks'] if receipt.exists() else []
    for name, argv in CHECKS:
        if any(check['name'] == name for check in checks):
            continue
        log = output / (name + '.log')
        record = dict(name=name, argv=argv, started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
        with log.open('x') as stream:
            result = subprocess.run(argv, cwd=REPO, env=env, stdout=stream, stderr=subprocess.STDOUT)
        record.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                      log=log.name, log_sha256=hashlib.sha256(log.read_bytes()).hexdigest())
        checks.append(record)
        (output / 'receipt.json').write_text(json.dumps(dict(checks=checks, environment={k: env[k] for k in ['RUSTUP_TOOLCHAIN', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'RUSTDOCFLAGS']}), indent=2)+'\n')
        print(json.dumps(record), flush=True)
        if result.returncode:
            raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()
