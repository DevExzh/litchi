#!/usr/bin/env python3
"""Run scoped ODG checks into a fresh evidence directory."""
import argparse
import json
import os
from pathlib import Path
import re
import subprocess

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--target-dir', type=Path, required=True)
parser.add_argument('--output', type=Path, required=True)
parser.add_argument('--without-boundaries', action='store_true', help='Run crate checks while boundaries are recorded separately')
args = parser.parse_args()
args.output.mkdir(parents=True, exist_ok=True)
cargo_env = {'CARGO_TARGET_DIR': str(args.target_dir.resolve()), 'CARGO_BUILD_JOBS': '2',
             'CARGO_PROFILE_DEV_DEBUG': '0', 'CARGO_PROFILE_TEST_DEBUG': '0', 'CARGO_INCREMENTAL': '0'}
commands = [
    ('test-all-targets', ['cargo', 'test', '-p', 'litchi-odg', '--all-targets']),
    ('fmt', ['cargo', 'fmt', '-p', 'litchi-odg', '--', '--check']),
    ('clippy', ['cargo', 'clippy', '-p', 'litchi-odg', '--all-targets', '--', '-D', 'warnings']),
    ('doctests', ['cargo', 'test', '-p', 'litchi-odg', '--doc']),
    ('rustdoc', ['cargo', 'doc', '-p', 'litchi-odg', '--no-deps']),
    ('downstream', ['cargo', 'check', '-p', 'litchi-odf', '--no-default-features', '--features', 'odg']),
    ('boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py']),
]
if args.without_boundaries:
    commands = [(name, command) for name, command in commands if name != 'boundaries']
results = []
for name, command in commands:
    env = {**os.environ, **cargo_env}
    if name == 'rustdoc':
        env['RUSTDOCFLAGS'] = '-D warnings'
    log = args.output / (name + '.log')
    with log.open('w') as output:
        completed = subprocess.run(command, env=env, stdout=output, stderr=subprocess.STDOUT)
    counts = [int(value) for value in re.findall(r'test result: ok\. (\d+) passed', log.read_text())]
    results.append({'name': name, 'command': command, 'exit_code': completed.returncode,
                    'log': log.name, 'passed_tests': sum(counts) if counts else None})
    (args.output / 'gates.json').write_text(json.dumps({
        'scope': 'ODG crate and ODF umbrella ODG-only feature; not full workspace, fuzz, sanitizer or live Office',
        'cargo_environment': cargo_env, 'rustdoc_environment': {'RUSTDOCFLAGS': '-D warnings'},
        'gates': results}, indent=2) + '\n')
    print(name, completed.returncode, flush=True)
    if completed.returncode:
        raise SystemExit(completed.returncode)
