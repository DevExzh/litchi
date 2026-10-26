#!/usr/bin/env python3
"""Run Cargo's default test target set in bounded groups, pruning executables.

Libraries, binary test harnesses, integration tests and doctests run normally.
Examples with test=false are built under the test profile; examples with
 test=true run as test harnesses. Only Cargo-reported executable artifacts are
removed, after their command finishes. Libraries and build scripts are kept.
"""
from __future__ import annotations
import argparse
import json
import os
from pathlib import Path
import shlex
import subprocess
import sys


def plan(metadata, selected, features, chunk_size):
    commands = []
    for package in metadata['packages']:
        if package['name'] not in selected:
            continue
        name = package['name']
        targets = package['targets']
        base = ['-p', name, *features]
        libraries = [t for t in targets if set(t['kind']) & {'lib', 'rlib', 'dylib', 'cdylib', 'staticlib', 'proc-macro'}]
        if any(t['test'] for t in libraries):
            commands.append(('test', [*base, '--lib']))
        for kind, switch in [('bin', '--bin'), ('test', '--test'), ('example', '--example')]:
            runnable = [t for t in targets if kind in t['kind'] and t['test']]
            for start in range(0, len(runnable), chunk_size):
                selectors = [item for t in runnable[start:start + chunk_size] for item in (switch, t['name'])]
                commands.append(('test', [*base, *selectors]))
        examples = [t for t in targets if 'example' in t['kind'] and not t['test']]
        for start in range(0, len(examples), chunk_size):
            selectors = [item for t in examples[start:start + chunk_size] for item in ('--example', t['name'])]
            commands.append(('build', [*base, '--profile', 'test', *selectors]))
        if any(t['doctest'] for t in libraries):
            commands.append(('test', [*base, '--doc']))
    return commands


def execute(command, target_dir):
    print('+ ' + shlex.join(command), flush=True)
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    artifacts = set()
    for line in process.stdout:
        try:
            message = json.loads(line)
        except (ValueError, TypeError):
            print(line, end='', flush=True)
            continue
        if not isinstance(message, dict) or 'reason' not in message:
            print(line, end='', flush=True)
            continue
        if message['reason'] == 'compiler-artifact':
            executable = message.get('executable')
            if executable and 'custom-build' not in message['target']['kind']:
                artifacts.add(Path(executable))
                # Cargo also reports the unhashed hard link in filenames.
                artifacts.update(Path(p) for p in message.get('filenames', []) if Path(p).is_file() and os.access(p, os.X_OK))
        elif message['reason'] == 'compiler-message':
            print(message['message'].get('rendered') or message['message']['message'], end='', flush=True)
    status = process.wait()
    for path in artifacts:
        if path.is_relative_to(target_dir) and path.suffix not in ('.so', '.dylib', '.dll'):
            path.unlink(missing_ok=True)
    print(f'command_status={status}; executable_artifacts_pruned={len(artifacts)}', flush=True)
    return status


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('-p', '--package', action='append', default=[])
    parser.add_argument('--workspace', action='store_true')
    parser.add_argument('--exclude', action='append', default=[])
    parser.add_argument('--all-features', action='store_true')
    parser.add_argument('--features')
    parser.add_argument('--no-default-features', action='store_true')
    parser.add_argument('--chunk-size', type=int, default=8)
    parser.add_argument('--plan', action='store_true')
    args = parser.parse_args()
    if args.chunk_size < 1:
        parser.error('--chunk-size must be positive')
    metadata = json.loads(subprocess.check_output(['cargo', 'metadata', '--format-version', '1', '--no-deps'], text=True))
    workspace = set(metadata['workspace_members'])
    names = {p['name'] for p in metadata['packages'] if p['id'] in workspace}
    selected = (names if args.workspace else set(args.package)) - set(args.exclude)
    if not selected or selected - names:
        parser.error('select existing workspace packages')
    features = []
    if args.all_features:
        features.append('--all-features')
    if args.no_default_features:
        features.append('--no-default-features')
    if args.features:
        features += ['--features', args.features]
    commands = plan(metadata, selected, features, args.chunk_size)
    failures = []
    target_dir = Path(metadata['target_directory']).resolve()
    for index, (verb, arguments) in enumerate(commands):
        command = ['cargo', verb, *arguments]
        if verb == 'test':
            command += ['--no-fail-fast']
        if args.plan:
            print(shlex.join(command))
            continue
        command += ['--message-format', 'json-render-diagnostics']
        subprocess.run([sys.executable, str(Path(__file__).with_name('target_headroom.py')), str(target_dir)], check=True)
        status = execute(command, target_dir)
        if status:
            failures.append({'index': index, 'status': status, 'command': command})
    print(json.dumps({'packages': sorted(selected), 'commands': len(commands), 'failures': failures}))
    return bool(failures)


if __name__ == '__main__':
    sys.exit(main())
