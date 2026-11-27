#!/usr/bin/env python3
"""Run Cargo's default test target set in bounded groups, pruning executables.

Libraries, binary test harnesses, integration tests and doctests run normally.
Examples with test=false are built under the test profile; examples with
 test=true run as test harnesses. Only Cargo-reported executable artifacts are
removed, after their command finishes. Libraries and build scripts are kept.

--part and --partition split the plan across runners. Every command keeps the
same package selection and feature arguments, so Cargo's feature unification
is identical in every part and partition.
"""
from __future__ import annotations
import argparse
import json
import os
from pathlib import Path
import shlex
import subprocess
import sys


def plan(packages, selection, enabled_features, chunk_size, lib_tests_only):
    """Keep Cargo's package and feature selection fixed across all batches."""
    targets = [
        target
        for package in packages
        for target in package['targets']
        if set(target.get('required-features', ())) <= enabled_features[package['name']]
    ]
    library_kinds = {'lib', 'rlib', 'dylib', 'cdylib', 'staticlib', 'proc-macro'}
    libraries = [target for target in targets if library_kinds.intersection(target['kind'])]
    commands = []
    if any(target['test'] for target in libraries):
        commands.append(('test', [*selection, '--lib']))

    def batches(verb, kind, test, extra=()):
        # A named target may occur in several selected crates. Cargo runs all
        # of them when the name is requested once.
        names = sorted({
            target['name'] for target in targets
            if kind in target['kind'] and target['test'] == test
        })
        for start in range(0, len(names), chunk_size):
            arguments = [*selection, *extra]
            for name in names[start:start + chunk_size]:
                arguments.extend([f'--{kind}', name])
            commands.append((verb, arguments))

    batches('test', 'bin', True)
    batches('test', 'test', True)
    batches('test', 'example', True)
    if not lib_tests_only:
        batches('build', 'example', False, ['--profile', 'test'])
        if any(target['doctest'] for target in libraries):
            commands.append(('test', [*selection, '--doc']))
    return commands


def is_lib_command(command):
    verb, arguments = command
    return verb == 'test' and arguments[-1:] == ['--lib']


def select(commands, part, partition):
    """Keep the commands of one part and, within it, of one partition.

    The `--lib` command belongs to part `lib`; every other command belongs to
    part `tests`. A partition K/N applies to the non-`--lib` commands only and
    keeps the i-th of them when i % N == K - 1.
    """
    index, count = partition
    lib = [command for command in commands if is_lib_command(command)]
    others = [command for command in commands if not is_lib_command(command)]
    others = [command for position, command in enumerate(others) if position % count == index - 1]
    return [*(lib if part in ('all', 'lib') else []), *(others if part in ('all', 'tests') else [])]


def parse_partition(value):
    """Parse K/N with 1 <= K <= N."""
    try:
        index, count = (int(part) for part in value.split('/'))
    except ValueError:
        raise argparse.ArgumentTypeError(f'expected K/N, got {value!r}') from None
    if not 1 <= index <= count:
        raise argparse.ArgumentTypeError(f'partition {value!r} needs 1 <= K <= N')
    return index, count


def selected_features(selection):
    """Ask Cargo which features are active on each selected root package."""
    output = subprocess.check_output(
        ['cargo', 'tree', *selection, '--color', 'never', '--depth', '0',
         '--prefix', 'none', '--format', '{p}\t{f}'],
        text=True, encoding='utf-8',
    )
    features = {}
    for line in output.splitlines():
        if '\t' in line:
            package, active = line.split('\t', 1)
            features[package.split()[0]] = set(active.split(',')) if active else set()
    return features


def execute(command, target_dir):
    print('+ ' + shlex.join(command), flush=True)
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                               text=True, encoding='utf-8', errors='replace')
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
                # For an uplifted binary or example, Cargo reports only the
                # unhashed hard link, so the hashed copy beside it stays.
                # Test executables are not uplifted and are removed whole.
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
    # Test output and diagnostics carry non-ASCII text; a Windows pipe would
    # otherwise encode it with the console code page and fail.
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('-p', '--package', action='append', default=[])
    parser.add_argument('--workspace', action='store_true')
    parser.add_argument('--exclude', action='append', default=[])
    parser.add_argument('--all-features', action='store_true')
    parser.add_argument('--features')
    parser.add_argument('--no-default-features', action='store_true')
    parser.add_argument('--lib-tests', action='store_true',
                        help='run the targets selected by cargo test --lib --tests')
    parser.add_argument('--chunk-size', type=int, default=8)
    parser.add_argument('--part', choices=('all', 'lib', 'tests'), default='all',
                        help='lib: only the --lib command; tests: every other command')
    parser.add_argument('--partition', type=parse_partition, default=(1, 1), metavar='K/N',
                        help='of the commands other than --lib, keep the i-th when i %% N == K - 1')
    parser.add_argument('--plan', action='store_true',
                        help='print the selected commands without running them')
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
    selection = ['--workspace'] if args.workspace else []
    package_flag = '--exclude' if args.workspace else '-p'
    for name in sorted(args.exclude if args.workspace else selected):
        selection.extend([package_flag, name])
    selection.extend(features)
    packages = [p for p in metadata['packages'] if p['name'] in selected]
    commands = select(
        plan(packages, selection, selected_features(selection), args.chunk_size, args.lib_tests),
        args.part, args.partition,
    )
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
    print(json.dumps({'packages': sorted(selected), 'part': args.part,
                      'partition': '{}/{}'.format(*args.partition),
                      'commands': len(commands), 'failures': failures}))
    return bool(failures)


if __name__ == '__main__':
    sys.exit(main())
