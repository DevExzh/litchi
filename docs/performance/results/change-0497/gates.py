#!/usr/bin/env python3
"""Run the applicable production and harness gates serially."""
import argparse
import subprocess
import sys
import tomllib
from build import ROOT, write


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--attempt', required=True)
    parser.add_argument('--start-at')
    args = parser.parse_args()
    base = ['--release', '--locked', '--offline']
    harness = [*base, '--manifest-path', 'tools/perf-baseline/Cargo.toml']
    serial = ['--', '--test-threads=1']
    iwork = (
        'litchi-iwa-archive', 'litchi-iwa-structured', 'litchi-iwa-protos',
        'litchi-iwa-common', 'litchi-iwa', 'litchi-numbers', 'litchi-iwa-text',
        'litchi-iwa-graph', 'litchi-iwa-cache', 'litchi-pages', 'litchi-iwa-detect',
        'litchi-iwa-text-wire', 'litchi-iwa-index', 'litchi-iwa-package',
        'litchi-keynote', 'litchi-iwa-core',
    )
    exclusions = [argument for package in iwork for argument in ('--exclude', package)]
    workspace = tomllib.loads((ROOT.parents[3] / 'Cargo.toml').read_text())
    packages = [tomllib.loads((directory / 'Cargo.toml').read_text())['package']['name'] for member in workspace['workspace']['members'] for directory in ROOT.parents[3].glob(member) if (directory / 'Cargo.toml').is_file()]
    if not packages:
        raise RuntimeError('workspace package inventory must not be empty')
    fmt_packages = [argument for package in packages if package not in iwork for argument in ('-p', package)]
    jobs = [
        ('focused', ['cargo', 'test', *base, '-p', 'litchi-docx', '--test', 'source_backed_tail_append_stream', *serial]),
        ('docx-default', ['cargo', 'test', *base, '-p', 'litchi-docx', '--all-targets', *serial]),
        ('docx-features', ['cargo', 'test', *base, '-p', 'litchi-docx', '--all-features', '--all-targets', *serial]),
        ('opc-atomic', ['cargo', 'test', *base, '-p', 'litchi-opc', '--lib', 'atomic', *serial]),
        ('opc-preservation', ['cargo', 'test', *base, '-p', 'litchi-opc', '--all-features', '--all-targets', *serial]),
        ('doc-tests', ['cargo', 'test', *base, '-p', 'litchi-docx', '-p', 'litchi-opc', '--all-features', '--doc', *serial]),
        ('workspace-check', ['cargo', 'check', *base, '--workspace', '--all-targets', *exclusions]),
        ('production-clippy', ['cargo', 'clippy', *base, '-p', 'litchi-docx', '-p', 'litchi-opc', '--all-features', '--all-targets', '--', '-D', 'warnings']),
        ('production-rustdoc', ['cargo', 'doc', *base, '-p', 'litchi-docx', '-p', 'litchi-opc', '--all-features', '--no-deps']),
        ('harness-allocator', ['cargo', 'test', *harness, '--lib', '--features', 'allocator-metrics', *serial]),
        ('harness-clippy', ['cargo', 'clippy', *harness, '--all-targets', '--features', 'allocator-metrics', '--', '-D', 'warnings']),
        ('harness-rustdoc', ['cargo', 'doc', *harness, '--lib', '--features', 'allocator-metrics', '--no-deps']),
        ('fmt', ['cargo', 'fmt', *fmt_packages, '--', '--check']),
        ('harness-fmt', ['cargo', 'fmt', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '-p', 'litchi-perf-baseline', '--', '--check']),
        ('boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py']),
    ]
    if args.start_at:
        names = [name for name, _ in jobs]
        if args.start_at not in names:
            raise ValueError('unknown start gate')
        jobs = jobs[names.index(args.start_at):]
    receipts = {}
    for label, argv in jobs:
        name = args.attempt + '-' + label
        command = [sys.executable, '-B', str(ROOT / 'gate.py'), name, *argv]
        code = subprocess.call(command)
        receipts[label] = {'receipt': str(ROOT / 'validation' / (name + '.json')), 'wrapper_exit_code': code}
        if code:
            write(ROOT / ('gates-' + args.attempt + '-failed.json'), receipts)
            raise SystemExit(code)
    write(ROOT / ('gates-' + args.attempt + '.json'), receipts)


if __name__ == '__main__':
    main()
