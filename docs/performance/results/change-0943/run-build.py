#!/usr/bin/env python3
from __future__ import annotations
import argparse, hashlib, json, os, subprocess, sys
from datetime import datetime, timezone
from pathlib import Path

def now() -> str:
    return datetime.now(timezone.utc).isoformat().replace('+00:00','Z')

def digest(path: Path) -> dict[str, object]:
    value = {'path': str(path.resolve()), 'bytes': path.stat().st_size, 'sha256': None}
    if path.is_file():
        h = hashlib.sha256()
        with path.open('rb') as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b''):
                h.update(chunk)
        value['sha256'] = h.hexdigest()
    return value

def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument('--name', required=True)
    p.add_argument('--source', required=True, type=Path)
    p.add_argument('--target', required=True, type=Path)
    p.add_argument('--receipt-root', required=True, type=Path)
    p.add_argument('--binary', required=True)
    p.add_argument('--source-receipt', required=True, type=Path)
    args = p.parse_args()
    root = args.receipt_root / args.name
    root.mkdir(parents=True, exist_ok=False)
    source = args.source.resolve()
    target = args.target.resolve()
    manifest = source / 'docs/performance/results/change-0917/probe/Cargo.toml'
    binary_name = args.binary
    command = [
        '/home/zhuhe/.cargo/bin/cargo', '+1.95.0', 'build', '--offline', '--locked', '-j2', '--release', '--message-format=json-render-diagnostics',
    ]
    command += ['--bin', binary_name, '--manifest-path', str(manifest), '--target-dir', str(target)]
    keep = ('HOME', 'CARGO_HOME', 'RUSTUP_HOME', 'PATH', 'LC_ALL', 'LANG', 'TZ')
    environment = {key: os.environ[key] for key in keep if key in os.environ}
    environment.update({'CARGO_TERM_COLOR': 'never'})
    environment['CARGO_PROFILE_RELEASE_DEBUG'] = '1'
    environment['CARGO_INCREMENTAL'] = '0'
    environment.pop('CARGO_ENCODED_RUSTFLAGS', None)
    stdout_path = root / 'build.stdout.log'
    stderr_path = root / 'build.stderr.log'
    start = now()
    start_record = {
        'schema': 'litchi.performance.0943.build-start.v1',
        'name': args.name,
        'argv': command,
        'cwd': str(source),
        'source_root': str(source),
        'source_receipt': digest(args.source_receipt),
        'manifest': digest(manifest),
        'lock': digest(manifest.parent / 'Cargo.lock'),
        'target_dir': str(target),
        'environment_policy': 'allowlisted build variables; normal release debug1; no instrumentation flags; encoded flags unset',
        'environment': dict(sorted(environment.items())),
        'started_utc': start,
    }
    (root / 'start.json').write_text(json.dumps(start_record, indent=2, sort_keys=True) + '\n')
    with stdout_path.open('xb') as stdout, stderr_path.open('xb') as stderr:
        process = subprocess.Popen(command, cwd=str(source), env=environment, stdout=stdout, stderr=stderr)
        exit_code = process.wait()
    ended = now()
    terminal = {
        'schema': 'litchi.performance.0943.build-terminal.v1',
        'name': args.name,
        'argv': command,
        'cwd': str(source),
        'source_root': str(source),
        'source_receipt': digest(args.source_receipt),
        'manifest': digest(manifest),
        'lock': digest(manifest.parent / 'Cargo.lock'),
        'target_dir': str(target),
        'environment_policy': start_record['environment_policy'],
        'environment': dict(sorted(environment.items())),
        'started_utc': start,
        'ended_utc': ended,
        'exit_code': exit_code,
        'stdout': digest(stdout_path),
        'stderr': digest(stderr_path),
        'binary_name': binary_name,
        'terminal_written_before_binary_descriptor': True,
    }
    terminal_path = root / 'terminal.json'
    terminal_path.write_text(json.dumps(terminal, indent=2, sort_keys=True) + '\n')
    binary_path = target / 'release' / binary_name
    binary = digest(binary_path) if binary_path.is_file() else None
    (root / 'binary.json').write_text(json.dumps(binary, indent=2, sort_keys=True) + '\n')
    return exit_code

if __name__ == '__main__':
    raise SystemExit(main())
