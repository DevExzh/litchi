#!/usr/bin/env python3
"""Bind the managed read-ahead results, final gates, and cleanup inventory."""

import json
from pathlib import Path
import sys
import cleanup as cleanup_driver

import measure
from support import ROOT, TEMP, TARGET_DIR, meta, read, write


FINAL_GATES_SCHEMA = 'docx-managed-read-ahead-final-gates-v1'
FINAL_GATES_VERSION = 1


def _expected_final_commands():
    """Return the exact command vectors accepted as final validation."""

    tmp = f'TMPDIR={TEMP / "test-tmp"}'
    cargo = ['env', tmp, 'cargo']
    python = ['env', tmp, 'python3']
    rustdoc = ['env', tmp, 'RUSTDOCFLAGS=-Dwarnings', 'cargo']
    fmt = [
        'rustfmt', '+1.98.1', '--edition', '2024', '--check', '--config',
        'skip_children=true',
        'crates/litchi-opc/src/lib.rs',
        'crates/litchi-opc/src/source_backed.rs',
        'crates/litchi-opc/src/source_backed/artifact_restore.rs',
        'crates/litchi-opc/src/source_backed/splice.rs',
        'crates/litchi-opc/src/source_backed/read_ahead.rs',
        'crates/litchi-opc/tests/source_read_ahead.rs',
        'crates/litchi-docx/src/source_backed.rs',
        'crates/litchi-docx/tests/source_read_policy.rs',
        'tools/perf-baseline/src/lib.rs',
        'tools/perf-baseline/src/main.rs',
        'tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs',
        'tools/perf-baseline/src/docx_managed_read_ahead.rs',
    ]
    return (
        tuple(cargo + ['test', '--release', '--locked', '-p', 'litchi-opc', '--all-features']),
        tuple(cargo + ['test', '--release', '--locked', '-p', 'litchi-docx']),
        tuple(cargo + [
            'clippy', '--release', '--locked', '-p', 'litchi-opc', '-p',
            'litchi-docx', '--all-targets', '--all-features', '--', '-D',
            'warnings',
        ]),
        tuple(cargo + [
            'clippy', '--release', '--locked', '--manifest-path',
            'tools/perf-baseline/Cargo.toml', '--all-targets', '--features',
            'allocator-metrics', '--', '-D', 'warnings',
        ]),
        tuple(cargo + [
            'test', '--release', '--locked', '--manifest-path',
            'tools/perf-baseline/Cargo.toml', '--lib', '--features',
            'allocator-metrics', '--', '--test-threads=1',
        ]),
        tuple(rustdoc + [
            'doc', '--release', '--locked', '-p', 'litchi-opc', '-p',
            'litchi-docx', '--all-features', '--no-deps',
        ]),
        tuple(rustdoc + [
            'doc', '--release', '--locked', '--manifest-path',
            'tools/perf-baseline/Cargo.toml', '--lib', '--features',
            'allocator-metrics', '--no-deps',
        ]),
        ('python3', '-B', 'tools/check_crate_boundaries.py'),
        tuple(python + [
            '-B', '-m', 'unittest', 'discover', '-s', str(ROOT),
            '-p', 'test_*.py',
        ]),
        tuple(measure._expected_build_command('normal')),
        tuple(measure._expected_build_command('allocator')),
        tuple(fmt),
    )


def _safe_gate_label(value):
    measure.require(
        isinstance(value, str) and value and value not in {'.', '..'}
        and '/' not in value and '\\' not in value
        and not any(character.isspace() for character in value),
        'final gate label is not path-safe',
    )
    return value


def _validate_final_gates(value):
    """Validate the strict command manifest and return its label/argv pairs."""

    measure._exact(value, ('schema', 'version', 'commands'), 'final-gates.json')
    measure.require(
        value['schema'] == FINAL_GATES_SCHEMA and value['version'] == FINAL_GATES_VERSION,
        'final-gates.json schema differs',
    )
    commands = value['commands']
    measure.require(isinstance(commands, dict), 'final-gates.json commands must be an object')
    expected = set(_expected_final_commands())
    actual = []
    for label, command in commands.items():
        _safe_gate_label(label)
        measure.require(
            isinstance(command, list)
            and all(isinstance(item, str) and item for item in command),
            f'{label}: final gate argv is malformed',
        )
        actual.append(tuple(command))
    measure.require(len(actual) == len(expected), 'final-gate count differs')
    measure.require(len(set(actual)) == len(actual), 'final-gate commands are duplicated')
    measure.require(set(actual) == expected, 'final-gate command set differs')
    return [(label, list(command)) for label, command in commands.items()]


def _validate_gate_receipt(label, command, path, source):
    """Bind one final gate to its filename, source, and exact command."""

    measure.require(path.is_file() and not path.is_symlink(), f'{label}: gate receipt is missing')
    receipt = read(path)
    measure.require(isinstance(receipt, dict), f'{label}: gate receipt is not an object')
    measure.require(receipt.get('label') == path.stem == label,
                    f'{label}: receipt label differs from filename')
    binding = measure._gate_binding(
        {'path': str(path), 'sha256': meta(path)['sha256']}, path
    )
    measure.require(binding['source'] == source, f'{label}: source differs')
    measure.require(binding['argv'] == command, f'{label}: command differs')
    return binding


def _verify_cleanup():
    """Run the canonical cleanup verifier before sealing evidence."""

    proof = cleanup_driver.verify(root=ROOT, temp=TEMP, target=TARGET_DIR)
    measure.require(isinstance(proof, dict) and proof.get('status') == 'pass',
                    'cleanup verification did not pass')
    return proof


def inventory(root=ROOT):
    files = {}
    for path in sorted(root.rglob('*')):
        if path.is_symlink():
            raise RuntimeError(f'symlink in evidence: {path}')
        if path == root / 'seal.json':
            if not path.is_file():
                raise RuntimeError(f'special file in evidence: {path}')
            continue
        if path.is_file():
            if '__pycache__' in path.parts or path.suffix == '.pyc':
                raise RuntimeError(f'generated bytecode in evidence: {path}')
            files[path.relative_to(root).as_posix()] = meta(path)
        elif not path.is_dir():
            raise RuntimeError(f'special file in evidence: {path}')
    return files


def verify_evidence():
    builds = measure.load_builds()
    protocol, protocol_hash = measure._load_protocol(builds)
    measure.require(protocol == measure.protocol_value(builds), 'protocol/build mismatch')
    accepted = read(ROOT / 'accepted-evidence.json')
    measure.require(accepted['formal_samples'] == 480 and accepted['pilot_samples'] == 24,
                    'accepted sample inventory differs')
    for pilot, key in ((True, 'pilot_attempt'), (False, 'formal_attempt')):
        measure.verify(accepted[key], pilot=pilot)

    gates = read(ROOT / 'final-gates.json')
    final_commands = _validate_final_gates(gates)
    for label, command in final_commands:
        path = ROOT / 'validation' / f'{label}.json'
        _validate_gate_receipt(label, command, path, builds['normal']['source'])

    tested = read(ROOT / 'helper-test-custody.json')
    measure.require(tested['helpers'] == {name: meta(ROOT / name) for name in tested['helpers']},
                    'tested helper changed')
    measure.require(set(measure.DRIVER_FILES).issubset(tested['helpers']),
                    'capture helpers missing from tested inventory')
    measure.require(tested['gate'] == meta(ROOT / 'validation' / (tested['label'] + '.json')),
                    'helper test gate changed')
    measure.require(tested['label'] in gates['commands'], 'helper test is not a final gate')

    reproduction = read(ROOT / 'source-reproduction.json')
    measure.require(reproduction['patch'] == meta(ROOT / 'build-source.patch'),
                    'source reproduction patch changed')
    measure.require(reproduction['source'] == builds['normal']['source'],
                    'source reproduction manifest differs')

    for started in (ROOT / 'validation').glob('*.started.json'):
        terminal = started.with_name(started.name.replace('.started.json', '.json'))
        measure.require(terminal.is_file(), f'unfinished validation command: {started.name}')

    cleanup_proof = _verify_cleanup()
    cleanup = read(ROOT / 'cleanup.json')
    expected_files = {Path(build['binary']['path']) for build in builds.values()}
    actual_files = {p for p in TEMP.rglob('*') if p.is_file()}
    measure.require(actual_files == expected_files, 'unexpected temporary files remain')
    expected_dirs = set()
    for binary in expected_files:
        measure.require(binary.is_relative_to(TEMP), 'retained binary escaped scratch root')
        expected_dirs.update(p for p in binary.parents if p != TEMP and p.is_relative_to(TEMP))
    measure.require({p for p in TEMP.rglob('*') if p.is_dir()} == expected_dirs,
                    'unexpected temporary directories remain')
    measure.require(not any(p.is_symlink() for p in TEMP.rglob('*')), 'scratch symlink remains')
    return {'status': 'pass', 'formal_samples': 480, 'pilot_samples': 24,
            'protocol_sha256': protocol_hash, 'gates': sorted(gates['commands']),
            'cleanup_receipt': meta(ROOT / 'cleanup.json'),
            'cleanup_verification_schema': cleanup_proof['schema']}


def main():
    if sys.argv[1:] not in [['create'], ['verify']]:
        raise SystemExit('usage: verify_bundle.py create|verify')
    result = {'schema': 'docx-managed-read-ahead-bundle-v1',
              'evidence': verify_evidence(), 'files': inventory()}
    path = ROOT / 'seal.json'
    if sys.argv[1] == 'create':
        write(path, result)
    elif read(path) != result:
        raise RuntimeError('bundle inventory or evidence changed')
    print(json.dumps({'status': 'pass', 'files': len(result['files']), **meta(path)}, sort_keys=True))


if __name__ == '__main__':
    main()
