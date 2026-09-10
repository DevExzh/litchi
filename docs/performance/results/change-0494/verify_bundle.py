#!/usr/bin/env python3
"""Bind the DOCX edit/provider results, validation gates, and cleanup inventory."""

import json
import importlib
from pathlib import Path
import re
import sys
import cleanup as cleanup_driver

import measure
from support import ROOT, TEMP, TARGET_DIR, meta, read, write


FINAL_GATES_SCHEMA = 'docx-edit-provider-final-gates-v1'
FINAL_GATES_VERSION = 1
FORMAT_FILES = (
    'tools/perf-baseline/src/lib.rs',
    'tools/perf-baseline/src/main.rs',
    'tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs',
    'tools/perf-baseline/src/docx_edit_provider.rs',
    'tools/perf-baseline/src/docx_edit_provider/cold.rs',
)


def _expected_final_commands():
    """Return the frozen command vectors required for the 0494 result."""

    cargo = ['cargo']
    return (
        tuple(cargo + [
            'clippy', '--release', '--locked', '--manifest-path',
            'tools/perf-baseline/Cargo.toml', '--all-targets', '--features',
            'allocator-metrics', '--', '-D', 'warnings',
        ]),
        ('env', 'RUSTDOCFLAGS=-Dwarnings', *cargo, 'doc', '--release', '--locked',
         '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--lib',
         '--features', 'allocator-metrics', '--no-deps'),
        tuple(cargo + [
            'test', '--release', '--locked', '--manifest-path',
            'tools/perf-baseline/Cargo.toml', '--lib', '--features',
            'allocator-metrics', '--', '--test-threads=1',
        ]),
        ('rustfmt', '+1.98.1', '--edition', '2024', '--check', '--config',
         'skip_children=true', *FORMAT_FILES),
        ('python3', '-B', 'tools/check_crate_boundaries.py'),
        ('python3', '-B', '-m', 'unittest', 'discover', '-s', str(ROOT),
         '-p', 'test_*.py'),
        tuple(measure._expected_build_command('normal')),
        tuple(measure._expected_build_command('allocator')),
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
    actual = []
    for label, command in commands.items():
        _safe_gate_label(label)
        measure.require(
            isinstance(command, list)
            and all(isinstance(item, str) and item for item in command),
            f'{label}: final gate argv is malformed',
        )
        actual.append(tuple(command))
    measure.require(actual, 'final-gate command set is empty')
    measure.require(len(set(actual)) == len(actual), 'final-gate commands are duplicated')
    expected = set(_expected_final_commands())
    measure.require(len(actual) == len(expected), 'final-gate count differs')
    measure.require(set(actual) == expected, 'final-gate command set differs')
    measure.require(expected.issubset(set(actual)), 'mandatory final-gate command is missing')
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
    if tuple(command) != tuple(_expected_final_commands()[5]):
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
    root = Path(root)
    if root.is_symlink() or not root.is_dir():
        raise RuntimeError(f'evidence root must be a real directory: {root}')
    if root.resolve(strict=True) != root.absolute():
        raise RuntimeError(f'evidence root is not canonical: {root}')
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


def _expected_helpers():
    """Return every helper and test file that must be under custody."""

    names = set(measure.DRIVER_FILES)
    names.update({
        'cleanup.py', 'verify_bundle.py', 'test_cleanup.py',
        'test_measure.py', 'test_verify_bundle.py',
    })
    for optional in ('profile.py', 'test_profile.py', 'test_profile_strace_status.py',
                     'cold_measure.py', 'test_cold_measure.py',
                     'recover_profile.py', 'test_recover_profile.py',
                     'recover_observer_counters.py',
                     'test_recover_observer_counters.py'):
        if (ROOT / optional).is_file():
            names.add(optional)
    return names


def _cold_report_paths():
    """Find retained cold reports without treating prose or smoke gates as claims."""

    paths = []
    for path in (ROOT / 'captures', ROOT / 'analysis', ROOT / 'verification'):
        if not path.is_dir():
            continue
        for candidate in path.rglob('*.json'):
            if candidate.is_symlink():
                continue
            try:
                value = read(candidate)
            except (OSError, ValueError, json.JSONDecodeError):
                continue
            if not isinstance(value, dict):
                continue
            schema = value.get('schema')
            if (isinstance(schema, str) and 'cold' in schema
                    and ('edit_provider' in schema or 'edit-provider' in schema)):
                paths.append(candidate)
                continue
            if 'cold_verified_status' in value and 'cold_verified_samples' in value:
                paths.append(candidate)
    return sorted(paths)


def _cold_descriptor(value, key, *, expected_path=None):
    """Validate an accepted-evidence descriptor and return its digest."""

    measure.require(isinstance(value, dict), f'{key}: verification descriptor missing')
    path = value.get('path')
    digest = value.get('sha256')
    if path is None and expected_path is not None:
        candidate = Path(expected_path)
    else:
        measure.require(isinstance(path, str) and path, f'{key}.path: missing')
        candidate = Path(path)
        if not candidate.is_absolute():
            candidate = ROOT / candidate
    try:
        candidate = candidate.resolve(strict=True)
    except OSError as error:
        measure.fail(f'{key}: verification artifact cannot be resolved: {error}')
    measure.require(candidate.is_relative_to(ROOT.resolve()),
                    f'{key}: verification artifact escapes evidence root')
    measure.require(candidate.is_file() and not candidate.is_symlink(),
                    f'{key}: verification artifact is missing')
    actual = meta(candidate)
    measure.require(actual['sha256'] == digest, f'{key}: verification artifact changed')
    if 'bytes' in value:
        measure.require(value['bytes'] == actual['bytes'], f'{key}: verification artifact length changed')
    return actual['sha256']


_COLD_FORMAL_ALLOCATOR_ROWS = 60
_COLD_PILOT_ALLOCATOR_ROWS = 3
_COLD_ALLOCATOR_ROWS = _COLD_FORMAL_ALLOCATOR_ROWS + _COLD_PILOT_ALLOCATOR_ROWS
_COLD_ALLOCATOR_UINT_FIELDS = frozenset({
    'allocation_calls', 'deallocation_calls', 'reallocation_calls',
    'failed_allocation_calls', 'allocated_bytes', 'deallocated_bytes',
    'live_bytes_before', 'live_bytes_after', 'peak_live_bytes_before',
    'peak_live_bytes_after', 'region_peak_live_bytes',
})
_COLD_ALLOCATOR_FIELDS = frozenset({'status', 'scope'}) | _COLD_ALLOCATOR_UINT_FIELDS


def _strict_uint(value, path):
    """Require a JSON integer, excluding booleans and other numeric types."""

    measure.require(type(value) is int and value >= 0,
                    f'{path}: strict unsigned integer required')
    return value


def _validate_cold_allocator_row(allocation, path):
    """Validate allocator conservation and peak bounds for one raw row."""

    measure.require(isinstance(allocation, dict), f'{path}: allocator evidence missing')
    measure.require(set(allocation) == _COLD_ALLOCATOR_FIELDS,
                    f'{path}: allocator fields differ')
    measure.require(
        allocation['status'] == 'measured'
        and allocation['scope'] == measure.SAMPLE_ALLOCATION_SCOPE,
        f'{path}: allocator evidence is not measured operation-scoped data',
    )
    for field in _COLD_ALLOCATOR_UINT_FIELDS:
        _strict_uint(allocation[field], f'{path}.{field}')

    before = allocation['live_bytes_before']
    after = allocation['live_bytes_after']
    allocated = allocation['allocated_bytes']
    deallocated = allocation['deallocated_bytes']
    measure.require(
        after == before + allocated - deallocated,
        f'{path}: live-byte conservation differs',
    )

    peak_before = allocation['peak_live_bytes_before']
    peak_after = allocation['peak_live_bytes_after']
    region_peak = allocation['region_peak_live_bytes']
    measure.require(peak_before >= before,
                    f'{path}: pre-operation peak is below live bytes')
    measure.require(peak_after >= peak_before and peak_after >= after,
                    f'{path}: post-operation peak bound differs')
    measure.require(region_peak >= before and region_peak >= after,
                    f'{path}: operation peak is below live bytes')
    measure.require(region_peak <= peak_after,
                    f'{path}: operation peak exceeds process peak')
    measure.require(
        allocation['allocation_calls'] >= allocation['reallocation_calls'],
        f'{path}: reallocation calls exceed allocation calls',
    )
    return allocation


def _validate_cold_allocator_entries(formal_entries, pilot_entries):
    """Validate every allocator row and enforce the frozen 60/3 inventory."""

    def check(entries, phase):
        measure.require(isinstance(entries, list), f'{phase}: raw capture entries are malformed')
        count = 0
        for entry_index, entry in enumerate(entries):
            measure.require(isinstance(entry, dict),
                            f'{phase}[{entry_index}]: raw capture entry is malformed')
            spec = entry.get('spec')
            rows = entry.get('rows')
            measure.require(isinstance(spec, dict) and isinstance(rows, list),
                            f'{phase}[{entry_index}]: raw capture entry is incomplete')
            role = spec.get('role')
            measure.require(role in {'normal', 'allocator'},
                            f'{phase}[{entry_index}]: capture role is unknown')
            for row_index, row in enumerate(rows):
                measure.require(isinstance(row, dict),
                                f'{phase}[{entry_index}][{row_index}]: raw row is malformed')
                if role == 'allocator':
                    _validate_cold_allocator_row(
                        row.get('allocation'),
                        f'{phase}[{entry_index}][{row_index}].allocation',
                    )
                    count += 1
                else:
                    measure.require(row.get('allocation') is None,
                                    f'{phase}[{entry_index}][{row_index}]: normal row has allocator data')
        return count

    formal_count = check(formal_entries, 'cold formal')
    pilot_count = check(pilot_entries, 'cold pilot')
    measure.require(formal_count == _COLD_FORMAL_ALLOCATOR_ROWS,
                    'cold formal allocator row inventory differs')
    measure.require(pilot_count == _COLD_PILOT_ALLOCATOR_ROWS,
                    'cold pilot allocator row inventory differs')
    measure.require(formal_count + pilot_count == _COLD_ALLOCATOR_ROWS,
                    'cold allocator row inventory differs')
    return {
        'formal_allocator_rows': formal_count,
        'pilot_allocator_rows': pilot_count,
        'allocator_rows': formal_count + pilot_count,
    }


def _verify_cold_allocator_conservation(
    cold_measure, builds, warm_protocol_hash, formal_attempt, pilot_attempt,
):
    """Re-read accepted cold raw reports and check allocator conservation."""

    load_protocol = getattr(cold_measure, '_load_protocol', None)
    collect = getattr(cold_measure, '_collect', None)
    measure.require(callable(load_protocol) and callable(collect),
                    'cold raw-report collection API is missing')
    protocol, protocol_hash = load_protocol(builds)
    warm_binding = protocol.get('warm_protocol') if isinstance(protocol, dict) else None
    measure.require(
        isinstance(warm_binding, dict)
        and warm_binding.get('sha256') == warm_protocol_hash,
        'cold raw reports are not bound to the warm protocol',
    )
    formal_entries = collect(
        formal_attempt, builds, protocol, protocol_hash, pilot=False,
    )
    pilot_entries = collect(
        pilot_attempt, builds, protocol, protocol_hash, pilot=True,
    )
    return _validate_cold_allocator_entries(formal_entries, pilot_entries)


def _verify_cold_acceptance(accepted, warm_protocol_hash, builds=None):
    """Require both cold captures whenever a retained cold report exists."""

    cold = accepted.get('cold')
    flat_claim = any(key in accepted for key in (
        'cold_formal_attempt', 'cold_pilot_attempt',
        'cold_formal_verification', 'cold_pilot_verification',
    ))
    claimed = bool(_cold_report_paths()) or flat_claim or (
        isinstance(cold, dict)
        and (cold.get('claimed') is True
             or 'formal_attempt' in cold
             or 'cold_formal_attempt' in cold)
    )
    if not claimed:
        return None

    if not isinstance(cold, dict):
        cold = accepted
    try:
        cold_measure = importlib.import_module('cold_measure')
    except ImportError as error:
        measure.fail(f'cold report is present but cold_measure.py is unavailable: {error}')

    formal_attempt = cold.get('formal_attempt', cold.get('cold_formal_attempt'))
    pilot_attempt = cold.get('pilot_attempt', cold.get('cold_pilot_attempt'))
    measure.require(isinstance(formal_attempt, str) and formal_attempt,
                    'cold formal attempt is missing')
    measure.require(isinstance(pilot_attempt, str) and pilot_attempt,
                    'cold pilot attempt is missing')
    measure.require(cold.get('formal_samples', cold.get('cold_formal_samples')) == 120,
                    'cold formal sample inventory differs')
    measure.require(cold.get('pilot_samples', cold.get('cold_pilot_samples')) == 6,
                    'cold pilot sample inventory differs')

    base = cold.get('base_protocol_sha256', cold.get('warm_protocol_sha256'))
    if isinstance(base, dict):
        base = base.get('sha256')
    if base is None:
        cold_protocol_path = getattr(cold_measure, 'COLD_REPORT_PATH', None)
        if cold_protocol_path is not None and Path(cold_protocol_path).is_file():
            cold_protocol = cold_measure._json(Path(cold_protocol_path))
            if isinstance(cold_protocol, dict):
                warm_binding = cold_protocol.get('warm_protocol')
                if isinstance(warm_binding, dict):
                    base = warm_binding.get('sha256')
    measure.require(base == warm_protocol_hash,
                    'cold protocol is not bound to the warm protocol')

    formal_descriptor = cold.get('formal_verification', cold.get('cold_formal_verification'))
    pilot_descriptor = cold.get('pilot_verification', cold.get('cold_pilot_verification'))
    formal_path = Path(cold_measure.COLD_VERIFICATION_ROOT) / f'{formal_attempt}.json'
    pilot_path = Path(cold_measure.COLD_VERIFICATION_ROOT) / f'{pilot_attempt}-pilot.json'
    _cold_descriptor(formal_descriptor, 'cold formal verification', expected_path=formal_path)
    _cold_descriptor(pilot_descriptor, 'cold pilot verification', expected_path=pilot_path)

    verify = getattr(cold_measure, 'verify', None)
    measure.require(callable(verify), 'cold_measure.verify API is missing')
    for attempt, pilot in ((formal_attempt, False), (pilot_attempt, True)):
        try:
            result = verify(attempt, pilot=pilot)
        except TypeError:
            result = verify(attempt, pilot)
        measure.require(result is not None, f'cold verification returned no result: {attempt}')
        verification_path = Path(result)
        verification = cold_measure._json(verification_path)
        measure.require(isinstance(verification, dict)
                        and verification.get('status') == 'pass',
                        f'cold verification is not a passing proof: {attempt}')
    inventory = getattr(cold_measure, 'formal_inventory', None)
    if callable(inventory):
        measure.require(
            sum(int(item['samples']) for item in inventory()) == 120,
            'cold formal inventory is not 120 samples',
        )
        measure.require(
            sum(int(item['samples']) for item in inventory(pilot=True)) == 6,
            'cold pilot inventory is not 6 samples',
        )
    allocator_summary = _verify_cold_allocator_conservation(
        cold_measure, measure.load_builds() if builds is None else builds, warm_protocol_hash,
        formal_attempt, pilot_attempt,
    )
    return {'formal_attempt': formal_attempt, 'pilot_attempt': pilot_attempt,
            'formal_samples': 120, 'pilot_samples': 6,
            **allocator_summary}


def _profile_meta(path, label):
    """Return a rooted, immutable descriptor for one profile artifact."""

    path = Path(path)
    measure.require(path.is_file() and not path.is_symlink(),
                    f'{label}: regular file is missing')
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        measure.fail(f'{label}: cannot resolve artifact: {error}')
    measure.require(resolved.is_relative_to(ROOT.resolve()),
                    f'{label}: artifact escapes evidence root')
    return {'path': str(resolved), **meta(resolved)}


def _profile_helper_core(value, label='profile helper archive'):
    """Project one archive receipt to its immutable profile/test bindings."""

    measure.require(isinstance(value, dict), 'profile helper archive is not an object')
    required = ('custody', 'profile', 'test_profile', 'label')
    measure.require(all(key in value for key in required),
                    f'{label} fields are incomplete')
    current = value.get('current')
    matches = value.get('current_matches_archived')
    if current is not None or matches is not None:
        measure.require(isinstance(current, dict) and isinstance(matches, dict),
                        f'{label} helper drift custody is incomplete')
        expected_matches = {}
        for name in ('profile.py', 'test_profile.py'):
            current_meta = _profile_meta(ROOT / name, f'current {name}')
            measure.require(current.get(name) == current_meta,
                            f'current {name} custody changed')
            expected_matches[name] = all(
                current_meta[key] == value[name[:-3]][key]
                for key in ('bytes', 'sha256')
            )
        measure.require(matches == expected_matches,
                        f'{label} helper drift flags differ')
    return {key: value[key] for key in required}


def _profile_helper_archive(recovery, path):
    """Validate archived profile helpers while recording later helper drift honestly."""

    value = recovery._helper_archive(path)
    core = _profile_helper_core(value)
    for name in ('profile.py', 'test_profile.py'):
        archived_path = Path(path) / name
        archived = _profile_meta(archived_path, f'archived {name}')
        key = name[:-3]
        measure.require(core[key] == archived,
                        f'archived {name} changed')
    return core


def _historical_helper_descriptor(value, name, *, allow_unarchived=False):
    """Authenticate a helper hash against the current file or a retained archive."""

    measure.require(isinstance(value, dict), f'{name}: helper descriptor is missing')
    measure._exact(value, ('bytes', 'path', 'sha256'), f'{name}: helper descriptor')
    expected_path = (ROOT / name).resolve()
    measure.require(Path(value['path']).resolve() == expected_path,
                    f'{name}: helper path differs')
    measure.require(type(value['bytes']) is int and value['bytes'] >= 0
                    and isinstance(value['sha256'], str),
                    f'{name}: helper descriptor is malformed')
    current = _profile_meta(expected_path, f'current {name}')
    if current['bytes'] == value['bytes'] and current['sha256'] == value['sha256']:
        return value
    for candidate in sorted(ROOT.rglob(name)):
        if candidate.resolve() == expected_path or candidate.is_symlink() or not candidate.is_file():
            continue
        actual = _profile_meta(candidate, f'archived {name}')
        if actual['bytes'] != value['bytes'] or actual['sha256'] != value['sha256']:
            continue
        custody_path = candidate.parent / 'helper-custody.json'
        if custody_path.is_file() and not custody_path.is_symlink():
            try:
                custody = read(custody_path)
            except (OSError, ValueError, json.JSONDecodeError) as error:
                measure.fail(f'{custody_path}: helper archive is invalid: {error}')
            expected = custody.get('helpers', {}).get(name) if isinstance(custody, dict) else None
            measure.require(isinstance(expected, dict)
                            and expected.get('bytes') == actual['bytes']
                            and expected.get('sha256') == actual['sha256'],
                            f'{custody_path}: {name} binding differs')
        return value
    measure.require(allow_unarchived,
                    f'{name}: changed helper has no retained historical archive')
    return value


def _repo_relative(path):
    """Return one evidence path in the command spelling used by gate.py."""

    path = Path(path).resolve(strict=True)
    measure.require(path.is_relative_to(measure.REPO.resolve()),
                    'profile postprocessing path escapes repository')
    return path.relative_to(measure.REPO.resolve()).as_posix()


def _profile_recovery_command(profile_dir, build_path, gate_path, archive_dir,
                              output_dir):
    """Return the exact retained-data recovery argv for one output directory."""

    return [
        'python3', '-B', _repo_relative(ROOT / 'recover_profile.py'),
        '--profile-dir', _repo_relative(profile_dir),
        '--build-record', _repo_relative(build_path),
        '--gate-receipt', _repo_relative(gate_path),
        '--helper-archive', _repo_relative(archive_dir),
        '--cpu', '2',
        '--output-dir', _repo_relative(output_dir),
    ]


def _observer_correction_command(profile_dir, recovery_summary_path,
                                 recovery_script, build_path, gate_path,
                                 archive_dir, output_dir,
                                 recovery_helper_archive=None):
    """Return the exact retained-observer correction argv."""

    command = [
        'python3', '-B', _repo_relative(ROOT / 'recover_observer_counters.py'),
        '--profile-dir', _repo_relative(profile_dir),
        '--recovery-summary', _repo_relative(recovery_summary_path),
        '--recovery-script', _repo_relative(recovery_script),
        '--build-record', _repo_relative(build_path),
        '--gate-receipt', _repo_relative(gate_path),
        '--helper-archive', _repo_relative(archive_dir),
    ]
    if recovery_helper_archive is not None:
        command.extend(['--recovery-helper-archive',
                        _repo_relative(recovery_helper_archive)])
    command.extend(['--output-dir', _repo_relative(output_dir)])
    return command


def _passing_postprocess_gate(path, command, label):
    """Validate one passing postprocessing gate without frozen-Rust binding."""

    measure.require(path.is_file() and not path.is_symlink(),
                    f'{label}: postprocessing gate is missing')
    receipt = read(path)
    measure.require(isinstance(receipt, dict) and receipt.get('exit_code') == 0,
                    f'{label}: postprocessing gate did not pass')
    binding = measure._gate_binding(
        {'path': str(path), 'sha256': meta(path)['sha256']}, path
    )
    measure.require(binding['argv'] == command,
                    f'{label}: postprocessing command differs')
    return binding


def _attempt_rank(path):
    """Order retained ``*-rN`` postprocessing attempts by their numeric suffix."""

    path = Path(path)
    for component in reversed(path.parts):
        match = re.search(r'-r(\d+)(?:$|\.json$)', component)
        if match is not None:
            return int(match.group(1)), str(path)
    return -1, str(path)


def _select_profile_recovery(recovery, profile_dir, build_path, gate_path,
                             archive_dir):
    """Select exactly one passing retained-data recovery attempt."""

    selected = []
    for summary_path in sorted(ROOT.glob('profiling-recovery-*/recovery-summary.json')):
        if summary_path.is_symlink() or not summary_path.is_file():
            continue
        try:
            summary = recovery._json(summary_path)
        except (AttributeError, OSError, TypeError, ValueError,
                recovery.RecoveryError) as error:
            measure.fail(f'{summary_path}: profile recovery summary is invalid: {error}')
        if not (isinstance(summary, dict)
                and summary.get('schema') == recovery.SCHEMA
                and summary.get('version') == recovery.VERSION
                and summary.get('status') == 'pass'):
            continue
        output_dir = summary_path.parent
        expected = _profile_recovery_command(
            profile_dir, build_path, gate_path, archive_dir, output_dir
        )
        matches = []
        for gate_path_candidate in sorted(
                (ROOT / 'validation').glob('profile-recovery-*.json')):
            if gate_path_candidate.is_symlink() or not gate_path_candidate.is_file():
                continue
            try:
                receipt = read(gate_path_candidate)
            except (OSError, ValueError, json.JSONDecodeError) as error:
                measure.fail(
                    f'{gate_path_candidate}: profile recovery gate is invalid: {error}'
                )
            if not isinstance(receipt, dict) or receipt.get('exit_code') != 0:
                continue
            if receipt.get('argv') != expected:
                continue
            matches.append((gate_path_candidate, _passing_postprocess_gate(
                gate_path_candidate, expected, gate_path_candidate.stem
            )))
        measure.require(len(matches) == 1,
                        f'{summary_path}: passing recovery gate binding is ambiguous')
        selected.append((summary_path, output_dir, matches[0][0], matches[0][1]))
    measure.require(selected, 'successful profile recovery summary is missing')
    selected.sort(key=lambda item: _attempt_rank(item[0]))
    highest = _attempt_rank(selected[-1][0])
    measure.require(sum(_attempt_rank(item[0]) == highest for item in selected) == 1,
                    'successful profile recovery summary is ambiguous')
    return selected[-1]


def _verify_profile_recovery(builds):
    """Authenticate the failed original export and its retained-data recovery."""

    profile_dir = ROOT / 'profiling-r1'
    summary_path = profile_dir / 'profile-summary.json'
    if not summary_path.is_file():
        return None
    try:
        recovery = importlib.import_module('recover_profile')
    except ImportError as error:
        measure.fail(f'profile evidence is present but recover_profile.py is unavailable: {error}')

    try:
        summary = recovery._json(summary_path)
    except (AttributeError, OSError, TypeError, ValueError,
            recovery.RecoveryError) as error:
        measure.fail(f'profile summary custody validation failed: {error}')
    measure.require(
        isinstance(summary, dict)
        and summary.get('schema') == recovery.PROFILE_SCHEMA
        and summary.get('version') == recovery.VERSION,
        'profile summary schema differs',
    )
    build_path = ROOT / 'build-normal.json'
    measure.require(
        builds['normal']['path'] == str(build_path)
        or Path(builds['normal']['path']).resolve() == build_path.resolve(),
        'profile normal build differs from accepted build',
    )
    try:
        archived_helpers = _profile_helper_archive(
            recovery, ROOT / 'profiling-r1-helper-sources'
        )
        original_gate = recovery._profile_gate(
            ROOT / 'validation' / 'profile-r1.json', archived_helpers['profile']
        )
        profile_builds, _protocol, profile_custody = recovery._build_protocol(
            build_path, summary
        )
        record_inputs = recovery._record_inputs(profile_dir, summary, profile_builds)
    except (AttributeError, KeyError, OSError, TypeError, ValueError,
            recovery.RecoveryError, measure.ProviderMatrixError) as error:
        measure.fail(f'profile custody validation failed: {error}')
    measure.require(
        profile_builds['normal']['receipt_sha256'] == builds['normal']['receipt_sha256']
        and profile_builds['allocator']['receipt_sha256'] == builds['allocator']['receipt_sha256'],
        'profile build receipts differ from accepted builds',
    )

    owned = summary.get('owned_record')
    measure.require(isinstance(owned, dict) and owned.get('status') == 'failed',
                    'original profile failure was not retained honestly')
    script = owned.get('script')
    measure.require(isinstance(script, dict) and script.get('status') == 'failed',
                    'original failed perf-script export is missing')
    measure.require(
        original_gate.get('exit_code') != 0
        and record_inputs['failed_script_status'] == 'failed'
        and record_inputs['record_terminal_status'] == 'pass',
        'original profile/export failure custody is incomplete',
    )

    recovery_summary_path, recovery_dir, recovery_gate_path, recovery_gate = (
        _select_profile_recovery(
            recovery, profile_dir, build_path,
            ROOT / 'validation' / 'profile-r1.json',
            ROOT / 'profiling-r1-helper-sources',
        )
    )
    try:
        recovery_summary = recovery._json(recovery_summary_path)
    except (AttributeError, OSError, TypeError, ValueError,
            recovery.RecoveryError) as error:
        measure.fail(f'{recovery_summary_path}: recovery summary is invalid: {error}')
    measure._exact(
        recovery_summary,
        ('schema', 'version', 'status', 'reason', 'created_utc', 'mode',
         'source_attempt', 'custody', 'command', 'process', 'artifacts',
         'stack_summary', 'interpretation'),
        str(recovery_summary_path.relative_to(ROOT)),
    )
    measure.require(
        recovery_summary['schema'] == recovery.SCHEMA
        and recovery_summary['version'] == recovery.VERSION
        and recovery_summary['status'] == 'pass'
        and recovery_summary['reason'] is None
        and recovery_summary['mode'] == 'retained-perf-data-script-without-cpu-field',
        'profile recovery status or mode differs',
    )

    profile_summary_descriptor = _profile_meta(summary_path, 'profile summary')
    source_attempt = recovery_summary['source_attempt']
    measure.require(
        isinstance(source_attempt, dict)
        and source_attempt.get('profile_summary') == profile_summary_descriptor
        and source_attempt.get('profile_summary_status') == 'failed'
        and source_attempt.get('record') == record_inputs,
        'profile recovery source attempt differs',
    )

    custody = recovery_summary.get('custody')
    measure.require(isinstance(custody, dict), 'profile recovery custody is missing')
    recovery_helper = _historical_helper_descriptor(
        custody.get('recovery_helper'), 'recover_profile.py'
    )
    source_before = custody.get('source_before')
    source_after = custody.get('source_after')
    measure.require(isinstance(source_before, dict) and isinstance(source_after, dict),
                    'profile recovery source custody is missing')
    measure._source_binding(source_before, 'profile recovery custody.source_before')
    measure._source_binding(source_after, 'profile recovery custody.source_after')
    expected_custody = {
        'recovery_helper': recovery_helper,
        'helper_archive': archived_helpers,
        'original_gate': original_gate,
        'build_protocol': profile_custody,
        'source_before': source_before,
        'source_after': source_after,
    }
    measure.require(custody.get('recovery_helper') == recovery_helper,
                    'profile recovery helper changed')
    measure.require(
        _profile_helper_core(custody.get('helper_archive'),
                             'profile recovery custody.helper_archive') == archived_helpers,
                    'profile helper archive changed')
    measure.require(custody.get('original_gate') == original_gate,
                    'original profile gate changed')
    measure.require(custody.get('build_protocol') == profile_custody,
                    'profile build/protocol custody changed')
    measure.require(
        source_before == source_after,
        'profile recovery source changed during postprocessing',
    )
    measure.require(
        expected_custody['source_before'] == custody['source_before']
        and expected_custody['source_after'] == custody['source_after'],
        'profile recovery source receipt differs',
    )

    data_path = Path(record_inputs['perf_data']['path'])
    command = recovery.script_command(data_path, 2)
    measure.require(recovery_summary['command'] == command,
                    'profile recovery command differs')
    process = recovery_summary['process']
    measure.require(isinstance(process, dict), 'profile recovery process is missing')
    measure.require(
        process.get('argv') == command
        and process.get('exit_code') == 0
        and process.get('timed_out') is False
        and process.get('termination') is None
        and process.get('launch_error') is None,
        'profile recovery process did not pass',
    )

    expected_paths = {
        'started': recovery_dir / 'perf-script.started.json',
        'terminal': recovery_dir / 'perf-script.terminal.json',
        'stdout': recovery_dir / 'perf-script.txt',
        'stderr': recovery_dir / 'perf-script.stderr',
        'folded': recovery_dir / 'perf-folded.txt',
        'stack_summary': recovery_dir / 'perf-stack-summary.json',
    }
    artifacts = recovery_summary['artifacts']
    measure.require(isinstance(artifacts, dict)
                    and set(artifacts) == set(expected_paths),
                    'profile recovery artifact inventory differs')
    actual_artifacts = {
        name: _profile_meta(path, f'profile recovery {name}')
        for name, path in expected_paths.items()
    }
    measure.require(artifacts == actual_artifacts,
                    'profile recovery artifact changed')

    started = recovery._json(expected_paths['started'])
    measure._exact(
        started,
        ('schema', 'role', 'argv', 'cwd', 'environment', 'new_session', 'input',
         'recovery_helper', 'source_before', 'started_utc', 'scope'),
        'profile recovery started receipt',
    )
    measure.require(
        started['schema'] == recovery.PROCESS_STARTED_SCHEMA
        and started['role'] == 'owned-perf-script-recovery'
        and started['argv'] == command
        and started['cwd'] == str(recovery.REPO)
        and started['new_session'] is True
        and started['input'] == record_inputs['perf_data']
        and started['recovery_helper'] == recovery_helper
        and started['source_before'] == custody['source_before']
        and started['scope'] == 'reprocess retained perf.data only; no DOCX workload is launched',
        'profile recovery started receipt differs',
    )

    terminal = recovery._json(expected_paths['terminal'])
    measure._exact(
        terminal,
        ('schema', 'role', 'status', 'reason', 'process', 'source_before',
         'source_after', 'source_unchanged', 'finished_utc', 'artifacts', 'input'),
        'profile recovery terminal receipt',
    )
    terminal_artifacts = {
        'perf-script.txt': actual_artifacts['stdout'],
        'perf-script.stderr': actual_artifacts['stderr'],
    }
    measure.require(
        terminal['schema'] == recovery.PROCESS_TERMINAL_SCHEMA
        and terminal['role'] == 'owned-perf-script-recovery'
        and terminal['status'] == 'pass'
        and terminal['reason'] is None
        and terminal['process'] == process
        and terminal['source_before'] == custody['source_before']
        and terminal['source_after'] == custody['source_after']
        and terminal['source_unchanged'] is True
        and terminal['artifacts'] == terminal_artifacts
        and terminal['input'] == record_inputs['perf_data'],
        'profile recovery terminal receipt differs',
    )

    stack_descriptor = recovery_summary['stack_summary']
    measure.require(isinstance(stack_descriptor, dict),
                    'profile recovery stack summary is missing')
    measure.require(
        stack_descriptor.get('path') == actual_artifacts['stack_summary']['path']
        and stack_descriptor.get('bytes') == actual_artifacts['stack_summary']['bytes']
        and stack_descriptor.get('sha256') == actual_artifacts['stack_summary']['sha256'],
        'profile recovery stack summary artifact differs',
    )
    stack_summary = recovery._json(expected_paths['stack_summary'])
    measure.require(isinstance(stack_summary, dict),
                    'profile recovery stack summary is not an object')
    sample_count = stack_summary.get('sample_count')
    total_period = stack_summary.get('total_period')
    measure.require(
        type(sample_count) is int and sample_count > 0
        and type(total_period) is int and total_period > 0
        and stack_descriptor.get('sample_count') == sample_count
        and stack_descriptor.get('total_period') == total_period,
        'profile recovery stack summary counts differ',
    )
    measure.require(expected_paths['folded'].read_text(encoding='utf-8').strip(),
                    'profile recovery folded stack output is empty')
    interpretation = recovery_summary['interpretation']
    measure.require(
        isinstance(interpretation, str)
        and 'retained perf.data' in interpretation
        and 'no workload samples' in interpretation,
        'profile recovery interpretation overclaims workload samples',
    )
    return {
        'status': 'pass',
        'original_status': 'failed',
        'profile_summary': profile_summary_descriptor,
        'recovery_gate': recovery_gate,
        'recovery_summary_path': str(recovery_summary_path),
        'recovery_summary': _profile_meta(recovery_summary_path, 'recovery summary'),
        'sample_count': sample_count,
        'total_period': total_period,
    }


def _verify_observer_correction(builds, profile_result):
    """Authenticate corrected observer counters derived from retained profile files."""

    try:
        correction = importlib.import_module('recover_observer_counters')
    except ImportError as error:
        measure.fail(f'profile observer evidence is present but correction helper is unavailable: {error}')

    profile_dir = ROOT / 'profiling-r1'
    summary_path = profile_dir / 'profile-summary.json'
    build_path = ROOT / 'build-normal.json'
    original_gate_path = ROOT / 'validation' / 'profile-r1.json'
    archive_dir = ROOT / 'profiling-r1-helper-sources'
    try:
        profile_summary = correction._json(summary_path)
        archived_helpers = _profile_helper_archive(correction.recovery, archive_dir)
        original_gate = correction.recovery._profile_gate(
            original_gate_path, archived_helpers['profile']
        )
        profile_builds, _protocol, profile_custody = correction.recovery._build_protocol(
            build_path, profile_summary
        )
        record_inputs = correction.recovery._record_inputs(
            profile_dir, profile_summary, profile_builds
        )
    except (AttributeError, KeyError, OSError, TypeError, ValueError,
            correction.CorrectionError, correction.recovery.RecoveryError,
            measure.ProviderMatrixError) as error:
        measure.fail(f'profile observer custody validation failed: {error}')
    measure.require(
        profile_builds['normal']['receipt_sha256'] == builds['normal']['receipt_sha256']
        and profile_builds['allocator']['receipt_sha256'] == builds['allocator']['receipt_sha256'],
        'profile observer build receipts differ from accepted builds',
    )

    recovery_summary_path = Path(profile_result['recovery_summary_path'])
    recovery_dir = recovery_summary_path.parent
    recovery_script = recovery_dir / 'perf-script.txt'
    correction_candidates = []
    for correction_path in sorted(ROOT.glob('profiling-counter-recovery-*/observer-correction.json')):
        if correction_path.is_symlink() or not correction_path.is_file():
            continue
        try:
            value = correction._json(correction_path)
        except (AttributeError, OSError, TypeError, ValueError,
                correction.CorrectionError) as error:
            measure.fail(f'{correction_path}: observer correction is invalid: {error}')
        if not (isinstance(value, dict)
                and value.get('schema') == correction.SCHEMA
                and value.get('version') == correction.VERSION
                and value.get('status') == 'pass'):
            continue
        recovery_archive = None
        recovery_archive_binding = value.get('custody', {}).get('recovery_helper_archive')
        if isinstance(recovery_archive_binding, dict):
            archive_custody = recovery_archive_binding.get('custody')
            if isinstance(archive_custody, dict) and isinstance(archive_custody.get('path'), str):
                recovery_archive = Path(archive_custody['path']).parent
        expected = _observer_correction_command(
            profile_dir, recovery_summary_path, recovery_script, build_path,
            original_gate_path, archive_dir, correction_path.parent,
            recovery_archive,
        )
        matching_gates = []
        for gate_path in sorted((ROOT / 'validation').glob('profile-counter-recovery-*.json')):
            if gate_path.is_symlink() or not gate_path.is_file():
                continue
            try:
                receipt = read(gate_path)
            except (OSError, ValueError, json.JSONDecodeError) as error:
                measure.fail(f'{gate_path}: observer correction gate is invalid: {error}')
            if (not isinstance(receipt, dict)
                    or receipt.get('exit_code') != 0
                    or receipt.get('source_unchanged') is not True):
                continue
            if receipt.get('argv') != expected:
                continue
            matching_gates.append((gate_path, _passing_postprocess_gate(
                gate_path, expected, gate_path.stem
            )))
        if not matching_gates:
            continue
        measure.require(len(matching_gates) == 1,
                        f'{correction_path}: observer correction gate is ambiguous')
        correction_candidates.append((correction_path, value,
                                     matching_gates[0][0], matching_gates[0][1],
                                     recovery_archive))

    measure.require(correction_candidates,
                    'successful profile observer correction is missing')
    correction_candidates.sort(key=lambda item: _attempt_rank(item[0]))
    highest = _attempt_rank(correction_candidates[-1][0])
    measure.require(
        sum(_attempt_rank(item[0]) == highest for item in correction_candidates) == 1,
        'successful profile observer correction is ambiguous',
    )
    correction_path, correction_value, correction_gate_path, correction_gate, recovery_archive = (
        correction_candidates[-1]
    )

    measure._exact(
        correction_value,
        ('created_utc', 'custody', 'interpretation', 'providers', 'schema',
         'scope', 'source_attempt', 'stack', 'status', 'version'),
        str(correction_path.relative_to(ROOT)),
    )
    measure.require(
        correction_value['schema'] == correction.SCHEMA
        and correction_value['version'] == correction.VERSION
        and correction_value['status'] == 'pass',
        'profile observer correction status differs',
    )

    profile_summary_descriptor = _profile_meta(summary_path, 'profile summary')
    source_attempt = correction_value['source_attempt']
    measure.require(isinstance(source_attempt, dict),
                    'profile observer correction source attempt is missing')
    measure.require(
        source_attempt.get('profile_summary') == profile_summary_descriptor
        and source_attempt.get('profile_gate') == original_gate
        and source_attempt.get('protocol') == profile_custody['protocol']
        and source_attempt.get('build_record') == {
            role: profile_custody['builds'][role]
            for role in correction.canonical_measure.ROLES
        }
        and source_attempt.get('retained_record') == record_inputs,
        'profile observer correction source custody differs',
    )

    stack_recovery = correction._validate_recovery_summary(
        recovery_summary_path, profile_summary_descriptor, record_inputs['perf_data']
    )
    corrected_stack = correction._corrected_stack_summary(
        stack_recovery, recovery_script
    )
    measure.require(source_attempt.get('stack_recovery') == stack_recovery,
                    'profile observer correction stack recovery differs')
    measure.require(source_attempt.get('corrected_stack') == corrected_stack
                    and correction_value['stack'] == corrected_stack,
                    'profile observer correction stack summary differs')

    recovery_archive_binding = correction_value['custody'].get('recovery_helper_archive')
    measure.require(recovery_archive is not None
                    and isinstance(recovery_archive_binding, dict),
                    'profile observer recovery-helper archive is missing')
    expected_recovery_archive = correction._recovery_helper_archive(recovery_archive)
    measure.require(recovery_archive_binding == expected_recovery_archive,
                    'profile observer recovery-helper archive changed')
    archived_recovery_helper = expected_recovery_archive['archived']['recover_profile.py']
    retained_recovery_helper = stack_recovery.get('recovery_helper')
    measure.require(
        isinstance(retained_recovery_helper, dict)
        and retained_recovery_helper.get('bytes') == archived_recovery_helper.get('bytes')
        and retained_recovery_helper.get('sha256') == archived_recovery_helper.get('sha256'),
        'profile observer recovery helper differs from its archive',
    )

    custody = correction_value['custody']
    measure.require(isinstance(custody, dict),
                    'profile observer correction custody is missing')
    correction_helper = _historical_helper_descriptor(
        custody.get('recovery_helper'), 'recover_observer_counters.py',
        allow_unarchived=True,
    )
    measure.require(
        _profile_helper_core(custody.get('helper_archive'),
                             'profile observer custody.helper_archive') == archived_helpers
        and custody.get('build_protocol') == profile_custody
        and custody.get('recovery_helper') == correction_helper
        and custody.get('source_unchanged_during_correction') is True,
        'profile observer correction helper custody differs',
    )
    source_before = custody.get('source_before')
    source_after = custody.get('source_after')
    measure.require(isinstance(source_before, dict) and isinstance(source_after, dict),
                    'profile observer correction source custody is missing')
    measure._source_binding(source_before, 'profile observer correction source_before')
    measure._source_binding(source_after, 'profile observer correction source_after')
    measure.require(source_before == source_after,
                    'profile observer correction source changed')

    providers = correction_value['providers']
    measure.require(isinstance(providers, dict)
                    and set(providers) == set(correction.PROVIDERS),
                    'profile observer correction provider inventory differs')
    old_providers = correction._provider_map(profile_summary)
    for provider in correction.PROVIDERS:
        provider_value = providers[provider]
        measure.require(isinstance(provider_value, dict)
                        and set(provider_value) == {'perf_stat', 'strace'},
                        f'{provider}: observer correction sections differ')
        provider_dir = profile_dir / provider
        raw_paths = {
            'perf_stat': provider_dir / 'perf-stat.csv',
            'strace': provider_dir / 'strace-summary.txt',
        }
        for section, raw_path in raw_paths.items():
            section_value = provider_value[section]
            measure.require(isinstance(section_value, dict)
                            and set(section_value) == {
                                'raw', 'old_profile_summary', 'corrected', 'interpretation'
                            },
                            f'{provider}.{section}: observer correction fields differ')
            raw = correction._file_meta(raw_path)
            old_section = 'perf' if section == 'perf_stat' else 'strace'
            old = {
                'path': str(summary_path),
                'sha256': profile_summary_descriptor['sha256'],
                'parsed': old_providers[provider][old_section].get('parsed'),
            }
            if section == 'perf_stat':
                expected_corrected = correction.parse_perf_stat_corrected(
                    raw_path.read_text(encoding='utf-8', errors='replace')
                )
            else:
                expected_corrected = correction.parse_strace_corrected(
                    raw_path.read_text(encoding='utf-8', errors='replace')
                )
            measure.require(
                section_value['raw'] == raw
                and section_value['old_profile_summary'] == old
                and section_value['corrected'] == expected_corrected,
                f'{provider}.{section}: corrected observer data changed',
            )

    interpretation = correction_value['interpretation']
    measure.require(
        isinstance(interpretation, str)
        and 'immutable raw observer files' in interpretation
        and 'no cache-miss' in interpretation,
        'profile observer correction interpretation overclaims evidence',
    )
    return {
        'status': 'pass',
        'correction_gate': correction_gate,
        'correction_summary': _profile_meta(
            correction_path, 'observer correction summary'
        ),
    }


def verify_evidence():
    builds = measure.load_builds()
    protocol, protocol_hash = measure._load_protocol(builds)
    measure.require(protocol == measure.protocol_value(builds), 'protocol/build mismatch')
    accepted = read(ROOT / 'accepted-evidence.json')
    measure.require(isinstance(accepted, dict), 'accepted evidence is not an object')
    accepted_schema = accepted.get('schema')
    measure.require(
        accepted_schema == 'docx-edit-provider-accepted-v1',
        'accepted evidence schema differs',
    )
    formal_samples = sum(
        int(item['samples']) for item in measure.formal_inventory()
    )
    pilot_samples = sum(
        int(item['samples']) for item in measure.formal_inventory(pilot=True)
    )
    measure.require(accepted.get('formal_samples') == formal_samples,
                    'accepted formal sample inventory differs')
    measure.require(accepted.get('pilot_samples') == pilot_samples,
                    'accepted pilot sample inventory differs')
    for pilot, key in ((True, 'pilot_attempt'), (False, 'formal_attempt')):
        measure.verify(accepted[key], pilot=pilot)
    cold_summary = _verify_cold_acceptance(accepted, protocol_hash, builds)
    profile_summary = _verify_profile_recovery(builds)
    observer_summary = (
        _verify_observer_correction(builds, profile_summary)
        if profile_summary is not None else None
    )

    gates = read(ROOT / 'final-gates.json')
    final_commands = _validate_final_gates(gates)
    manifest_labels = {label for label, _ in final_commands}
    validation_json = {
        path.stem
        for path in (ROOT / 'validation').glob('*.json')
        if not path.name.endswith('.started.json')
    }
    measure.require(manifest_labels <= validation_json,
                    'final-gate receipt is missing from validation')
    gate_bindings = {}
    for label, command in final_commands:
        path = ROOT / 'validation' / f'{label}.json'
        gate_bindings[label] = _validate_gate_receipt(
            label, command, path, builds['normal']['source']
        )

    tested = read(ROOT / 'helper-test-custody.json')
    expected_helpers = _expected_helpers()
    measure.require(isinstance(tested, dict), 'helper custody is not an object')
    measure.require(set(tested.get('helpers', {})) == expected_helpers,
                    'helper custody inventory is incomplete or unexpected')
    measure.require(tested['helpers'] == {name: meta(ROOT / name) for name in expected_helpers},
                    'tested helper changed')
    measure.require(set(measure.DRIVER_FILES).issubset(tested['helpers']),
                    'capture helpers missing from tested inventory')
    helper_label = tested.get('label')
    measure.require(isinstance(helper_label, str) and helper_label in gates['commands'],
                    'helper test label is not a final gate')
    measure.require(
        tuple(gates['commands'][helper_label]) == tuple(_expected_final_commands()[5]),
        'helper custody label does not select the Python helper gate',
    )
    measure.require(tested['gate'] == meta(ROOT / 'validation' / (helper_label + '.json')),
                    'helper test gate changed')
    helper_binding = gate_bindings[helper_label]
    measure.require(tested.get('source') == helper_binding['source'],
                    'helper test gate source differs from its custody source')

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
    result = {'status': 'pass', 'formal_samples': formal_samples, 'pilot_samples': pilot_samples,
            'protocol_sha256': protocol_hash, 'gates': sorted(gates['commands']),
            'cleanup_receipt': meta(ROOT / 'cleanup.json'),
            'cleanup_verification_schema': cleanup_proof['schema']}
    if cold_summary is not None:
        result['cold'] = cold_summary
    if profile_summary is not None:
        result['profile_recovery'] = profile_summary
    if observer_summary is not None:
        result['profile_observer_correction'] = observer_summary
    return result


def main():
    if sys.argv[1:] not in [['create'], ['verify']]:
        raise SystemExit('usage: verify_bundle.py create|verify')
    result = {'schema': 'docx-edit-provider-bundle-v1',
              'evidence': verify_evidence(), 'files': inventory()}
    path = ROOT / 'seal.json'
    if sys.argv[1] == 'create':
        write(path, result)
    elif read(path) != result:
        raise RuntimeError('bundle inventory or evidence changed')
    print(json.dumps({'status': 'pass', 'files': len(result['files']), **meta(path)}, sort_keys=True))


if __name__ == '__main__':
    main()
