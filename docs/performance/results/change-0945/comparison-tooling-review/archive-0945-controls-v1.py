#!/usr/bin/env python3
"""Offline controls for the frozen 0945 archive selector.

This imports the frozen archive script, redirects its roots in memory, and
uses only a temporary fake source/work tree. It never invokes archive_mode,
verify_mode, plan_mode, or cleanup_mode against the real root.
"""
from __future__ import annotations
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import tempfile
from datetime import datetime, timezone
from typing import Any, Callable

SCRIPT = Path('/home/zhuhe/litchi-performance-0945/archive-0945-v1.py')
WORK = Path('/home/zhuhe/litchi-performance-0945-archive')
RECEIPT = WORK / 'comparison-tooling-review' / 'archive-0945-controls-v1.json'


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1 << 20), b''):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    stat = path.stat()
    return {'path': str(path), 'bytes': stat.st_size, 'sha256': sha(path)}


def write(path: Path, data: bytes = b'x') -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def expect_archive_error(label: str, fn: Callable[[], Any], checks: list[dict[str, Any]]) -> None:
    try:
        fn()
    except module.ArchiveError as error:  # type: ignore[name-defined]
        checks.append({'id': label, 'status': 'pass', 'error': str(error)})
    else:
        checks.append({'id': label, 'status': 'fail', 'error': 'expected ArchiveError was not raised'})


def load_module():
    spec = importlib.util.spec_from_file_location('archive_0945_frozen_controls', SCRIPT)
    if spec is None or spec.loader is None:
        raise RuntimeError('could not load frozen archive script')
    loaded = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(loaded)
    return loaded


module = load_module()
script_sha_before = sha(SCRIPT)
checks: list[dict[str, Any]] = []
status = 'pass'
errors: list[str] = []
fake_parent = Path(tempfile.mkdtemp(prefix='archive-0945-controls-', dir=WORK))
fake_source = fake_parent / 'fake-source'
fake_workspace = fake_parent / 'fake-work'
fake_source.mkdir()
fake_workspace.mkdir()

try:
    module.SOURCE_ROOT = fake_source.resolve()
    module.WORK_ROOT = fake_workspace.resolve()
    module.ROOT_NAME = module.SOURCE_ROOT.name
    module.BINARY_DESCRIPTOR = fake_workspace / 'binary-descriptor.json'

    # Exact cache roots have one discarded file each; each build cache also
    # contains its one explicitly allowlisted release binary.
    for directory in sorted(module.EXCLUDED_DIRECTORIES, key=Path.as_posix):
        write(fake_source / directory / 'discard.bin', b'discard')
    for allowed in sorted(module.ALLOWLISTED_TARGET_FILES, key=Path.as_posix):
        write(fake_source / allowed, b'allowed-release')
    # Names beginning with target- elsewhere are ordinary source/corpus paths.
    preserved = [
        Path('target-unlisted-top/keep.bin'),
        Path('crates/example/source/fuzz/target-corpus/input.bin'),
        Path('crates/example/source/target-generated/keep.bin'),
        Path('quality-v2/candidate-source/crates/fuzz/target-samples/case.bin'),
    ]
    for relative in preserved:
        write(fake_source / relative, b'preserve')
    write(fake_source / 'ordinary.txt', b'ordinary')

    selected = {relative.as_posix() for relative, _path in module.iter_selected_files()}
    expected_allowed = {path.as_posix() for path in module.ALLOWLISTED_TARGET_FILES}
    checks.append({'id': 'exact-cache-exclusions', 'status': 'pass' if all(
        path not in selected for path in [
            *(f'{directory.as_posix()}/discard.bin' for directory in module.EXCLUDED_DIRECTORIES),
        ]) else 'fail', 'selected_count': len(selected)})
    checks.append({'id': 'all-three-allowlisted-release-files-selected', 'status': 'pass' if expected_allowed <= selected else 'fail', 'expected': sorted(expected_allowed)})
    checks.append({'id': 'nested-target-names-preserved', 'status': 'pass' if all(path.as_posix() in selected for path in preserved) else 'fail', 'paths': [path.as_posix() for path in preserved]})
    checks.append({'id': 'unlisted-target-top-level-preserved', 'status': 'pass' if 'target-unlisted-top/keep.bin' in selected else 'fail'})

    # Both file and directory symlinks must fail closed; remove each before
    # continuing so no symlink reaches the fake cleanup test.
    symlink_file = fake_source / 'symlink-file'
    symlink_file.symlink_to(fake_source / 'ordinary.txt')
    expect_archive_error('symlink-file-rejected', lambda: list(module.iter_selected_files()), checks)
    symlink_file.unlink()
    symlink_dir = fake_source / 'symlink-directory'
    symlink_dir.symlink_to(fake_source / 'crates', target_is_directory=True)
    expect_archive_error('symlink-directory-rejected', lambda: list(module.iter_selected_files()), checks)
    symlink_dir.unlink()

    # Missing and mismatched frozen controls fail closed.
    control = fake_source / 'control.json'
    write(control, b'control')
    module.FROZEN_CONTROL_FILES = {Path('control.json'): {'bytes': control.stat().st_size, 'sha256': sha(control)}}
    module.verify_frozen_control_files()
    control.unlink()
    expect_archive_error('missing-frozen-control-rejected', module.verify_frozen_control_files, checks)
    write(control, b'changed')
    expect_archive_error('mismatched-frozen-control-rejected', module.verify_frozen_control_files, checks)
    control.write_bytes(b'control')

    # Missing and malformed/mismatched binary descriptors fail closed.
    module.BINARY_DESCRIPTOR.unlink(missing_ok=True)
    expect_archive_error('missing-binary-descriptor-rejected', module.expected_binary_descriptors, checks)
    # Create all five bound binaries and a correct descriptor.
    inventory = [Path('binaries/baseline-native'), Path('binaries/baseline-allocator')]
    all_binary_paths = sorted(module.BOUND_BINARY_PATHS, key=Path.as_posix)
    for relative in all_binary_paths:
        write(fake_source / relative, b'binary-' + relative.as_posix().encode())
    binary_entries = []
    for relative in all_binary_paths:
        path = fake_source / relative
        binary_entries.append({'path': relative.as_posix(), 'bytes': path.stat().st_size, 'sha256': sha(path)})
    module.BINARY_DESCRIPTOR.write_text(json.dumps({
        'schema': module.BINARY_DESCRIPTOR_SCHEMA,
        'source_root': str(fake_source),
        'binaries': binary_entries,
    }) + '\n', encoding='utf-8')
    module.expected_binary_descriptors()
    checks.append({'id': 'complete-binary-descriptor-accepted', 'status': 'pass'})
    first_binary = fake_source / all_binary_paths[0]
    first_binary.write_bytes(b'changed-binary')
    expect_archive_error('mismatched-binary-rejected', module.expected_binary_descriptors, checks)
    first_binary.write_bytes(b'binary-' + all_binary_paths[0].as_posix().encode())

    # Source manifest rows and quality input copies reject missing/mismatched
    # selected records, without consulting archive output.
    manifest_path = fake_source / 'manifest.json'
    plain = fake_source / 'ordinary.txt'
    manifest_path.write_text(json.dumps({'file_count': 1, 'bytes': plain.stat().st_size, 'files': [
        {'path': 'ordinary.txt', 'bytes': plain.stat().st_size, 'sha256': sha(plain)},
    ]}) + '\n', encoding='utf-8')
    module.SOURCE_MANIFESTS = ((Path('manifest.json'), Path('.')),)
    module.QUALITY_SOURCE_MANIFEST_COPIES = ()
    expect_archive_error('missing-source-manifest-selection-rejected', lambda: module.verify_source_manifests({}), checks)
    expect_archive_error('mismatched-source-manifest-selection-rejected', lambda: module.verify_source_manifests({'ordinary.txt': {'bytes': plain.stat().st_size, 'sha256': '0' * 64}}), checks)
    origin = fake_source / 'origin.bin'; write(origin, b'origin')
    quality_copy = fake_source / 'q/candidate-source/fixture.bin'; write(quality_copy, b'wrong')
    origin_descriptor = {'path': str(origin), 'bytes': origin.stat().st_size, 'sha256': sha(origin)}
    copy_descriptor = {'path': 'fixture.bin', 'bytes': origin.stat().st_size, 'sha256': sha(origin)}
    quality_item = {'kind': 'control', 'relative_destination': 'fixture.bin', 'origin': origin_descriptor, 'copies': [{'role': 'candidate', 'destination': copy_descriptor}]}
    expect_archive_error('mismatched-quality-input-copy-rejected', lambda: module.verify_standard_quality_inputs([quality_item], None, label='fake-quality', quality_root=Path('q')), checks)

    # Fake cleanup exercises exact containment and preserves all three release
    # binaries plus ordinary target-* source/fuzz names.
    outside = fake_parent / 'outside.txt'; write(outside, b'outside')
    expect_archive_error('cleanup-outside-source-rejected', lambda: module.remove_path(outside), checks)
    removed = module.cleanup_named_directories()
    module.assert_post_cleanup()
    for directory in module.EXCLUDED_DIRECTORIES:
        if directory in {Path('quality/target-baseline'), Path('quality/target-candidate'), Path('quality-v2/target-baseline'), Path('quality-v2/target-candidate')}:
            if (fake_source / directory).exists():
                raise AssertionError(f'quality cache survived cleanup: {directory}')
    for allowed in module.ALLOWLISTED_TARGET_FILES:
        if not (fake_source / allowed).is_file():
            raise AssertionError(f'allowlisted binary removed: {allowed}')
    for relative in preserved:
        if not (fake_source / relative).is_file():
            raise AssertionError(f'preserved target-like source removed: {relative}')
    checks.append({'id': 'fake-cleanup-exact-and-contained', 'status': 'pass', 'removed_count': len(removed), 'removed_paths': sorted(removed)})

    if any(item['status'] != 'pass' for item in checks):
        raise AssertionError('one or more control checks failed')
except Exception as error:
    status = 'fail'
    errors.append(f'{type(error).__name__}: {error}')
finally:
    # Only this harness-owned fake tree is removed. The real source/work roots
    # are never passed to archive or cleanup modes.
    if fake_parent.exists():
        shutil.rmtree(fake_parent)

script_sha_after = sha(SCRIPT)
if script_sha_after != script_sha_before:
    status = 'fail'
    errors.append('frozen archive script changed during controls')
RECEIPT.parent.mkdir(parents=True, exist_ok=True)
if RECEIPT.exists():
    raise RuntimeError(f'refusing to replace receipt: {RECEIPT}')
receipt = {
    'schema': 'litchi.performance.0945.archive-controls-review.v1',
    'status': status,
    'created_utc': datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
    'frozen_script': {'path': str(SCRIPT), 'bytes': SCRIPT.stat().st_size, 'sha256_before': script_sha_before, 'sha256_after': script_sha_after, 'expected_sha256': 'a61c6140b83f5fd1cd72453940eba8bfbec8f40f7195242f9f0c65c3b3774296'},
    'command_scope': {'real_archive_invoked': False, 'real_verify_invoked': False, 'real_cleanup_invoked': False, 'fake_source_root': str(fake_source), 'fake_tree_removed_after_tests': not fake_parent.exists()},
    'checks': checks,
    'errors': errors,
    'controls': ['exact top-level cache exclusions', 'all three allowlisted release files', 'nested source/fuzz target-* preservation', 'file and directory symlink rejection', 'cleanup containment', 'missing and mismatched frozen controls', 'missing/mismatched binary descriptor', 'missing/mismatched source manifest selection', 'mismatched quality input copy'],
    'regeneration_command': '/usr/bin/python3 -B /home/zhuhe/litchi-performance-0945/comparison-tooling-review/archive-0945-controls-v1.py',
}
with RECEIPT.open('x', encoding='utf-8') as stream:
    json.dump(receipt, stream, indent=2, sort_keys=True)
    stream.write('\n')
    stream.flush()
    os.fsync(stream.fileno())
print(json.dumps({'status': status, 'receipt': str(RECEIPT), 'checks': len(checks), 'errors': len(errors), 'receipt_sha256': sha(RECEIPT)}, sort_keys=True))
if status != 'pass':
    raise SystemExit(1)
