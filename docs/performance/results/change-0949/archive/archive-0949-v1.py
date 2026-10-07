#!/usr/bin/env python3
"""Create a manifest and, after explicit GO, a low-compression 0949 archive.

The packet is evidence only: source/build trees and binaries remain external
0948 dependencies. This helper never deletes the packet or its dependencies.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PACKET_DEFAULT = Path('/home/zhuhe/litchi-performance-0949')
ARCHIVE_DEFAULT = Path('/home/zhuhe/litchi-performance-0949-archive')
SCHEMA = 'litchi.performance.0949.external-archive-manifest.v1'


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def descriptor(path: Path, *, relative: str | None = None) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        raise RuntimeError(f'not a regular file: {path}')
    st = path.stat()
    return {
        'path': relative if relative is not None else str(path.resolve()),
        'bytes': st.st_size,
        'mode': stat.S_IMODE(st.st_mode),
        'executable': bool(st.st_mode & 0o111),
        'sha256': sha256(path),
    }


def dependency_descriptor(archive_root: Path) -> dict[str, Any]:
    deps = {
        '0948_archive': archive_root.parent / 'litchi-performance-0948-archive' / 'archive-0948-v1.tar.zst',
        '0948_manifest': archive_root.parent / 'litchi-performance-0948-archive' / 'manifest-v1.json',
        'candidate_source_manifest': archive_root.parent / 'litchi-performance-0948' / 'candidate-build-manifest-v3.json',
        'baseline_source_manifest': archive_root.parent / 'litchi-performance-0948' / 'baseline-source-manifest-v1.json',
        'baseline_allocator_binary': archive_root.parent / 'litchi-performance-0948' / 'binaries' / 'baseline-allocator',
        'candidate_allocator_binary': archive_root.parent / 'litchi-performance-0948' / 'binaries' / 'candidate-v3-preserved' / 'litchi-perf-baseline-alloc',
    }
    out: dict[str, Any] = {
        'schema': 'litchi.performance.0949.external-dependency-binding.v1',
        'status': 'prepared_external_dependency_only',
        'note': '0949 packet intentionally contains no duplicated source tree or benchmark binary; the reconciler reads the external 0948 candidate source snapshot.',
        'dependencies': {},
    }
    for name, path in deps.items():
        out['dependencies'][name] = descriptor(path)
    out['source_roots'] = {
        'candidate_for_reconciliation': str((archive_root.parent / 'litchi-performance-0948' / 'candidate-build-v3').resolve()),
        'baseline_build_source': str((archive_root.parent / 'litchi-performance-0948' / 'baseline-source').resolve()),
    }
    out['required_replay_inputs'] = [
        'candidate_source_manifest', 'candidate_allocator_binary', 'baseline_allocator_binary',
        '0948_archive', '0948_manifest',
    ]
    return out


def collect(packet: Path) -> tuple[list[dict[str, Any]], list[str], list[str]]:
    if not packet.is_dir() or packet.is_symlink():
        raise RuntimeError(f'packet is not a real directory: {packet}')
    entries: list[dict[str, Any]] = []
    excluded: list[str] = []
    elf_files: list[str] = []
    for current, dirs, files in os.walk(packet, topdown=True, followlinks=False):
        current_path = Path(current)
        kept_dirs: list[str] = []
        for name in sorted(dirs):
            path = current_path / name
            rel = path.relative_to(packet).as_posix()
            if path.is_symlink():
                raise RuntimeError(f'symlink directory is not allowed: {path}')
            if name == '__pycache__':
                excluded.append(rel + '/')
                continue
            kept_dirs.append(name)
        dirs[:] = kept_dirs
        for name in sorted(files):
            path = current_path / name
            rel = path.relative_to(packet).as_posix()
            if path.is_symlink():
                raise RuntimeError(f'symlink file is not allowed: {path}')
            if path.suffix == '.pyc':
                excluded.append(rel)
                continue
            if not path.is_file():
                raise RuntimeError(f'non-regular packet entry: {path}')
            item = descriptor(path, relative=rel)
            entries.append(item)
            with path.open('rb') as stream:
                magic = stream.read(4)
            if magic == b'\x7fELF':
                elf_files.append(rel)
    entries.sort(key=lambda item: item['path'])
    excluded.sort()
    elf_files.sort()
    return entries, excluded, elf_files


def make_manifest(packet: Path, archive_root: Path, *, status: str, dependency: dict[str, Any]) -> dict[str, Any]:
    entries, excluded, elf_files = collect(packet)
    return {
        'schema': SCHEMA,
        'status': status,
        'generated_utc': datetime.now(timezone.utc).isoformat(),
        'packet_root': str(packet.resolve()),
        'entry_count': len(entries),
        'packet_bytes': sum(int(item['bytes']) for item in entries),
        'entries': entries,
        'excluded_generated_entries': excluded,
        'excluded_policy': ['directories named __pycache__', 'files ending in .pyc'],
        'elf_files': elf_files,
        'elf_file_count': len(elf_files),
        'final_selection': {
            'attribution': 'attribution-final-v1',
            'attribution_tool': 'tools/attribute_heaptrack_0949_v1.py',
            'decoder': 'tools/decode_heaptrack_0949_v1.py',
            'attribution_config': 'tools/attribution-config-root-v1.json',
            'reconciliation': 'review/xml-multiwindow-source-profile-reconciliation-v2.py',
        },
        'preserved_exploratory_history': ['attribution-v1', 'attribution-v2', 'raw-decoder-v1', 'heaptrack-exports-v1', 'heaptrack-exports-v2'],
        'dependency_binding': dependency,
        'archive_policy': {
            'source_duplication': False,
            'binary_duplication': False,
            'compression': 'zstd level 3 after deterministic GNU tar stream',
            'source_deletion': False,
            'cleanup_requires_explicit_GO': True,
        },
    }


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')


def sha_file(path: Path) -> str:
    return sha256(path)


def seal(packet: Path, archive_root: Path, manifest_path: Path) -> None:
    if not manifest_path.is_file():
        raise RuntimeError(f'manifest missing: {manifest_path}')
    manifest = json.loads(manifest_path.read_text())
    if manifest.get('status') != 'final_sealed_input':
        raise RuntimeError('seal requires a final_sealed_input manifest')
    entries = manifest['entries']
    selection = archive_root / 'selection-v1.list'
    selection.write_bytes(b''.join(str(item['path']).encode() + b'\0' for item in entries))
    tar_path = archive_root / 'archive-0949-v1.tar'
    compressed_path = archive_root / 'archive-0949-v1.tar.zst'
    for path in (tar_path, compressed_path):
        if path.exists() or path.is_symlink():
            raise RuntimeError(f'refusing to overwrite archive artifact: {path}')
    subprocess.run([
        'tar', '--create', '--sort=name', '--mtime=@0', '--owner=0', '--group=0', '--numeric-owner',
        '--format=pax', '--file', str(tar_path), '--null', '--files-from', str(selection),
    ], cwd=packet, check=True)
    temp_compressed = archive_root / '.archive-0949-v1.tar.zst.tmp'
    if temp_compressed.exists():
        temp_compressed.unlink()
    subprocess.run(['zstd', '-3', '--no-progress', '--force', str(tar_path), '-o', str(temp_compressed)], check=True)
    temp_compressed.replace(compressed_path)
    tar_desc = descriptor(tar_path)
    tar_desc['retained'] = False
    tar_desc['deleted_after_compression'] = True
    compressed_desc = descriptor(compressed_path)
    receipt = {
        'schema': 'litchi.performance.0949.external-archive-receipt.v1',
        'status': 'sealed',
        'packet_manifest': descriptor(manifest_path),
        'selection': descriptor(selection),
        'uncompressed_tar': tar_desc,
        'compressed_archive': compressed_desc,
        'regeneration': f'python3 -B {Path(__file__).resolve()} --seal --go --packet {packet} --archive-root {archive_root} --manifest {manifest_path}',
        'source_cleanup_performed': False,
        'moved_replay_required_before_cleanup': True,
    }
    write_json(archive_root / 'archive-receipt-v1.json', receipt)
    tar_path.unlink()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--packet', type=Path, default=PACKET_DEFAULT)
    parser.add_argument('--archive-root', type=Path, default=ARCHIVE_DEFAULT)
    parser.add_argument('--manifest', type=Path, default=None)
    parser.add_argument('--manifest-only', action='store_true')
    parser.add_argument('--seal', action='store_true', help='explicitly seal archive; requires final_sealed_input manifest')
    parser.add_argument('--go', action='store_true', help='required with --seal as explicit coordinator authorization')
    args = parser.parse_args()
    packet = args.packet.resolve()
    archive_root = args.archive_root.resolve()
    archive_root.mkdir(parents=True, exist_ok=True)
    dependency_path = archive_root / 'dependency-0949-v1.json'
    dependency = dependency_descriptor(archive_root)
    write_json(dependency_path, dependency)
    manifest_path = (args.manifest or archive_root / ('manifest-v1.json' if args.seal else 'manifest-preseal-v1.json')).resolve()
    if args.seal and not args.go:
        raise SystemExit('--seal requires --go; no archive was created')
    status = 'final_sealed_input' if args.seal else 'preseal_not_sealed'
    manifest = make_manifest(packet, archive_root, status=status, dependency=descriptor(dependency_path))
    write_json(manifest_path, manifest)
    if args.seal:
        seal(packet, archive_root, manifest_path)
    print(json.dumps({
        'status': status,
        'manifest': str(manifest_path),
        'manifest_sha256': sha_file(manifest_path),
        'entries': manifest['entry_count'],
        'packet_bytes': manifest['packet_bytes'],
        'elf_files': manifest['elf_file_count'],
        'archive_sealed': bool(args.seal),
    }, sort_keys=True))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
