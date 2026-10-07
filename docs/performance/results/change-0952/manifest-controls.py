#!/usr/bin/env python3
"""Actual extraction-manifest verifier and fail-closed controls for 0952."""
from __future__ import annotations
import copy
import hashlib
import json
import os
import shutil
import stat
import tempfile
from pathlib import Path
from typing import Any

ARCHIVE = Path('/home/zhuhe/litchi-performance-0952-archive')
MANIFEST = ARCHIVE / 'manifest-v1.json'
EXTRACT = Path('/home/zhuhe/litchi-performance-0952-replay-v1')

class ManifestError(RuntimeError):
    pass

def digest(path: Path) -> tuple[int, int, str]:
    if path.is_symlink() or not path.is_file():
        raise ManifestError(f'not a regular file: {path}')
    h = hashlib.sha256(); size = 0
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1 << 20), b''):
            h.update(block); size += len(block)
    return size, stat.S_IMODE(path.stat().st_mode), h.hexdigest()

def safe_rel(raw: Any) -> str:
    if not isinstance(raw, str) or not raw or Path(raw).is_absolute():
        raise ManifestError('manifest path is not relative')
    p = Path(raw)
    if any(part in ('', '.', '..') for part in p.parts):
        raise ManifestError('manifest path escapes root')
    return p.as_posix()

def verify_manifest(root: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    entries = manifest.get('entries')
    if not isinstance(entries, list):
        raise ManifestError('entries missing')
    expected: dict[str, dict[str, Any]] = {}
    for entry in entries:
        if not isinstance(entry, dict): raise ManifestError('entry is not object')
        rel = safe_rel(entry.get('path'))
        if rel in expected: raise ManifestError(f'duplicate entry: {rel}')
        for key in ('bytes','mode','sha256','executable'):
            if key not in entry: raise ManifestError(f'missing {key}: {rel}')
        expected[rel] = entry
    if any(p.is_symlink() for p in root.rglob('*')):
        raise ManifestError('symlink present in extracted root')
    observed: dict[str, Path] = {}
    for p in root.rglob('*'):
        if p.is_file(): observed[p.relative_to(root).as_posix()] = p
    if set(observed) != set(expected):
        raise ManifestError(f'file set differs: missing={sorted(set(expected)-set(observed))} extra={sorted(set(observed)-set(expected))}')
    total = 0
    for rel, entry in expected.items():
        p = observed[rel]; size, mode, sha = digest(p); total += size
        if size != entry['bytes'] or mode != entry['mode'] or sha != entry['sha256']:
            raise ManifestError(f'descriptor differs: {rel}')
        if bool(p.stat().st_mode & 0o111) != bool(entry['executable']):
            raise ManifestError(f'executable bit differs: {rel}')
    if total != manifest.get('logical_bytes') or len(expected) != manifest.get('entry_count'):
        raise ManifestError('census differs')
    return {'entries': len(expected), 'logical_bytes': total}

def hardlink_tree(dst: Path) -> None:
    dst.mkdir(parents=True)
    for src in EXTRACT.rglob('*'):
        rel = src.relative_to(EXTRACT); out = dst / rel
        if src.is_dir(): out.mkdir()
        elif src.is_file(): os.link(src, out)
        else: raise RuntimeError(f'unexpected source: {src}')

def reject(name: str, mutate, manifest: dict[str, Any]) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix='0952-manifest-control-', dir=ARCHIVE) as td:
        root = Path(td) / 'root'; hardlink_tree(root)
        candidate = copy.deepcopy(manifest)
        mutate(root, candidate)
        try:
            verify_manifest(root, candidate)
        except ManifestError as exc:
            return {'name': name, 'rejected': True, 'reason': str(exc)}
        raise RuntimeError(f'control accepted: {name}')

def main() -> int:
    manifest = json.loads(MANIFEST.read_text())
    positive = verify_manifest(EXTRACT, manifest)
    controls = [
      reject('missing-file', lambda root,m: (root / m['entries'][0]['path']).unlink(), manifest),
      reject('extra-file', lambda root,m: (root / 'unexpected-extra.bin').write_bytes(b'x'), manifest),
      reject('symlink-file', lambda root,m: (root / m['entries'][0]['path']).unlink() or (root / m['entries'][0]['path']).symlink_to('/etc/hosts'), manifest),
      reject('symlink-directory', lambda root,m: (root / 'fake-dir').symlink_to('/tmp', target_is_directory=True), manifest),
      reject('wrong-sha', lambda root,m: m['entries'][0].__setitem__('sha256', '0'*64), manifest),
      reject('wrong-mode', lambda root,m: m['entries'][0].__setitem__('mode', 0), manifest),
      reject('duplicate-entry', lambda root,m: m['entries'].append(copy.deepcopy(m['entries'][0])), manifest),
      reject('path-escape', lambda root,m: m['entries'][0].__setitem__('path', '../outside'), manifest),
    ]
    out = {'schema':'litchi.performance.0952.relocated-manifest-controls.v1',
           'status':'pass', 'manifest':str(MANIFEST), 'extraction':str(EXTRACT),
           'positive':positive, 'controls':controls,
           'all_rejected':all(x['rejected'] for x in controls),
           'scope':'Manifest custody controls only; no workload or perf operation.'}
    print(json.dumps(out, indent=2, sort_keys=True))
    return 0

if __name__ == '__main__': raise SystemExit(main())
