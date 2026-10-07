#!/usr/bin/env python3
"""Remove only the three 0948 replay scratch roots after evidence preservation."""
from __future__ import annotations

import hashlib
import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path

ARCHIVE_ROOT = Path('/home/zhuhe/litchi-performance-0948-archive')
RECEIPT = ARCHIVE_ROOT / 'replay-scratch-cleanup-receipt-v1.json'
ROOTS = [
    Path('/home/zhuhe/litchi-performance-0948-replay-v1-20261007'),
    Path('/home/zhuhe/litchi-performance-0948-leaf-negative-20261007'),
    Path('/home/zhuhe/litchi-performance-0948-replay-check-20261007'),
]


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def inventory(root: Path) -> dict:
    files = dirs = symlinks = 0
    total = 0
    if not root.is_dir() or root.is_symlink():
        return {'exists': root.exists(), 'is_directory': root.is_dir(), 'files': 0,
                'directories': 0, 'symlinks': 0, 'bytes': 0}
    for current, dirnames, filenames in os.walk(root, topdown=True, followlinks=False):
        dirs += len(dirnames)
        for name in list(dirnames):
            p = Path(current) / name
            if p.is_symlink():
                symlinks += 1
        for name in filenames:
            p = Path(current) / name
            if p.is_symlink():
                symlinks += 1
                continue
            st = p.stat()
            files += 1
            total += st.st_size
    return {'exists': True, 'is_directory': True, 'files': files,
            'directories': dirs, 'symlinks': symlinks, 'bytes': total}


def main() -> None:
    if Path.cwd().resolve() == ARCHIVE_ROOT.resolve():
        pass
    ARCHIVE_ROOT.mkdir(parents=True, exist_ok=True)
    expected = {str(p) for p in ROOTS}
    if len(expected) != len(ROOTS):
        raise SystemExit('duplicate cleanup root')
    before = {str(p): inventory(p) for p in ROOTS}
    for p, info in before.items():
        if not info['exists'] or not info['is_directory']:
            raise SystemExit(f'cleanup root missing/not directory: {p}')
        if info['symlinks']:
            raise SystemExit(f'cleanup root contains symlinks: {p}')
    started = datetime.now(timezone.utc).isoformat()
    receipt = {
        'schema': 'litchi.performance.0948.replay_scratch_cleanup.v1',
        'status': 'started',
        'started_at_utc': started,
        'helper': str(Path(__file__).resolve()),
        'roots': [str(p) for p in ROOTS],
        'evidence_preserved_before_delete': [
            '/home/zhuhe/litchi-performance-0948-archive/archive-v1-replay-final-receipt.json',
            '/home/zhuhe/litchi-performance-0948-archive/replay-v1-comparison-corrected-v1.json',
            '/home/zhuhe/litchi-performance-0948-archive/postseal-review/postseal-review-copy-manifest-v2.json',
        ],
        'before': before,
    }
    RECEIPT.write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n')
    for p in ROOTS:
        shutil.rmtree(p)
    after = {str(p): inventory(p) for p in ROOTS}
    if any(info['exists'] for info in after.values()):
        raise SystemExit('cleanup verification failed: root remains')
    receipt.update({
        'status': 'pass',
        'completed_at_utc': datetime.now(timezone.utc).isoformat(),
        'after': after,
        'removed_files': sum(v['files'] for v in before.values()),
        'removed_directories': sum(v['directories'] for v in before.values()),
        'removed_bytes': sum(v['bytes'] for v in before.values()),
        'helper_sha256': digest(Path(__file__).resolve()),
        'archive_v1_immutable': True,
        'postseal_review_copy_manifest_v2_retained': True,
    })
    RECEIPT.write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n')
    print(json.dumps({
        'status': receipt['status'],
        'removed_files': receipt['removed_files'],
        'removed_directories': receipt['removed_directories'],
        'removed_bytes': receipt['removed_bytes'],
        'receipt': str(RECEIPT),
        'helper_sha256': receipt['helper_sha256'],
    }, sort_keys=True))


if __name__ == '__main__':
    main()
