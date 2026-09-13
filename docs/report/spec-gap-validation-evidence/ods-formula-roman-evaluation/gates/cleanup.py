#!/usr/bin/env python3
"""Remove this batch's disposable roots after final review and capture close."""
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tarfile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[4]
OWNED = [Path('/home/zhuhe/code/litchi-roman-' + suffix)
         for suffix in ['worktree', 'target', 'tmp']]
RECOVERY = Path('/home/zhuhe/code/litchi-tmp-recovery-20260913/roman-final')


def now():
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def write(path, value):
    path.write_text(json.dumps(value, indent=2) + '\n')


def references(roots):
    found = []
    for proc in Path('/proc').iterdir():
        if not proc.name.isdigit():
            continue
        links = [proc / 'cwd', proc / 'exe']
        try:
            links.extend((proc / 'fd').iterdir())
        except OSError:
            pass
        for link in links:
            try:
                target = os.readlink(link).removesuffix(' (deleted)')
            except OSError:
                continue
            if any(target == str(p) or target.startswith(str(p) + '/') for p in roots):
                found.append({'pid': proc.name, 'link': str(link), 'target': target})
        try:
            lines = (proc / 'maps').read_text().splitlines()
        except OSError:
            continue
        for line in lines:
            if any(str(p) in line for p in roots):
                found.append({'pid': proc.name, 'mapping': line})
    return found


def main():
    binaries = {}
    for variant in ['baseline', 'candidate']:
        folder = HERE.parent / 'performance' / variant
        provenance = json.loads((folder / 'binary-provenance.json').read_text())
        binary = folder / 'ods-formula-roman-evaluation-profile'
        assert sha(binary) == provenance['sha256']
        assert binary.stat().st_size == provenance['size_bytes']
        binaries[variant] = {'path': str(binary.relative_to(ROOT)),
                             'sha256': sha(binary), 'bytes': binary.stat().st_size}
    roots = OWNED + [ROOT / item['path'] for item in binaries.values()]
    assert not references(roots), 'owned scratch or binary is still in use'
    assert all(p.is_dir() and not p.is_symlink() for p in OWNED)
    seen = set()
    reclaimed = 0
    for root in OWNED:
        for path in [root, *root.rglob('*')]:
            stat = path.lstat()
            inode = (stat.st_dev, stat.st_ino)
            if inode not in seen:
                reclaimed += stat.st_blocks * 512
                seen.add(inode)
    RECOVERY.mkdir(parents=True, exist_ok=False)
    archive = RECOVERY / 'unique-scratch.tar.gz'
    unique = []
    with tarfile.open(archive, 'w:gz') as out:
        for path in sorted(OWNED[0].rglob('*')):
            if path.name == '.git' or not path.is_file():
                continue
            assert not path.is_symlink(), path
            relative = path.relative_to(OWNED[0])
            digest = sha(path)
            counterpart = ROOT / relative
            if counterpart.is_file() and sha(counterpart) == digest:
                continue
            unique.append({'path': str(relative), 'sha256': digest,
                           'bytes': path.stat().st_size})
            out.add(path, arcname=str(relative), recursive=False)
    with tarfile.open(archive, 'r:gz') as saved:
        assert len(saved.getmembers()) == len(unique)
        for item in unique:
            with saved.extractfile(item['path']) as stream:
                assert hashlib.file_digest(stream, 'sha256').hexdigest() == item['sha256']
    recovery = {'archive_sha256': sha(archive), 'files': unique}
    write(RECOVERY / 'recovery-manifest.json', recovery)
    write(HERE / 'recovery-manifest.json', recovery)
    assert not references(roots), 'process started during recovery capture'
    subprocess.run(['git', 'worktree', 'remove', '--force', str(OWNED[0])],
                   cwd=ROOT, check=True)
    for path in OWNED[1:]:
        shutil.rmtree(path)
    removed = []
    for variant, item in binaries.items():
        binary = ROOT / item['path']
        assert sha(binary) == item['sha256']
        binary.unlink()
        removed.append({'variant': variant, **item})
    remaining = references(roots)
    assert not remaining and all(not p.exists() for p in roots)
    write(HERE / 'root-binary-verification.json', {
        'verified_at': now(), 'binaries': binaries, 'removed_binaries': removed,
        'removed_after_verification_at': now()})
    receipt = {'removed_directories': [str(p) for p in OWNED],
               'all_removed_paths_absent': True, 'remaining_process_references': remaining,
               'reclaimed_allocated_bytes': reclaimed, 'unique_archive': str(archive),
               'unique_archive_sha256': recovery['archive_sha256'],
               'unique_file_count': len(unique), 'finished_at': now()}
    write(HERE / 'cleanup.json', receipt)
    print(json.dumps(receipt, indent=2))


if __name__ == '__main__':
    main()
