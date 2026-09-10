#!/usr/bin/env python3
"""Remove only this batch's disposable clones/build tree after custody checks."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time
from build import ROOT, TEMP, source_manifest, write
import capture


def active_references():
    references = []
    for process in Path('/proc').iterdir():
        if not process.name.isdigit() or int(process.name) == os.getpid():
            continue
        try:
            for name in ('cmdline', 'environ', 'maps'):
                references.append((process / name).read_bytes().decode(errors='replace'))
            for name in ('cwd', 'exe'):
                try:
                    references.append(os.readlink(process / name))
                except OSError:
                    pass
            for fd in (process / 'fd').iterdir():
                try:
                    references.append(os.readlink(fd))
                except OSError:
                    pass
        except OSError:
            pass
    return '\n'.join(references)


def main():
    assert not (ROOT / 'cleanup.json').exists()
    protocol, _, builds = capture._load_protocol()
    entries = capture._collect('formal1', protocol, builds)
    assert len(entries) == 32
    for phase in ('before', 'after'):
        source = TEMP / phase
        assert source.is_dir() and not source.is_symlink()
        assert (source / '.git').is_dir()  # standalone clone, not shared worktree metadata
        manifest = json.loads((ROOT / 'builds' / (phase + '-final1-source.json')).read_text())
        assert source_manifest(source) == manifest
        subprocess.run(['git', 'apply', '--reverse', '--check', str(ROOT / (phase + '-harness.patch'))],
                       cwd=source, check=True)
    protected = json.loads((ROOT / 'protected-primary.json').read_text())
    for name, digest in protected.items():
        assert hashlib.sha256((capture.REPO / name).read_bytes()).hexdigest() == digest
    removed = []
    free_before = shutil.disk_usage(TEMP).free
    for name in ('before', 'after', 'target', 'tmp', 'projections', 'runs'):
        path = TEMP / name
        if not path.exists():
            continue
        assert path.is_dir() and not path.is_symlink()
        assert str(path) not in active_references(), f'active process references {path}'
        allocated = 0
        files = 0
        for directory, dirs, names in os.walk(path, followlinks=False):
            for item in names:
                allocated += (Path(directory) / item).lstat().st_blocks * 512
                files += 1
        shutil.rmtree(path)
        removed.append({'path': str(path), 'allocated_bytes': allocated, 'files': files})
    pycache = ROOT / '__pycache__'
    if pycache.exists():
        shutil.rmtree(pycache)
    capture.load_builds()  # retained executable and gate custody survives removal
    write(ROOT / 'cleanup.json', {'schema': 'docx-phase-cleanup-v1', 'status': 'pass',
          'finished_ns': time.time_ns(), 'removed': removed, 'source_manifests_verified': True,
          'patches_reverse_checked': True, 'protected_files_unchanged': True,
          'retained_builds': list(builds), 'free_before': free_before,
          'free_after': shutil.disk_usage(TEMP).free,
          'scope': 'Owned standalone clones, Cargo target and private scratch only; retained executables preserved.'})
    print(json.dumps({'removed_GiB': sum(x['allocated_bytes'] for x in removed) / 2**30,
                      'retained_executables': len(builds)}))


if __name__ == '__main__':
    main()
