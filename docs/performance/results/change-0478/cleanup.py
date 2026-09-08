#!/usr/bin/env python3
"""Remove only authenticated batch runtime files after live verification."""
from common import ROOT, REPO, TEMP, read, write, meta, sha, now


def main():
    live = read(ROOT / 'live-state.json')
    assert live['status'] == 'pass'
    assert not (ROOT / 'cleanup.json').exists()
    binaries = read(ROOT / 'binaries.json')
    expected = {TEMP / mode / 'pptx_metadata_spool' for mode in binaries}
    expected.add(TEMP / 'cpu.lock')
    files = {p for p in TEMP.rglob('*') if p.is_file()}
    assert files == expected, files ^ expected
    assert not any(p.is_symlink() for p in TEMP.rglob('*'))
    removed = {}
    for mode, spec in binaries.items():
        path = TEMP / mode / 'pptx_metadata_spool'
        identity = {key: spec[key] for key in ('bytes', 'sha256')}
        assert meta(path) == identity
        removed[str(path)] = identity
    assert (TEMP / 'cpu.lock').stat().st_size == 0
    for name, digest in live['protected_files'].items():
        assert sha(REPO / name) == digest
    caches = [REPO / 'target', REPO / 'tools/perf-baseline/target']
    assert all(p.is_dir() for p in caches)
    for path in expected:
        path.unlink()
    for path in sorted(TEMP.rglob('*'), key=lambda p: len(p.parts), reverse=True):
        path.rmdir()
    TEMP.rmdir()
    write(ROOT / 'cleanup.json', dict(
        schema='pptx-metadata-spool-cleanup-v1', checked_utc=now(),
        live_state_sha256=sha(ROOT / 'live-state.json'),
        preliminary_cleanup_sha256=sha(ROOT / 'preliminary-runtime-cleanup.json'),
        removed_binaries=removed, temporary_root_absent=not TEMP.exists(),
        scratch_files_remaining=0, shared_caches_retained=[str(p) for p in caches],
        protected_files=live['protected_files'], status='pass'))
    print('Authenticated final binaries and empty batch directories removed; shared caches retained.')


if __name__ == '__main__':
    main()
