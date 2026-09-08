#!/usr/bin/env python3
"""Remove authenticated batch runtimes while preserving shared Cargo caches."""
from common import ROOT, REPO, TEMP, read, write, meta, sha, now


def main():
    live = read(ROOT / 'live-state.json')
    assert live['status'] == 'pass'
    assert not (ROOT / 'cleanup.json').exists()
    expected = {TEMP / mode / 'docx_plain_paragraph_tail_append'
                for mode in live['binaries']}
    expected.add(TEMP / 'cpu.lock')
    files = {p for p in TEMP.rglob('*') if p.is_file()}
    assert files == expected, files ^ expected
    assert not any(p.is_symlink() for p in TEMP.rglob('*'))
    for mode, binary in live['binaries'].items():
        assert meta(TEMP / mode / 'docx_plain_paragraph_tail_append') == {
            key: binary[key] for key in ('bytes', 'sha256')}
    assert (TEMP / 'cpu.lock').stat().st_size == 0
    caches = [REPO / 'target', REPO / 'tools/perf-baseline/target']
    assert all(p.is_dir() for p in caches)
    assert all(sha(REPO / name) == digest for name, digest in live['protected_files'].items())
    for path in expected:
        path.unlink()
    for path in sorted(TEMP.rglob('*'), key=lambda p: len(p.parts), reverse=True):
        path.rmdir()
    TEMP.rmdir()
    write(ROOT / 'cleanup.json', dict(
        schema='docx-plain-paragraph-tail-append-cleanup-v1', checked_utc=now(),
        live_state_sha256=sha(ROOT / 'live-state.json'),
        removed_binaries=live['binaries'], temporary_root_absent=not TEMP.exists(),
        shared_caches_retained=[str(p) for p in caches],
        protected_files=live['protected_files'], status='pass'))
    print('Authenticated batch runtimes removed; shared caches retained.')


if __name__ == '__main__':
    main()
