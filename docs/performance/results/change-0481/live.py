#!/usr/bin/env python3
"""Check final source, templates and live binaries before runtime cleanup."""
import subprocess
from common import ROOT, REPO, TEMP, read, write, meta, sha, now
import gate


def main():
    binaries = {f'{arm}-{mode}': b for arm in ('control', 'candidate') for mode,b in read(ROOT / f'{arm}-binaries.json').items()}
    source = gate.snapshot()
    assert all(source['sha256'] == b['source_manifest_sha256'] for mode,b in binaries.items() if mode.startswith('candidate-'))
    for mode, binary in binaries.items():
        path = TEMP / mode / 'docx_plain_paragraph_tail_append'
        assert str(path) == binary['path']
        assert meta(path) == {key: binary[key] for key in ('bytes', 'sha256')}
    inputs = read(ROOT / 'embedded-inputs.json')['files']
    for name, expected in inputs.items():
        assert meta(REPO / name) == meta(ROOT / expected['path']) == {
            key: expected[key] for key in ('bytes', 'sha256')}
    protected = read(ROOT / 'initial-state.json')['protected_files']
    assert all(sha(REPO / name) == digest for name, digest in protected.items())
    write(ROOT / 'live-state.json', dict(
        schema='docx-borrowed-names-live-state-v1',
        checked_utc=now(), source=source, binaries=binaries,
        embedded_input_count=len(inputs), protected_files=protected,
        revision=subprocess.check_output(['git', 'rev-parse', 'HEAD'],
                                         cwd=REPO, text=True).strip(), status='pass'))
    print('Final source, eight templates, all four binaries and protected files verified.')


if __name__ == '__main__':
    main()
