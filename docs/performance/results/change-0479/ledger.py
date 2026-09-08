#!/usr/bin/env python3
"""Bind every validation attempt and the exact required final gate set."""
from common import ROOT, read, write, sha
import verify


def main():
    binaries = read(ROOT / 'binaries.json')
    source = binaries['normal']['source_manifest_sha256']
    assert source == binaries['allocator']['source_manifest_sha256']
    attempts = {}
    for path in sorted((ROOT / 'validation').glob('*.json')):
        if path.name.endswith('.started.json'):
            continue
        receipt = read(path)
        attempts[path.stem] = dict(path=path.relative_to(ROOT).as_posix(),
                                   sha256=sha(path), argv=receipt['argv'],
                                   exit_code=receipt['exit_code'],
                                   source_unchanged=receipt['source_unchanged'])
    for label in verify.REQUIRED_FINAL_LABELS:
        assert attempts[label]['exit_code'] == 0
        receipt = read(ROOT / attempts[label]['path'])
        assert receipt['source_before']['sha256'] == receipt['source_after']['sha256'] == source
    write(ROOT / 'rust-validation.json', dict(
        schema='docx-plain-paragraph-tail-append-rust-validation-v1',
        final_source_sha256=source, required=list(verify.REQUIRED_FINAL_LABELS),
        attempts=attempts))
    print('All required final gates and every development attempt bound.')


if __name__ == '__main__':
    main()
