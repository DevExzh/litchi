#!/usr/bin/env python3
"""Bind every validation attempt and the exact required final gate set."""
from common import ROOT, read, write, sha
import verify


def main():
    binaries = {arm:read(ROOT / f'{arm}-binaries.json') for arm in ('control','candidate')}
    source = {arm:b['normal']['source_manifest_sha256'] for arm,b in binaries.items()}
    assert all(source[arm] == b['allocator']['source_manifest_sha256'] for arm,b in binaries.items())
    attempts = {}
    for path in sorted((ROOT / 'validation').glob('*.json')):
        if path.name.endswith('.started.json'):
            continue
        receipt = read(path)
        attempts[path.stem] = dict(path=path.relative_to(ROOT).as_posix(),
                                   sha256=sha(path), argv=receipt['argv'],
                                   exit_code=receipt['exit_code'],
                                   source_unchanged=receipt['source_unchanged'])
    for label in verify.REQUIRED_LABELS:
        assert attempts[label]['exit_code'] == 0
        receipt = read(ROOT / attempts[label]['path'])
        arm = 'control' if label.startswith(('build-control-', 'pilot-control-')) else 'candidate'
        assert receipt['source_before']['sha256'] == receipt['source_after']['sha256'] == source[arm]
    write(ROOT / 'rust-validation.json', dict(
        schema='docx-borrowed-names-validation-v1',
        source_sha256=source, required=list(verify.REQUIRED_LABELS),
        attempts=attempts))
    print('All required final gates and every development attempt bound.')


if __name__ == '__main__':
    main()
