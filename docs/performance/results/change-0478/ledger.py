#!/usr/bin/env python3
"""Assemble the final ledger without discarding any development receipt."""
from common import ROOT, read, write, sha
import verify


def main():
    protocol = read(ROOT / 'protocol.json')
    binaries = verify.check_binaries(ROOT, protocol)
    attempts = {}
    for path in sorted((ROOT / 'validation').glob('*.json')):
        if path.name.endswith('.started.json'):
            continue
        gate = read(path)
        record = dict(path=path.relative_to(ROOT).as_posix(), sha256=sha(path),
                      argv=gate['argv'], exit_code=gate['exit_code'],
                      source_unchanged=gate['source_unchanged'])
        attempts[path.stem] = record
    for label in verify.REQUIRED_FINAL_LABELS:
        assert attempts[label]['exit_code'] == 0
        assert attempts[label]['source_unchanged'] is True
    ledger = dict(schema='pptx-metadata-spool-rust-validation-v1',
                  final_source_sha256=binaries['normal']['source_manifest_sha256'],
                  required=list(verify.REQUIRED_FINAL_LABELS), attempts=attempts)
    write(ROOT / 'rust-validation.json', ledger)
    print(verify.check_rust_validation(ROOT, binaries, protocol))


if __name__ == '__main__':
    main()
