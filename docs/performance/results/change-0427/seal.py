#!/usr/bin/env python3
"""Losslessly compress logs and refresh the complete bundle inventory."""
import gzip
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def main():
    compression = ROOT / 'compression.json'
    records = json.loads(compression.read_text()) if compression.exists() else {}
    for path in sorted(ROOT.rglob('*.log')):
        stored_path = path.with_suffix(path.suffix + '.gz')
        assert not stored_path.exists()
        raw = path.read_bytes()
        stored = gzip.compress(raw, compresslevel=9, mtime=0)
        assert gzip.decompress(stored) == raw
        stored_path.write_bytes(stored)
        records[str(stored_path.relative_to(ROOT))] = {
            'original_path': str(path.relative_to(ROOT)),
            'original_sha256': sha(raw), 'original_bytes': len(raw),
            'stored_sha256': sha(stored), 'stored_bytes': len(stored)}
        path.unlink()
    compression.write_text(json.dumps(records, sort_keys=True, indent=2) + '\n')
    checks = {}
    for path in sorted((ROOT / 'checks').glob('*.json')):
        row = json.loads(path.read_text())
        if 'source_before' in row:
            assert row['status'] in ['pass', 'failed'], path
            checks[path.stem] = row['status']
    (ROOT / 'expected-checks.json').write_text(json.dumps(checks, sort_keys=True, indent=2) + '\n')
    files = sorted(path for path in ROOT.rglob('*') if path.is_file() and path.name != 'SHA256SUMS')
    (ROOT / 'SHA256SUMS').write_text(''.join(
        sha(path.read_bytes()) + '  ' + str(path.relative_to(ROOT)) + '\n' for path in files))
    print(json.dumps({'inventory_files': len(files), 'command_receipts': len(checks),
                      'compressed_logs': len(records)}))


if __name__ == '__main__':
    main()
