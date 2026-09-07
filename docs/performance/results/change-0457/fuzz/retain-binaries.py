#!/usr/bin/env python3
"""Losslessly retain captured fuzz executables without committing raw binaries."""
import gzip
import hashlib
import json
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parent
POST = ROOT / 'artifacts/post-run'


def identity(path):
    with path.open('rb') as stream:
        return {'bytes': path.stat().st_size,
                'sha256': hashlib.file_digest(stream, 'sha256').hexdigest()}


def main():
    receipt = ROOT / 'binary-retention.json'
    assert not receipt.exists()
    inventory_path = POST / 'inventory.json'
    inventory = json.loads(inventory_path.read_text())
    rows = {row['path']: row for row in inventory['files']}
    records = []
    for name in ('zip/parse_zip', 'xml/scan_xml'):
        source = POST / name
        target = source.with_name(source.name + '.gz')
        assert source.is_file() and not source.is_symlink()
        assert not target.exists()
        expected = rows[name]
        assert identity(source) == {k: expected[k] for k in ('bytes', 'sha256')}
        with source.open('rb') as incoming, target.open('xb') as outgoing:
            with gzip.GzipFile(filename='', mode='wb', fileobj=outgoing,
                               compresslevel=9, mtime=0) as compressed:
                shutil.copyfileobj(incoming, compressed, 1 << 20)
        digest = hashlib.sha256()
        size = 0
        with gzip.open(target, 'rb') as decoded:
            while chunk := decoded.read(1 << 20):
                size += len(chunk)
                digest.update(chunk)
        assert size == expected['bytes'] and digest.hexdigest() == expected['sha256']
        records.append({'original': expected,
                        'retained': {'path': target.relative_to(POST).as_posix(),
                                     **identity(target)}, 'encoding': 'gzip'})
    value = {'schema': 'litchi-0457-fuzz-binary-retention-v1',
             'original_inventory': {'path': 'artifacts/post-run/inventory.json',
                                    **identity(inventory_path)},
             'driver': {'path': Path(__file__).name, **identity(Path(__file__))},
             'records': records,
             'scope': 'lossless storage transformation; original capture inventory unchanged; shared target cache preserved'}
    with receipt.open('x') as stream:
        json.dump(value, stream, indent=2)
        stream.write('\n')
    for record in records:
        (POST / record['original']['path']).unlink()
    print(json.dumps({'status': 'pass', 'executables': len(records),
                      'original_bytes': sum(r['original']['bytes'] for r in records),
                      'retained_bytes': sum(r['retained']['bytes'] for r in records)}))


if __name__ == '__main__':
    main()
