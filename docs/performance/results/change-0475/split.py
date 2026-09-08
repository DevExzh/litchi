#!/usr/bin/env python3
"""Retain original Heaptrack bytes in bounded Git blobs after all exports."""
import datetime
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PART_BYTES = 64 * 1024 * 1024


def main():
    assert (ROOT / 'compression.json').is_file()
    assert (ROOT / 'supplement-compression.json').is_file()
    assert not (ROOT / 'chunking.json').exists()
    records = []
    originals = []
    for lane in ('heap-H1', 'heap-H2'):
        path = ROOT / lane / 'heaptrack.zst'
        digest = hashlib.sha256()
        count = 0
        parts = []
        with path.open('rb') as stream:
            while data := stream.read(PART_BYTES):
                digest.update(data)
                count += len(data)
                part = path.with_name(path.name + f'.part-{len(parts):03d}')
                with part.open('xb') as output:
                    output.write(data)
                parts.append(dict(path=part.relative_to(ROOT).as_posix(), bytes=len(data),
                                  sha256=hashlib.sha256(data).hexdigest()))
        expected = json.loads((path.parent / 'receipt.json').read_text())['artifacts'][path.name]
        assert dict(bytes=count, sha256=digest.hexdigest()) == expected
        replay = hashlib.sha256()
        for part in parts:
            data = (ROOT / part['path']).read_bytes()
            assert len(data) == part['bytes'] and hashlib.sha256(data).hexdigest() == part['sha256']
            replay.update(data)
        assert replay.hexdigest() == expected['sha256']
        records.append(dict(path=path.relative_to(ROOT).as_posix(), **expected, parts=parts))
        originals.append(path)
    with (ROOT / 'chunking.json').open('x') as stream:
        json.dump(dict(schema='litchi-0475-chunking-v1', chunk_bytes=PART_BYTES,
            recorded_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
            driver_sha256=hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
            artifacts=records), stream, indent=2, sort_keys=True)
        stream.write('\n')
    for path in originals:
        path.unlink()
    print('split two original traces; concatenated hashes match capture receipts')


if __name__ == '__main__':
    main()
