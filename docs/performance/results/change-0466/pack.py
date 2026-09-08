#!/usr/bin/env python3
"""Losslessly retain large profile artifacts; record and verify original bytes."""
import gzip
import hashlib
import json
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parent


def identity(path, opener=open):
    digest, size = hashlib.sha256(), 0
    with opener(path, 'rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
            size += len(block)
    return dict(sha256=digest.hexdigest(), bytes=size)


def main():
    names = ('samples/perf.data', 'samples/perf-script.txt',
             'samples-fp/perf.data', 'samples-fp/perf-script.stdout', 'heaptrack/print.txt')
    rows = {}
    for name in names:
        source = ROOT / name
        original = identity(source)
        target = source.with_name(source.name + '.gz')
        with source.open('rb') as input_stream, target.open('xb') as output_stream:
            with gzip.GzipFile(filename='', mode='wb', fileobj=output_stream, mtime=0) as zipper:
                shutil.copyfileobj(input_stream, zipper)
        assert identity(target, gzip.open) == original
        rows[name] = dict(original=original, compressed=dict(path=str(target.relative_to(ROOT)), **identity(target)))
        source.unlink()
    (ROOT / 'compression.json').write_text(json.dumps(rows, indent=2) + '\n')


if __name__ == '__main__':
    main()
