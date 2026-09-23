#!/usr/bin/env python3
"""Measure the Deflate framing that zlib adds to incompressible bytes.

Change 0742 bounds a transferred member's compressed size by its decoded size
plus one 5-byte stored-block header per started 4 KiB plus 64 bytes. This
script prints, for 2 MiB of seeded pseudo-random bytes, the framing (compressed
minus decoded bytes) of raw Deflate at levels 1, 6 and 9 for several zlib
memory levels, next to what the prescribed 65,535-byte unit and the chosen
4 KiB unit allow. It uses Python's zlib module (the system zlib), whose block
structure zlib-rs reproduces.

Usage: zlib_framing.py
"""

from __future__ import annotations

import random
import zlib

DECODED = 2 * 1024 * 1024
HEADER = 5
SLACK = 64


def allowed(unit: int) -> int:
    return -(-DECODED // unit) * HEADER + SLACK


def main() -> None:
    generator = random.Random(742)
    data = bytes(generator.getrandbits(8) for _ in range(DECODED))
    print(f"zlib runtime\t{zlib.ZLIB_RUNTIME_VERSION}")
    print(f"decoded bytes\t{DECODED}")
    print(f"allowed framing, 65,535-byte unit\t{allowed(65_535)}")
    print(f"allowed framing, 4 KiB unit\t{allowed(4 * 1024)}")
    print("memLevel\tlevel\tcompressed\tframing\tverdict")
    for memory in (9, 8, 4, 2, 1):
        for level in (1, 6, 9):
            compressor = zlib.compressobj(level, zlib.DEFLATED, -15, memory)
            compressed = len(compressor.compress(data) + compressor.flush())
            framing = compressed - DECODED
            verdict = "transfer" if framing <= allowed(4 * 1024) else "recompress"
            print(f"{memory}\t{level}\t{compressed}\t{framing}\t{verdict}")


if __name__ == "__main__":
    main()
