#!/usr/bin/env python3
"""Change 0769: derive gen/zvi-normalized.cfb from POI's BlockSize512.zvi.

The Zeiss AxioVision writer puts FREESECT (0xFFFFFFFF) in its storages'
start sector fields, which litchi's directory validation refuses before any
stream allocation is examined. This rewrites exactly those fields to
ENDOFCHAIN and nothing else, so the file reaches the mini-stream checks.
Uses census.py's independent reader to find the directory entries.

Usage: normalize_zvi.py SOURCE.zvi OUTPUT.cfb
"""

import struct
import sys

import census


def main():
    data = bytearray(open(sys.argv[1], 'rb').read())
    parsed = census.parse(bytes(data))
    sector_size = parsed['sector_size']
    chain = census.chain(parsed['fat'], struct.unpack_from('<I', data, 0x30)[0], len(parsed['fat']))
    changed = 0
    for sid, _name, kind, start, _size in parsed['entries']:
        if kind == 1 and start == 0xFFFFFFFF:
            offset = (chain[sid * 128 // sector_size] + 1) * sector_size + (sid * 128) % sector_size
            struct.pack_into('<I', data, offset + 0x74, 0xFFFFFFFE)
            changed += 1
    open(sys.argv[2], 'wb').write(bytes(data))
    print(f'{changed} storage start fields normalized')


if __name__ == '__main__':
    main()
