#!/usr/bin/env python3
"""Independently inspect Formula framing in the checked-in large CFB stream."""
import hashlib
import json
from pathlib import Path
import struct

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
FIXTURE = 'test-data/ole/xls/FormulaEvalTestData.xls'


def main():
    raw = (REPO / FIXTURE).read_bytes()
    assert raw[:8] == bytes.fromhex('d0cf11e0a1b11ae1')
    sector_size = 1 << struct.unpack_from('<H', raw, 30)[0]
    assert sector_size == 512 and struct.unpack_from('<I', raw, 72)[0] == 0
    sectors = len(raw) // sector_size - 1

    def sector(index):
        assert 0 <= index < sectors
        return raw[(index + 1) * sector_size:(index + 2) * sector_size]

    fat = []
    for index in struct.unpack_from('<109I', raw, 76):
        if index != 0xffffffff:
            fat.extend(struct.unpack('<128I', sector(index)))

    def chain(index):
        seen, chunks = set(), []
        while index != 0xfffffffe:
            assert index not in seen and len(seen) < sectors
            seen.add(index)
            chunks.append(sector(index))
            index = fat[index]
        return b''.join(chunks)

    directory = chain(struct.unpack_from('<I', raw, 48)[0])
    workbook = None
    for offset in range(0, len(directory), 128):
        entry = directory[offset:offset + 128]
        length = struct.unpack_from('<H', entry, 64)[0]
        if length < 2:
            continue
        name = entry[:length - 2].decode('utf-16le')
        if name in ('Workbook', 'Book'):
            assert workbook is None
            size = struct.unpack_from('<Q', entry, 120)[0]
            assert size >= 4096  # This bounded diagnostic does not implement MiniFAT.
            workbook = chain(struct.unpack_from('<I', entry, 116)[0])[:size]
            assert len(workbook) == size
    assert workbook is not None
    offset, formulas, extra_records = 0, 0, []
    while offset + 4 <= len(workbook):
        kind, length = struct.unpack_from('<HH', workbook, offset)
        payload = workbook[offset + 4:offset + 4 + length]
        assert len(payload) == length
        if kind == 6:
            formulas += 1
            assert length >= 22
            cce = struct.unpack_from('<H', payload, 20)[0]
            assert 22 + cce <= length
            if 22 + cce < length:
                extra_records.append({
                    'stream_offset': offset, 'row': struct.unpack_from('<H', payload)[0],
                    'column': struct.unpack_from('<H', payload, 2)[0],
                    'payload_bytes': length, 'token_bytes': cce,
                    'tokens_hex': payload[22:22 + cce].hex(),
                    'extra_hex': payload[22 + cce:].hex(),
                })
        offset += 4 + length
    print(json.dumps({
        'fixture': FIXTURE, 'fixture_sha256': hashlib.sha256(raw).hexdigest(),
        'stream_sha256': hashlib.sha256(workbook).hexdigest(),
        'formula_count': formulas, 'records_with_extra': extra_records,
    }, indent=2))


if __name__ == '__main__':
    main()
