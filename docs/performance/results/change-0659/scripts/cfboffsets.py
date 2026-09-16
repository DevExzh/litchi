"""Derives a byte inside the first FAT sector and a byte inside the first
directory sector of a CFB artifact, so a change-under-read sweep can place its
witness in the allocation table or in the directory rather than in the payload."""
import struct, sys

def offsets(path):
    with open(path, 'rb') as handle:
        header = handle.read(512)
    if header[:8] != bytes.fromhex('d0cf11e0a1b11ae1'):
        return None
    shift = struct.unpack_from('<H', header, 30)[0]
    sector = 1 << shift
    first_fat = struct.unpack_from('<I', header, 76)[0]
    first_dir = struct.unpack_from('<I', header, 48)[0]
    if first_fat >= 0xFFFFFFFA or first_dir >= 0xFFFFFFFA:
        return None
    return ((first_fat + 1) * sector + 8, (first_dir + 1) * sector + 96, sector)

for path in sys.argv[1:]:
    got = offsets(path)
    print(path, *(got if got else ('-', '-', '-')), sep='\t')
