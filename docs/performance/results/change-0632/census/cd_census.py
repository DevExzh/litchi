#!/usr/bin/env python3
"""Census of the ZIP tail geometry over every ZIP container under `test-data`.

For each container reports, from the bytes alone:

  cd_size          the EOCD-declared central-directory size
  cd_off           the EOCD-declared central-directory offset
  comment_len      the EOCD comment length
  zip64            whether the EOCD carries a ZIP64 sentinel
  fast_path        whether `ZipLocator::locate_in_reader`'s 22-byte probe at
                   `len - 22` parses with `comment_len == 0 && !is_zip64`,
                   which is the condition for today's one-read EOCD locate
  prefixed         whether the first central record is NOT at the declared
                   `cd_off` but is at `eocd_offset - cd_size` -- the
                   `finish_locate_in_reader` base-offset fallback
  tail_pad         bytes between the end of the central directory and the EOCD

Usage: cd_census.py <container-list> [<repo-root>]
"""
import struct, sys, os

EOCD_SIG = b'PK\x05\x06'
CD_SIG = b'PK\x01\x02'
Z64L_SIG = b'PK\x06\x07'
Z64_SIG = b'PK\x06\x06'
EOCD_FIXED = 22
CD_FIXED = 46


def parse_eocd_fixed(b, off):
    """Parse the 22-byte fixed EOCD record at `off`, or None."""
    if off < 0 or off + EOCD_FIXED > len(b):
        return None
    if b[off:off + 4] != EOCD_SIG:
        return None
    (disk, cd_disk, disk_entries, entries, cd_size, cd_off, comment_len) = struct.unpack_from(
        '<HHHHIIH', b, off + 4)
    return dict(disk=disk, cd_disk=cd_disk, disk_entries=disk_entries, entries=entries,
                cd_size=cd_size, cd_off=cd_off, comment_len=comment_len)


def is_zip64(e):
    return (e['disk'] == 0xFFFF or e['cd_disk'] == 0xFFFF or e['disk_entries'] == 0xFFFF
            or e['entries'] == 0xFFFF or e['cd_size'] == 0xFFFFFFFF
            or e['cd_off'] == 0xFFFFFFFF)


def find_eocd(b):
    """Backwards search for the EOCD signature, as the locator does."""
    for i in range(len(b) - EOCD_FIXED, -1, -1):
        if b[i:i + 4] == EOCD_SIG:
            return i
    return None


def main():
    listing = sys.argv[1]
    root = sys.argv[2] if len(sys.argv) > 2 else '.'
    rows = []
    with open(listing) as fh:
        paths = [line.strip() for line in fh if line.strip()]
    print('\t'.join(['path', 'file_bytes', 'eocd_off', 'cd_off', 'cd_size', 'entries',
                     'comment_len', 'zip64', 'fast_path', 'prefixed', 'tail_pad']))
    for rel in paths:
        path = os.path.join(root, rel)
        with open(path, 'rb') as fh:
            b = fh.read()
        eocd_off = find_eocd(b)
        if eocd_off is None:
            print('\t'.join([rel, str(len(b))] + ['-'] * 9))
            continue
        e = parse_eocd_fixed(b, eocd_off)
        z64 = is_zip64(e)
        probe = parse_eocd_fixed(b, len(b) - EOCD_FIXED)
        fast = bool(probe is not None and probe['comment_len'] == 0 and not is_zip64(probe))
        cd_size, cd_off = e['cd_size'], e['cd_off']
        # ZIP64: take the real directory geometry from the ZIP64 record.
        if z64:
            loc = eocd_off - 20
            if loc >= 0 and b[loc:loc + 4] == Z64L_SIG:
                z64_off, = struct.unpack_from('<Q', b, loc + 8)
                if 0 <= z64_off < len(b) and b[z64_off:z64_off + 4] == Z64_SIG:
                    cd_size, cd_off = struct.unpack_from('<QQ', b, z64_off + 40)
        at_declared = b[cd_off:cd_off + 4] == CD_SIG if cd_off + 4 <= len(b) else False
        fallback = eocd_off - cd_size
        at_fallback = (0 <= fallback and b[fallback:fallback + 4] == CD_SIG)
        prefixed = (not at_declared) and at_fallback
        real_cd_off = cd_off if at_declared else (fallback if at_fallback else cd_off)
        tail_pad = eocd_off - (real_cd_off + cd_size)
        print('\t'.join([rel, str(len(b)), str(eocd_off), str(cd_off), str(cd_size),
                         str(e['entries']), str(e['comment_len']), str(int(z64)),
                         str(int(fast)), str(int(prefixed)), str(tail_pad)]))


if __name__ == '__main__':
    main()
