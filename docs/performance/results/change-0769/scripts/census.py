#!/usr/bin/env python3
"""Independent census of CFB root mini-stream sizes and mini-stream tails.

Parses each compound file with its own small reader (no litchi code), and
classifies how its mini streams relate to the root entry's stream size (the
mini stream's declared size, MS-CFB 2.6.1):

- aligned:        the root size is a multiple of 64;
- unaligned-free: the root size is not a multiple of 64 and no mini stream
                  uses the partial last mini sector;
- tail-inside:    a mini stream's last sector is the partial last mini sector
                  and every byte it needs lies within the root size;
- tail-outside:   a mini stream needs a byte of the partial last mini sector
                  beyond the root size (genuinely out of bounds);
- beyond:         a mini stream's chain names a mini sector at or past
                  ceil(root size / 64);
- unparsed:       the independent reader could not walk the file.

Usage: census.py PATHS_FILE > census.json
"""

import json
import struct
import sys

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
MAXREGSECT = 0xFFFFFFFA


class Unparsed(Exception):
    pass


def chain(table, start, limit):
    out = []
    seen = set()
    sector = start
    while sector != ENDOFCHAIN:
        if sector >= MAXREGSECT or sector >= len(table) or sector in seen or len(out) > limit:
            raise Unparsed(f"bad chain at {sector:#x}")
        seen.add(sector)
        out.append(sector)
        sector = table[sector]
    return out


def parse(data):
    if len(data) < 512 or data[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
        raise Unparsed("not CFB")
    major = struct.unpack_from("<H", data, 0x1A)[0]
    shift = struct.unpack_from("<H", data, 0x1E)[0]
    if shift not in (9, 12):
        raise Unparsed("sector shift")
    ssz = 1 << shift
    nfat = struct.unpack_from("<I", data, 0x2C)[0]
    dir_start = struct.unpack_from("<I", data, 0x30)[0]
    minifat_start = struct.unpack_from("<I", data, 0x3C)[0]
    nminifat = struct.unpack_from("<I", data, 0x40)[0]
    difat_start = struct.unpack_from("<I", data, 0x44)[0]
    ndifat = struct.unpack_from("<I", data, 0x48)[0]
    nsect = (len(data) + ssz - 1) // ssz - 1

    def sector(index):
        off = (index + 1) * ssz
        if index >= MAXREGSECT or off >= len(data):
            raise Unparsed(f"sector {index} outside file")
        return data[off:off + ssz].ljust(ssz, b"\0")

    fat_locs = list(struct.unpack_from("<109I", data, 0x4C))[: min(nfat, 109)]
    difat = difat_start
    guard = 0
    while len(fat_locs) < nfat and difat < MAXREGSECT and guard <= nsect:
        words = struct.unpack(f"<{ssz // 4}I", sector(difat))
        fat_locs.extend(words[:-1][: nfat - len(fat_locs)])
        difat = words[-1]
        guard += 1
    if len(fat_locs) < nfat:
        raise Unparsed("short DIFAT")
    fat = []
    for loc in fat_locs:
        fat.extend(struct.unpack(f"<{ssz // 4}I", sector(loc)))
    dir_chain = chain(fat, dir_start, len(fat))
    directory = b"".join(sector(s) for s in dir_chain)
    minifat = []
    if minifat_start != ENDOFCHAIN:
        for s in chain(fat, minifat_start, len(fat)):
            minifat.extend(struct.unpack(f"<{ssz // 4}I", sector(s)))
    entries = []
    for offset in range(0, len(directory), 128):
        raw = directory[offset:offset + 128]
        name_len = struct.unpack_from("<H", raw, 0x40)[0]
        kind = raw[0x42]
        start = struct.unpack_from("<I", raw, 0x74)[0]
        size = struct.unpack_from("<Q", raw, 0x78)[0]
        if major == 3:
            size &= 0xFFFFFFFF
        name = raw[: max(0, min(name_len, 64) - 2)].decode("utf-16-le", "replace")
        entries.append((offset // 128, name, kind, start, size))
    if not entries or entries[0][2] != 5:
        raise Unparsed("no root entry")
    return {
        "major": major,
        "sector_size": ssz,
        "fat": fat,
        "minifat": minifat,
        "entries": entries,
        "minifat_sectors": nminifat,
        "ndifat": ndifat,
    }


def classify(data):
    cfb = parse(data)
    root = cfb["entries"][0]
    root_size = root[4]
    tail = root_size % 64
    capacity = (root_size + 63) // 64
    partial = root_size // 64 if tail else None
    result = {
        "major": cfb["major"],
        "root_size": root_size,
        "root_mod_64": tail,
        "root_sectors": (root_size + cfb["sector_size"] - 1) // cfb["sector_size"],
        "mini_streams": 0,
        "classes": [],
        "streams_in_partial": [],
    }
    classes = set()
    for sid, name, kind, start, size in cfb["entries"][1:]:
        if kind != 2 or size == 0 or size >= 4096:
            continue
        result["mini_streams"] += 1
        try:
            sectors = chain(cfb["minifat"], start, len(cfb["minifat"]))
        except Unparsed:
            classes.add("bad-mini-chain")
            continue
        count = (size + 63) // 64
        for position, sector in enumerate(sectors[:count]):
            need = min(64, size - 64 * position)
            if sector >= capacity:
                classes.add("beyond")
            elif sector == partial:
                end = sector * 64 + need
                shape = "tail-inside" if end <= root_size else "tail-outside"
                classes.add(shape)
                result["streams_in_partial"].append(
                    {
                        "sid": sid,
                        "name": name,
                        "size": size,
                        "position": position,
                        "last": position == count - 1,
                        "need": need,
                        "end": end,
                        "shape": shape,
                    }
                )
    if not classes:
        classes.add("aligned" if tail == 0 else "unaligned-free")
    result["classes"] = sorted(classes)
    return result


def main():
    rows = []
    for path in open(sys.argv[1]).read().splitlines():
        if not path:
            continue
        data = open(path, "rb").read()
        try:
            row = classify(data)
        except (Unparsed, struct.error, IndexError, UnicodeDecodeError) as error:
            row = {"classes": ["unparsed"], "error": str(error)}
        row["path"] = path
        row["bytes"] = len(data)
        rows.append(row)
    json.dump(rows, sys.stdout, indent=1)


if __name__ == "__main__":
    main()
