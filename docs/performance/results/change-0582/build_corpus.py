#!/usr/bin/env python3
"""Build the deterministic differential corpus for change 0582.

Usage:
    build_corpus.py <repo-root> <out-dir>

Writes four families under `<out-dir>`:

  seeds/      the ZIP-format seeds change 0457 retained
  testdata/   every real ZIP container under `<repo-root>/test-data`
  mutations/  deterministic mutations of a fixed base set
  crafted/    hand-built archives that target the strict-layout proof

Everything is a pure function of the repository contents plus the constant
`RNG_SEED`, so the corpus regenerates byte for byte.  Nothing here is random in
the wall-clock sense: `random.Random(RNG_SEED)` is seeded once, drawn from in a
fixed order, and every base input is processed in sorted path order.
"""

import hashlib
import json
import os
import random
import struct
import sys
import zlib

# The whole corpus is a function of this constant.
RNG_SEED = 0x05820580

LFH = b"PK\x03\x04"
CDH = b"PK\x01\x02"
EOCD = b"PK\x05\x06"
DD = b"PK\x07\x08"

MAX_BYTES = 1 << 20  # the fuzz target ignores anything larger

# How many real archives seed the mutation family.  Bounded so the corpus
# stays regenerable in minutes rather than hours.
MUTATION_TESTDATA_BASES = 40


# ---------------------------------------------------------------------------
# A minimal, fully explicit ZIP builder.  Every field is settable so a crafted
# archive can declare a layout no ordinary writer would emit.
# ---------------------------------------------------------------------------


class Member:
    def __init__(
        self,
        name,
        payload=b"",
        method=0,
        local_extra=b"",
        central_extra=b"",
        local_offset=None,
        local_name=None,
        local_name_len=None,
        local_extra_len=None,
        local_csize=None,
        local_usize=None,
        local_crc=None,
        central_csize=None,
        central_usize=None,
        central_crc=None,
        flags=0,
        local_flags=None,
        local_method=None,
        descriptor=None,  # None | "unsigned" | "signed" | "zip64"
        gap_before=0,
    ):
        self.name = name.encode() if isinstance(name, str) else name
        self.payload = payload
        self.method = method
        self.local_extra = local_extra
        self.central_extra = central_extra
        self.local_offset = local_offset
        self.local_name = self.name if local_name is None else local_name
        self.local_name_len = local_name_len
        self.local_extra_len = local_extra_len
        self.local_csize = local_csize
        self.local_usize = local_usize
        self.local_crc = local_crc
        self.central_csize = central_csize
        self.central_usize = central_usize
        self.central_crc = central_crc
        self.flags = flags
        self.local_flags = local_flags
        self.local_method = local_method
        self.descriptor = descriptor
        self.gap_before = gap_before


def _stored_or_deflated(member):
    raw = member.payload
    if member.method == 8:
        compressor = zlib.compressobj(6, zlib.DEFLATED, -15)
        body = compressor.compress(raw) + compressor.flush()
    else:
        body = raw
    return body, zlib.crc32(raw) & 0xFFFFFFFF


def build_zip(members, comment=b"", directory_offset=None, entry_count=None):
    """Assemble an archive from fully explicit member descriptions."""
    out = bytearray()
    placed = []
    for member in members:
        out.extend(b"\x00" * member.gap_before)
        if member.local_offset is not None:
            if member.local_offset < len(out):
                del out[member.local_offset :]
            else:
                out.extend(b"\x00" * (member.local_offset - len(out)))
        offset = len(out)
        body, crc = _stored_or_deflated(member)
        csize = len(body)
        usize = len(member.payload)

        flags = member.flags
        if member.descriptor:
            flags |= 0x08
        local_flags = flags if member.local_flags is None else member.local_flags
        local_method = member.method if member.local_method is None else member.local_method
        local_crc = (0 if member.descriptor else crc) if member.local_crc is None else member.local_crc
        local_csize = (0 if member.descriptor else csize) if member.local_csize is None else member.local_csize
        local_usize = (0 if member.descriptor else usize) if member.local_usize is None else member.local_usize
        name_len = len(member.local_name) if member.local_name_len is None else member.local_name_len
        extra_len = len(member.local_extra) if member.local_extra_len is None else member.local_extra_len

        out.extend(LFH)
        out.extend(
            struct.pack(
                "<HHHHHIIIHH",
                20,
                local_flags,
                local_method,
                0,
                0,
                local_crc,
                local_csize,
                local_usize,
                name_len,
                extra_len,
            )
        )
        out.extend(member.local_name)
        out.extend(member.local_extra)
        out.extend(body)
        if member.descriptor == "unsigned":
            out.extend(struct.pack("<III", crc, csize, usize))
        elif member.descriptor == "signed":
            out.extend(DD + struct.pack("<III", crc, csize, usize))
        elif member.descriptor == "zip64":
            out.extend(DD + struct.pack("<IQQ", crc, csize, usize))
        placed.append((member, offset, crc, csize, usize))

    directory = len(out) if directory_offset is None else directory_offset
    if directory_offset is not None:
        if directory < len(out):
            del out[directory:]
        else:
            out.extend(b"\x00" * (directory - len(out)))

    central = bytearray()
    for member, offset, crc, csize, usize in placed:
        flags = member.flags | (0x08 if member.descriptor else 0)
        central.extend(CDH)
        central.extend(
            struct.pack(
                "<HHHHHHIIIHHHHHII",
                20,
                20,
                flags,
                member.method,
                0,
                0,
                crc if member.central_crc is None else member.central_crc,
                csize if member.central_csize is None else member.central_csize,
                usize if member.central_usize is None else member.central_usize,
                len(member.name),
                len(member.central_extra),
                0,
                0,
                0,
                0,
                offset,
            )
        )
        central.extend(member.name)
        central.extend(member.central_extra)
    out.extend(central)

    count = len(placed) if entry_count is None else entry_count
    out.extend(EOCD)
    out.extend(struct.pack("<HHHHIIH", 0, 0, count, count, len(central), directory, len(comment)))
    out.extend(comment)
    return bytes(out)


# ---------------------------------------------------------------------------
# Structural parsing, for the mutation family.
# ---------------------------------------------------------------------------


def parse_layout(data):
    """Return (eocd_offset, directory_offset, [(central_off, local_off, name_len,
    extra_len, comment_len, csize)]) or None."""
    eocd = data.rfind(EOCD)
    if eocd < 0 or eocd + 22 > len(data):
        return None
    try:
        total = struct.unpack_from("<H", data, eocd + 10)[0]
        directory = struct.unpack_from("<I", data, eocd + 16)[0]
    except struct.error:
        return None
    if directory == 0xFFFFFFFF or directory > len(data):
        return None
    records = []
    cursor = directory
    for _ in range(total):
        if data[cursor : cursor + 4] != CDH:
            return None
        try:
            csize = struct.unpack_from("<I", data, cursor + 20)[0]
            name_len, extra_len, comment_len = struct.unpack_from("<HHH", data, cursor + 28)
            local = struct.unpack_from("<I", data, cursor + 42)[0]
        except struct.error:
            return None
        records.append((cursor, local, name_len, extra_len, comment_len, csize))
        cursor += 46 + name_len + extra_len + comment_len
        if cursor > len(data):
            return None
    return eocd, directory, records


def interesting_offsets(data):
    """Structurally meaningful truncation points, deterministic and sorted."""
    points = set()
    layout = parse_layout(data)
    if layout:
        eocd, directory, records = layout
        points.update({eocd, eocd + 4, eocd + 22, directory, directory + 4, directory + 46})
        for central, local, name_len, extra_len, _comment, csize in records[:32]:
            points.update(
                {
                    central,
                    central + 30,
                    central + 46,
                    local,
                    local + 4,
                    local + 30,
                    local + 30 + name_len,
                    local + 30 + name_len + extra_len,
                    local + 30 + name_len + extra_len + csize,
                }
            )
    for fraction in (0, 1, 2, 3, 4, 8, 16):
        points.add(len(data) * fraction // 16)
    points.add(len(data) - 1)
    return sorted(point for point in points if 0 <= point <= len(data))


# ---------------------------------------------------------------------------
# Corpus families
# ---------------------------------------------------------------------------

ZIP_MAGICS = (LFH, EOCD, b"PK\x06\x06", b"PK\x05\x05")


def is_zip_container(data):
    return data.startswith(LFH) and EOCD in data[-70000:]


def collect_testdata(root):
    found = []
    for directory, _dirs, names in os.walk(root):
        for name in sorted(names):
            path = os.path.join(directory, name)
            try:
                with open(path, "rb") as handle:
                    data = handle.read()
            except OSError:
                continue
            if is_zip_container(data):
                found.append((os.path.relpath(path, root), data))
    found.sort(key=lambda row: row[0])
    return found


def write(out_dir, family, name, data):
    path = os.path.join(out_dir, family, name)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(data)
    return path


def mutate(rng, base):
    """Deterministic single-field mutations of one base archive."""
    produced = []
    layout = parse_layout(base)

    # (a) truncations at structurally interesting offsets
    for point in interesting_offsets(base):
        produced.append(("trunc-%d" % point, base[:point]))

    if not layout:
        return produced
    _eocd, _directory, records = layout

    # (b) single-field corruption of local headers
    local_fields = [
        ("lfh-sig", 0, "<I", [0x04034B4F, 0x02014B50, 0]),
        ("lfh-flags", 6, "<H", [0x0008, 0x0001, 0x0041, 0xFFFF]),
        ("lfh-method", 8, "<H", [0, 8, 9, 99]),
        ("lfh-crc", 14, "<I", [0, 0xFFFFFFFF]),
        ("lfh-csize", 18, "<I", [0, 1, 0xFFFFFFFF, 1 << 20]),
        ("lfh-usize", 22, "<I", [0, 0xFFFFFFFF]),
        ("lfh-namelen", 26, "<H", [0, 1, 100, 0xFFFF]),
        ("lfh-extralen", 28, "<H", [0, 64, 4096, 0xFFFF]),
    ]
    for index, (_central, local, _n, _e, _c, _cs) in enumerate(records[:6]):
        for label, offset, fmt, values in local_fields:
            for value in values:
                position = local + offset
                if position + struct.calcsize(fmt) > len(base):
                    continue
                mutated = bytearray(base)
                struct.pack_into(fmt, mutated, position, value)
                produced.append(("%s-e%d-%08x" % (label, index, value), bytes(mutated)))

    # (c) single-field corruption of central headers
    central_fields = [
        ("cdh-flags", 8, "<H", [0x0008, 0x0001, 0]),
        ("cdh-method", 10, "<H", [0, 8, 99]),
        ("cdh-crc", 16, "<I", [0, 0xFFFFFFFF]),
        ("cdh-csize", 20, "<I", [0, 1, 0xFFFFFFFF]),
        ("cdh-usize", 24, "<I", [0, 0xFFFFFFFF]),
        ("cdh-localoff", 42, "<I", [0, 1, 30, 0xFFFFFFFE]),
    ]
    for index, (central, _local, _n, _e, _c, _cs) in enumerate(records[:6]):
        for label, offset, fmt, values in central_fields:
            for value in values:
                position = central + offset
                if position + struct.calcsize(fmt) > len(base):
                    continue
                mutated = bytearray(base)
                struct.pack_into(fmt, mutated, position, value)
                produced.append(("%s-e%d-%08x" % (label, index, value), bytes(mutated)))

    # (d) point every central record at the first record's local header, so
    #     several records claim one span
    if len(records) >= 2:
        mutated = bytearray(base)
        first_local = records[0][1]
        for central, _local, _n, _e, _c, _cs in records[1:]:
            struct.pack_into("<I", mutated, central + 42, first_local)
        produced.append(("all-share-first-local", bytes(mutated)))

    # (e) duplicated and reordered central records
    eocd, directory, records = layout
    blocks = []
    for index, (central, _local, name_len, extra_len, comment_len, _cs) in enumerate(records):
        size = 46 + name_len + extra_len + comment_len
        blocks.append(base[central : central + size])
    if blocks:
        tail = base[eocd:]
        head = base[:directory]

        def rebuild(order, count=None):
            central_bytes = b"".join(blocks[i] for i in order)
            body = head + central_bytes
            record = bytearray(tail)
            total = len(order) if count is None else count
            struct.pack_into("<HH", record, 8, total & 0xFFFF, total & 0xFFFF)
            struct.pack_into("<I", record, 12, len(central_bytes))
            struct.pack_into("<I", record, 16, directory)
            return body + bytes(record)

        produced.append(("central-dup-first", rebuild([0] + list(range(len(blocks))))))
        produced.append(("central-reversed", rebuild(list(reversed(range(len(blocks)))))))
        if len(blocks) >= 3:
            order = list(range(len(blocks)))
            rng.shuffle(order)
            produced.append(("central-shuffled", rebuild(order)))
        produced.append(("central-count-plus-one", rebuild(list(range(len(blocks))), len(blocks) + 1)))
        produced.append(("central-count-minus-one", rebuild(list(range(len(blocks))), max(0, len(blocks) - 1))))

    # (f) EOCD field corruption
    for label, offset, fmt, values in [
        ("eocd-count", 10, "<H", [0, 1, 0xFFFF]),
        ("eocd-size", 12, "<I", [0, 0xFFFFFFFF]),
        ("eocd-offset", 16, "<I", [0, 30, 0xFFFFFFFE]),
    ]:
        for value in values:
            mutated = bytearray(base)
            if eocd + offset + struct.calcsize(fmt) <= len(mutated):
                struct.pack_into(fmt, mutated, eocd + offset, value)
                produced.append(("%s-%08x" % (label, value), bytes(mutated)))

    # (g) a handful of seeded byte flips inside the framing regions only, so the
    #     input still reaches the strict path rather than dying in inflate
    framing = []
    for central, local, name_len, extra_len, comment_len, _cs in records[:8]:
        framing.extend(range(local, min(local + 30 + name_len + extra_len, len(base))))
        framing.extend(range(central, min(central + 46 + name_len + extra_len + comment_len, len(base))))
    framing = sorted(set(framing))
    for step in range(12):
        if not framing:
            break
        mutated = bytearray(base)
        position = framing[rng.randrange(len(framing))]
        mutated[position] ^= 1 << (step % 8)
        produced.append(("framing-flip-%d" % step, bytes(mutated)))

    return produced


def crafted_archives():
    """Hand-built archives aimed squarely at the strict-layout proof."""
    out = {}

    # --- change 0575's approved witness -----------------------------------
    # A.bin's LOCAL extra field is 4096 bytes while its CENTRAL record declares
    # zero, so A's declared span swallows B's whole local record and only A's
    # local header reveals it.
    witness = build_zip(
        [
            Member("A.bin", b"A" * 32, local_extra=struct.pack("<HH", 0xFACE, 4092) + b"\x00" * 4092),
            Member("B.bin", b"B" * 32, local_offset=2048),
            Member("C.bin", b"C" * 32, local_offset=4163),
        ],
        directory_offset=4230,
    )
    out["witness-0575"] = witness

    # The same shape with the overlap moved, so the target sits before, inside
    # and after the overlapping pair.
    out["overlap-swallow-next"] = build_zip(
        [
            Member("first.bin", b"0" * 16),
            Member("swallower.bin", b"1" * 16, local_extra=b"\x00" * 512),
            Member("swallowed.bin", b"2" * 16, local_offset=200),
            Member("after.bin", b"3" * 16, local_offset=1400),
        ]
    )
    out["overlap-swallow-two"] = build_zip(
        [
            Member("big.bin", b"0" * 16, local_extra=b"\x00" * 16384),
            Member("victim1.bin", b"1" * 16, local_offset=4096),
            Member("victim2.bin", b"2" * 16, local_offset=8192),
            Member("victim3.bin", b"3" * 16, local_offset=12288),
            Member("clear.bin", b"4" * 16, local_offset=17000),
        ]
    )

    # --- the residual-window boundary -------------------------------------
    # A predecessor whose LOCAL name length is inflated declares a span that
    # reaches past the bracket a central-name-length derivation would compute.
    out["predecessor-inflated-local-name"] = build_zip(
        [
            Member("short.bin", b"x" * 32, local_name=b"short.bin", local_name_len=100,
                   local_extra=b"\x00" * 65535, local_extra_len=65535),
            Member("target.bin", b"y" * 32, local_offset=65640),
        ]
    )
    # Just inside and just outside the 131,094-byte sound window.
    for label, distance in (("inside", 131_000), ("outside", 140_000)):
        out["predecessor-window-%s" % label] = build_zip(
            [
                Member("pred.bin", b"x" * 32),
                Member("target.bin", b"y" * 32, local_offset=distance),
            ]
        )
    # A predecessor whose central compressed size alone already reaches the
    # target: refutable with no read at all.
    out["predecessor-csize-reaches"] = build_zip(
        [
            Member("pred.bin", b"x" * 4096, central_csize=8192),
            Member("target.bin", b"y" * 32, local_offset=6000),
        ]
    )

    # --- spans reaching into the central directory ------------------------
    out["span-into-central-directory"] = build_zip(
        [
            Member("a.bin", b"a" * 16),
            Member("last.bin", b"b" * 16, local_extra=b"\x00" * 256, central_csize=400),
        ]
    )
    out["predecessor-span-into-directory"] = build_zip(
        [
            Member("reaching.bin", b"a" * 16, central_csize=4096),
            Member("later.bin", b"b" * 16, local_offset=8192),
        ]
    )

    # --- descriptor-bearing members at the boundaries ---------------------
    for flavour in ("unsigned", "signed", "zip64"):
        out["descriptor-gapless-%s" % flavour] = build_zip(
            [
                Member("one.bin", b"1" * 24, descriptor=flavour),
                Member("two.bin", b"2" * 24, descriptor=flavour),
                Member("three.bin", b"3" * 24, descriptor=flavour),
            ]
        )
        # A descriptor-bearing predecessor whose payload ends exactly at the
        # target's local header: the widest-descriptor reservation would refuse,
        # the exact resolution must accept.
        out["descriptor-boundary-%s" % flavour] = build_zip(
            [
                Member("pred.bin", b"1" * 24, descriptor=flavour),
                Member("target.bin", b"2" * 24, descriptor=flavour),
            ]
        )
        out["descriptor-overlap-%s" % flavour] = build_zip(
            [
                Member("pred.bin", b"1" * 24, descriptor=flavour, local_extra=b"\x00" * 128),
                Member("target.bin", b"2" * 24, descriptor=flavour, local_offset=80),
            ]
        )

    # --- duplicate and out-of-order local header offsets ------------------
    out["duplicate-local-offsets"] = build_zip(
        [
            Member("one.bin", b"1" * 16),
            Member("two.bin", b"2" * 16, local_offset=0),
        ]
    )
    out["offset-into-payload"] = build_zip(
        [
            Member("host.bin", b"H" * 512),
            Member("parasite.bin", b"P" * 16, local_offset=100),
        ]
    )

    # --- exact adjacency, one byte of slack, one byte of overlap ----------
    base = build_zip([Member("a.bin", b"a" * 32), Member("b.bin", b"b" * 32)])
    out["adjacent-exact"] = base
    layout = parse_layout(base)
    if layout:
        _eocd, _directory, records = layout
        second_local = records[1][1]
        out["adjacent-gap-one"] = build_zip(
            [Member("a.bin", b"a" * 32), Member("b.bin", b"b" * 32, local_offset=second_local + 1)]
        )
        out["adjacent-overlap-one"] = build_zip(
            [
                Member("a.bin", b"a" * 32, local_extra=b"\x00" * 1),
                Member("b.bin", b"b" * 32, local_offset=second_local),
            ]
        )

    # --- directory records participating ----------------------------------
    out["directory-record-overlap"] = build_zip(
        [
            Member("dir/", b"", local_extra=b"\x00" * 256),
            Member("dir/file.bin", b"d" * 16, local_offset=64),
            Member("other.bin", b"o" * 16, local_offset=600),
        ]
    )

    # --- local/central disagreement on a record nobody reads --------------
    out["neighbour-name-mismatch"] = build_zip(
        [
            Member("alpha.bin", b"a" * 16, local_name=b"DIFFERENT.bin"),
            Member("target.bin", b"t" * 16),
        ]
    )
    out["neighbour-crc-mismatch"] = build_zip(
        [
            Member("alpha.bin", b"a" * 16, local_crc=0xDEADBEEF),
            Member("target.bin", b"t" * 16),
        ]
    )
    out["neighbour-size-mismatch"] = build_zip(
        [
            Member("alpha.bin", b"a" * 16, local_csize=999),
            Member("target.bin", b"t" * 16),
        ]
    )
    out["neighbour-method-mismatch"] = build_zip(
        [
            Member("alpha.bin", b"a" * 16, local_method=8),
            Member("target.bin", b"t" * 16),
        ]
    )
    out["neighbour-flags-mismatch"] = build_zip(
        [
            Member("alpha.bin", b"a" * 16, local_flags=0x0800),
            Member("target.bin", b"t" * 16),
        ]
    )
    out["neighbour-bad-signature"] = build_zip(
        [
            Member("alpha.bin", b"a" * 16),
            Member("target.bin", b"t" * 16),
        ]
    )
    corrupted = bytearray(out["neighbour-bad-signature"])
    corrupted[3] = 0x05
    out["neighbour-bad-signature"] = bytes(corrupted)
    out["neighbour-encrypted-flag"] = build_zip(
        [
            Member("alpha.bin", b"a" * 16, flags=0x0001),
            Member("target.bin", b"t" * 16),
        ]
    )

    # --- local-header size smuggling --------------------------------------
    # A predecessor whose LOCAL header declares a compressed size far larger
    # than its central record.  A reader that trusts local headers places this
    # record's payload across the target; a reader that trusts the central
    # directory does not.  The target-scoped neighbour bound takes the payload
    # size from the CENTRAL record, so it never sees the local claim.
    out["neighbour-local-csize-smuggles"] = build_zip(
        [
            Member("pred.bin", b"P" * 16, local_csize=100_000),
            Member("target.bin", b"T" * 16, local_offset=200),
        ]
    )
    out["neighbour-local-csize-smuggles-deflate"] = build_zip(
        [
            Member("pred.bin", b"P" * 512, method=8, local_csize=100_000),
            Member("target.txt", b"secret payload " * 32, method=8, local_offset=400),
        ]
    )
    # The mirror image: the LOCAL header understates.  The bound uses the
    # central size, which is the larger of the two, so this must still refuse.
    out["neighbour-local-csize-understates"] = build_zip(
        [
            Member("pred.bin", b"P" * 16, central_csize=4096, local_csize=16),
            Member("target.bin", b"T" * 16, local_offset=200),
        ]
    )
    out["neighbour-local-usize-smuggles"] = build_zip(
        [
            Member("pred.bin", b"P" * 16, local_usize=100_000),
            Member("target.bin", b"T" * 16, local_offset=200),
        ]
    )

    # --- the extreme of the sound residual window --------------------------
    # 30 + 65535 (local name) + 65535 (local extra) + 32 = 131,132, the largest
    # span a u16 pair can describe for a 32-byte payload.  A target inside it
    # must be refused; one past it may be accepted.
    for label, offset in (("reaches", 131_100), ("clear", 131_140)):
        out["window-max-residual-%s" % label] = build_zip(
            [
                Member(
                    "pred.bin",
                    b"P" * 32,
                    local_name=b"pred.bin",
                    local_name_len=65535,
                    local_extra=b"\x00" * 65535,
                    local_extra_len=65535,
                ),
                Member("target.bin", b"T" * 32, local_offset=offset),
            ]
        )
    # Local extra field inflated rather than the name: the bound reads both
    # halves of the variable region, so this must refuse.
    out["neighbour-local-extra-inflated"] = build_zip(
        [
            Member("pred.bin", b"P" * 16, local_extra_len=4096),
            Member("target.bin", b"T" * 16, local_offset=1024),
        ]
    )

    # --- deflate members, so the accepted path decompresses ---------------
    out["deflate-pair"] = build_zip(
        [
            Member("a.txt", b"hello world " * 64, method=8),
            Member("b.txt", b"second member " * 64, method=8),
        ]
    )
    out["deflate-overlap"] = build_zip(
        [
            Member("a.txt", b"hello world " * 64, method=8, local_extra=b"\x00" * 1024),
            Member("b.txt", b"second member " * 64, method=8, local_offset=200),
            Member("c.txt", b"third member " * 64, method=8, local_offset=2200),
        ]
    )

    # --- degenerate shapes -------------------------------------------------
    out["empty-archive"] = build_zip([])
    out["single-member"] = build_zip([Member("only.bin", b"x" * 8)])
    out["prefix-bytes"] = b"\x00" * 64 + build_zip([Member("only.bin", b"x" * 8)])
    out["many-tiny"] = build_zip([Member("m%03d.bin" % i, bytes([i & 0xFF]) * 4) for i in range(200)])
    out["many-tiny-one-reaching"] = build_zip(
        [Member("m000.bin", b"0" * 4, central_csize=4096)]
        + [Member("m%03d.bin" % i, bytes([i & 0xFF]) * 4) for i in range(1, 200)]
    )

    return out


def main(repo_root, out_dir):
    rng = random.Random(RNG_SEED)
    manifest = {"rng_seed": RNG_SEED, "families": {}, "files": []}

    seen = set()
    written_paths = set()

    def emit(family, name, data):
        if len(data) == 0:
            return None
        digest = hashlib.sha256(data).hexdigest()
        if digest in seen:
            return None
        seen.add(digest)
        key = (family, name)
        assert key not in written_paths, "corpus path collision: %s/%s" % key
        written_paths.add(key)
        path = write(out_dir, family, name, data)
        manifest["files"].append(
            {
                "family": family,
                "name": name,
                "bytes": len(data),
                "sha256": digest,
            }
        )
        return path

    # (a) retained seeds
    seed_root = os.path.join(repo_root, "docs/performance/results/change-0457/fuzz/seeds")
    seeds = []
    for directory, _dirs, names in os.walk(seed_root):
        for name in sorted(names):
            path = os.path.join(directory, name)
            with open(path, "rb") as handle:
                data = handle.read()
            if is_zip_container(data):
                seeds.append((os.path.relpath(path, seed_root).replace(os.sep, "-"), data))
    seeds.sort(key=lambda row: row[0])
    for name, data in seeds:
        emit("seeds", name, data)
    manifest["families"]["seeds"] = len(seeds)

    # (b) every real ZIP container in test-data
    testdata = collect_testdata(os.path.join(repo_root, "test-data"))
    for relpath, data in testdata:
        emit("testdata", relpath.replace(os.sep, "__"), data)
    manifest["families"]["testdata"] = len(testdata)

    # (c) crafted archives
    crafted = crafted_archives()
    for name in sorted(crafted):
        emit("crafted", name + ".zip", crafted[name])
    manifest["families"]["crafted"] = len(crafted)

    # (d) deterministic mutations of a fixed base set:
    #     every seed, every crafted archive, and every test-data archive under
    #     48 KiB (sorted, so the selection is reproducible).
    bases = [("seed-" + name, data) for name, data in seeds]
    bases += [("crafted-" + name, crafted[name]) for name in sorted(crafted)]
    # A deterministic, bounded slice of the small real archives.  Sorted by
    # path, filtered by size, then a fixed stride so the selection is a pure
    # function of the corpus and not of how many fixtures happen to exist.
    small = [(relpath, data) for relpath, data in testdata if len(data) <= 24 * 1024]
    stride = max(1, len(small) // MUTATION_TESTDATA_BASES)
    picked = small[::stride][:MUTATION_TESTDATA_BASES]
    bases += [("td-" + relpath.replace(os.sep, "__"), data) for relpath, data in picked]

    mutation_count = 0
    for base_name, base_data in bases:
        if len(base_data) > MAX_BYTES:
            continue
        for label, data in mutate(rng, base_data):
            if len(data) > MAX_BYTES:
                continue
            if emit("mutations", "%s/%s.zip" % (base_name, label), data):
                mutation_count += 1
    manifest["families"]["mutations"] = mutation_count
    manifest["mutation_bases"] = len(bases)

    manifest["total"] = len(manifest["files"])
    manifest_path = out_dir.rstrip("/") + "-manifest.json"
    with open(manifest_path, "w") as handle:
        json.dump(manifest, handle, indent=1, sort_keys=True)
    print(json.dumps({k: v for k, v in manifest.items() if k != "files"}, indent=1, sort_keys=True))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
