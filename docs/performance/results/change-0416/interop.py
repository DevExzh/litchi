#!/usr/bin/env python3
"""Generate and independently verify small ZIP descriptor interoperability corpora.

The fixtures are produced through Python's standard-library ``zipfile``
writer and a non-seekable output wrapper.  ``force_zip64=True`` therefore
produces the same small-file shape used by Python callers that do not know a
member's final size: a version 4.5 local header, ZIP64 size sentinels and a
16-byte zero placeholder, followed by a signed 64-bit data descriptor.  The
unsigned variants remove only the descriptor signature and repair the
following central-directory offsets.  The central+local fixture additionally
promotes the central records and terminal archive to ZIP64 metadata.

The generated manifest is an independent oracle.  ``verify`` drains every
member through ``zipfile`` and ``zlib``, records all sizes, CRCs and SHA-256
digests, and checks the exact fixture shape used by the Rust integration
tests.  A seekable ``force_zip64`` case keeps actual values in the local ZIP64
extra while retaining ordinary central sizes, and empty members pin the
zero-size descriptor boundary.  No soapberry code is imported.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import struct
import sys
import zipfile
import zlib


DESCRIPTOR_SIGNATURE = 0x08074B50
ZIP32_MAX = 0xFFFFFFFF
ZIP64_EXTRA_ID = 0x0001
CHUNK_BYTES = 4096


class SequentialSink:
    """A deliberately non-seekable sink accepted by ``zipfile.ZipFile``."""

    def __init__(self, output):
        self.output = output

    def write(self, data):
        return self.output.write(data)

    def flush(self):
        self.output.flush()


def fixed_info(name: str, method: int) -> zipfile.ZipInfo:
    info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
    info.compress_type = method
    info.create_system = 3
    info.create_version = 20
    info.extract_version = 20
    info.external_attr = 0o100644 << 16
    return info


def crc_signature_payload() -> bytes:
    """Return deterministic bytes whose CRC is the ZIP descriptor marker.

    An unsigned descriptor begins with its CRC.  Making that CRC equal to the
    optional signed-descriptor marker exercises the only framing ambiguity
    that cannot be resolved by looking at those first four bytes alone.
    """

    prefix = b"unsigned descriptor CRC marker | PK\x07\x08 | "
    target = DESCRIPTOR_SIGNATURE

    # CRC32 is affine in the four appended bytes.  Compute the influence of
    # each suffix bit and solve the resulting 32x32 binary linear system.
    base = zlib.crc32(prefix + b"\0\0\0\0") & ZIP32_MAX
    columns = []
    for bit in range(32):
        suffix = bytearray(4)
        suffix[bit // 8] = 1 << (bit % 8)
        columns.append((zlib.crc32(prefix + suffix) & ZIP32_MAX) ^ base)
    rhs = target ^ base

    rows = []
    for row in range(32):
        mask = 0
        for column, value in enumerate(columns):
            if value & (1 << row):
                mask |= 1 << column
        if rhs & (1 << row):
            mask |= 1 << 32
        rows.append(mask)

    pivot = 0
    for column in range(32):
        selected = next(
            (candidate for candidate in range(pivot, 32) if rows[candidate] & (1 << column)),
            None,
        )
        if selected is None:
            raise RuntimeError("CRC32 suffix system is singular")
        rows[pivot], rows[selected] = rows[selected], rows[pivot]
        for candidate in range(32):
            if candidate != pivot and rows[candidate] & (1 << column):
                rows[candidate] ^= rows[pivot]
        pivot += 1

    suffix = bytearray(4)
    for row in rows:
        coefficient_bits = row & ((1 << 32) - 1)
        if coefficient_bits.bit_count() != 1:
            raise RuntimeError("CRC32 suffix system did not reduce to identity")
        column = coefficient_bits.bit_length() - 1
        if row & (1 << 32):
            suffix[column // 8] |= 1 << (column % 8)
    payload = prefix + suffix
    if zlib.crc32(payload) & ZIP32_MAX != target:
        raise RuntimeError("CRC32 suffix construction failed")
    return payload


CRC_SIGNATURE_PAYLOAD = crc_signature_payload()
STORE_PAYLOAD = b"stored payload with a deterministic ZIP descriptor\n"
DEFLATE_PAYLOAD = (
    b"deflated payload: "
    + bytes(range(32))
    + b"\n"
    + (b"repeatable deflate text; " * 11)
)


def write_member(archive: zipfile.ZipFile, info: zipfile.ZipInfo, payload: bytes, force_zip64: bool) -> None:
    with archive.open(info, "w", force_zip64=force_zip64) as destination:
        for start in range(0, len(payload), CHUNK_BYTES):
            destination.write(payload[start : start + CHUNK_BYTES])


def write_archive(path: Path, members: list[tuple[str, int, bytes]], force_zip64: bool) -> None:
    with path.open("wb") as output:
        with zipfile.ZipFile(SequentialSink(output), "w", allowZip64=True) as archive:
            for name, method, payload in members:
                write_member(archive, fixed_info(name, method), payload, force_zip64)


def write_seekable_archive(path: Path, members: list[tuple[str, int, bytes]], force_zip64: bool) -> None:
    """Write through a normal seekable file so local ZIP64 values are final."""

    with path.open("w+b") as output:
        with zipfile.ZipFile(output, "w", allowZip64=True) as archive:
            for name, method, payload in members:
                write_member(archive, fixed_info(name, method), payload, force_zip64)


def u16(data: bytes | bytearray, offset: int) -> int:
    return struct.unpack_from("<H", data, offset)[0]


def u32(data: bytes | bytearray, offset: int) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def u64(data: bytes | bytearray, offset: int) -> int:
    return struct.unpack_from("<Q", data, offset)[0]


def put_u16(data: bytearray, offset: int, value: int) -> None:
    struct.pack_into("<H", data, offset, value)


def put_u32(data: bytearray, offset: int, value: int) -> None:
    struct.pack_into("<I", data, offset, value)


def put_u64(data: bytearray, offset: int, value: int) -> None:
    struct.pack_into("<Q", data, offset, value)


def zip32_eocd(data: bytes | bytearray) -> tuple[int, int, int, int]:
    position = data.rfind(b"PK\x05\x06")
    if position < 0:
        raise ValueError("ZIP32 EOCD is missing")
    if u16(data, position + 20) != 0:
        raise ValueError("fixture generator expects an empty ZIP comment")
    count = u16(data, position + 10)
    central_size = u32(data, position + 12)
    central_offset = u32(data, position + 16)
    if count != 0xFFFF and central_size != ZIP32_MAX and central_offset != ZIP32_MAX:
        return position, count, central_offset, central_offset + central_size

    locator = position - 20
    if locator < 0 or u32(data, locator) != 0x07064B50:
        raise ValueError("ZIP64 locator is missing")
    zip64_offset = u64(data, locator + 8)
    if zip64_offset > len(data) - 56 or u32(data, zip64_offset) != 0x06064B50:
        raise ValueError("ZIP64 EOCD is missing")
    count = u64(data, zip64_offset + 32)
    central_size = u64(data, zip64_offset + 40)
    central_offset = u64(data, zip64_offset + 48)
    return position, count, central_offset, central_offset + central_size


def central_records(data: bytes | bytearray) -> list[tuple[int, int, int, int, int, int]]:
    """Return (record, local_offset, compressed, name_len, extra_len, flags)."""

    eocd, count, central_offset, central_end = zip32_eocd(data)
    cursor = central_offset
    records = []
    for _ in range(count):
        if u32(data, cursor) != 0x02014B50:
            raise ValueError("central-directory signature is missing")
        name_len, extra_len, comment_len = (
            u16(data, cursor + 28),
            u16(data, cursor + 30),
            u16(data, cursor + 32),
        )
        end = cursor + 46 + name_len + extra_len + comment_len
        if end > central_end:
            raise ValueError("central-directory record is truncated")
        records.append(
            (
                cursor,
                u32(data, cursor + 42),
                u32(data, cursor + 20),
                name_len,
                extra_len,
                u16(data, cursor + 8),
            )
        )
        cursor = end
    if cursor != central_end:
        raise ValueError("central-directory size does not match its records")
    return records


def local_descriptor_position(data: bytes | bytearray, central: tuple[int, int, int, int, int, int]) -> int:
    _, local_offset, compressed, _, _, _ = central
    if u32(data, local_offset) != 0x04034B50:
        raise ValueError("local-header signature is missing")
    name_len, extra_len = u16(data, local_offset + 26), u16(data, local_offset + 28)
    return local_offset + 30 + name_len + extra_len + compressed


def remove_descriptor_signatures(data: bytes) -> bytes:
    """Remove signed descriptor markers and repair ZIP32 central offsets."""

    records = central_records(data)
    positions = []
    for record in records:
        descriptor = local_descriptor_position(data, record)
        if u32(data, descriptor) != DESCRIPTOR_SIGNATURE:
            raise ValueError("expected a signed descriptor in local-only fixture")
        positions.append(descriptor)

    output = bytearray(data)
    for position in reversed(positions):
        del output[position : position + 4]

    old_eocd, _, old_central_offset, _ = zip32_eocd(data)
    removed_before = 0
    for position in positions:
        if position < old_eocd:
            removed_before += 4
    new_eocd = old_eocd - removed_before
    new_central_offset = old_central_offset - removed_before
    put_u32(output, new_eocd + 16, new_central_offset)

    # Every central record's local offset moves by the descriptors removed
    # before it.  The records themselves remain byte-for-byte otherwise.
    for record, local_offset, _, _, _, _ in records:
        removed = sum(4 for position in positions if position < local_offset)
        new_record = record - sum(4 for position in positions if position < record)
        put_u32(output, new_record + 42, local_offset - removed)

    if len(output) != len(data) - 4 * len(positions):
        raise ValueError("descriptor removal changed an unexpected number of bytes")
    return bytes(output)


def promote_central_and_archive_to_zip64(
    data: bytes, members: list[tuple[str, int, bytes]]
) -> bytes:
    """Promote every central record and the terminal archive to ZIP64."""

    records = central_records(data)
    eocd, count, central_offset, central_end = zip32_eocd(data)
    central_size = central_end - central_offset
    output = bytearray(data)

    # All fixture records have empty extras.  Add only the prescribed ZIP64
    # uncompressed/compressed size pair; local offsets remain ZIP32-valid.
    for record, _, compressed, name_len, extra_len, _ in reversed(records):
        if extra_len != 0:
            raise ValueError("central ZIP64 promotion expects empty central extras")
        name_start = record + 46
        name_end = name_start + name_len
        name = output[name_start:name_end]
        payload = next(payload for candidate_name, _, payload in members if candidate_name.encode() == name)
        extra = struct.pack("<HHQQ", ZIP64_EXTRA_ID, 16, len(payload), compressed)
        output[record + 46 + name_len : record + 46 + name_len] = extra
        put_u16(output, record + 30, len(extra))
        put_u32(output, record + 20, ZIP32_MAX)
        put_u32(output, record + 24, ZIP32_MAX)
        put_u16(output, record + 6, 45)
        put_u16(output, record + 4, (3 << 8) | 45)

    # Inserting central extras moves the terminal records.  Recompute the
    # central end after insertion and append the ZIP64 EOCD/locator before a
    # ZIP32-compatible sentinel EOCD.
    inserted = len(records) * 20
    new_eocd = eocd + inserted
    new_central_end = central_end + inserted
    old_terminal = bytes(output[new_eocd : new_eocd + 22])
    if u32(old_terminal, 0) != 0x06054B50:
        raise ValueError("terminal EOCD moved unexpectedly")
    zip64_eocd = struct.pack(
        "<IQHHIIQQQQ",
        0x06064B50,
        44,
        45,
        45,
        0,
        0,
        count,
        count,
        central_size + inserted,
        central_offset,
    )
    zip64_locator = struct.pack("<IIQI", 0x07064B50, 0, new_eocd, 1)
    terminal = bytearray(old_terminal)
    put_u16(terminal, 8, ZIP32_MAX & 0xFFFF)
    put_u16(terminal, 10, ZIP32_MAX & 0xFFFF)
    put_u32(terminal, 12, ZIP32_MAX)
    put_u32(terminal, 16, ZIP32_MAX)
    output[new_eocd:] = zip64_eocd + zip64_locator + terminal
    return bytes(output)


def promote_central_records_to_zip64(
    data: bytes, members: list[tuple[str, int, bytes]]
) -> bytes:
    """Promote central records while retaining an ordinary ZIP32 terminal."""

    records = central_records(data)
    eocd, _, central_offset, central_end = zip32_eocd(data)
    central_size = central_end - central_offset
    output = bytearray(data)

    for record, _, _, name_len, extra_len, _ in reversed(records):
        if extra_len != 0:
            raise ValueError("central ZIP64 promotion expects empty central extras")
        name_start = record + 46
        name = output[name_start : name_start + name_len]
        payload = next(payload for candidate_name, _, payload in members if candidate_name.encode() == name)
        extra = struct.pack("<HHQQ", ZIP64_EXTRA_ID, 16, len(payload), u32(data, record + 20))
        output[record + 46 + name_len : record + 46 + name_len] = extra
        put_u16(output, record + 30, len(extra))
        put_u32(output, record + 20, ZIP32_MAX)
        put_u32(output, record + 24, ZIP32_MAX)
        put_u16(output, record + 6, 45)
        put_u16(output, record + 4, (3 << 8) | 45)

    inserted = len(records) * 20
    new_eocd = eocd + inserted
    terminal = bytes(output[new_eocd : new_eocd + 22])
    if u32(terminal, 0) != 0x06054B50:
        raise ValueError("terminal EOCD moved unexpectedly")
    put_u32(output, new_eocd + 12, central_size + inserted)
    put_u32(output, new_eocd + 16, central_offset)
    return bytes(output)


def parse_zip64_extra(extra: bytes, compressed: int, uncompressed: int) -> tuple[int | None, int | None, int | None]:
    cursor = 0
    resolved_uncompressed = None
    resolved_compressed = None
    resolved_offset = None
    while cursor + 4 <= len(extra):
        field_id, size = struct.unpack_from("<HH", extra, cursor)
        field = extra[cursor + 4 : cursor + 4 + size]
        if len(field) != size:
            raise ValueError("truncated extra field")
        if field_id == ZIP64_EXTRA_ID:
            field_cursor = 0
            if uncompressed == ZIP32_MAX:
                resolved_uncompressed = u64(field, field_cursor)
                field_cursor += 8
            if compressed == ZIP32_MAX:
                resolved_compressed = u64(field, field_cursor)
                field_cursor += 8
            if field_cursor < len(field):
                resolved_offset = u64(field, field_cursor)
        cursor += 4 + size
    return resolved_uncompressed, resolved_compressed, resolved_offset


EXPECTED_DESCRIPTOR_ENCODING = {
    "zip32-signed-store-deflate.zip": "signed",
    "zip64-local-only-signed.zip": "signed",
    "zip64-local-only-unsigned-crc-marker.zip": "unsigned",
    "zip64-central-local-signed.zip": "signed",
    "opc-local-only-signed.zip": "signed",
    "many-small-local-only-signed.zip": "signed",
    "zip64-local-only-seekable-no-descriptor.zip": None,
    "many-small-zip32-signed.zip": "signed",
    "many-small-central-local-signed.zip": "signed",
    "zip64-local-only-empty-signed.zip": "signed",
    "zip64-local-only-empty-unsigned.zip": "unsigned",
    "zip64-central-local-zip32-tail-signed.zip": "signed",
    "many-small-central-local-zip32-tail-signed.zip": "signed",
}


def descriptor_shape(
    data: bytes, info: zipfile.ZipInfo, expected_encoding: str | None
) -> dict[str, object]:
    records = central_records(data)
    record = next(record for record in records if data[record[0] + 46 : record[0] + 46 + record[3]].decode() == info.filename)
    record_offset, local_offset, _, _, central_extra_len, flags = record
    local_name_len, local_extra_len = u16(data, local_offset + 26), u16(data, local_offset + 28)
    local_version = u16(data, local_offset + 4)
    local_compressed_field = u32(data, local_offset + 18)
    local_uncompressed_field = u32(data, local_offset + 22)
    local_extra = data[local_offset + 30 + local_name_len : local_offset + 30 + local_name_len + local_extra_len]
    central_extra = data[record_offset + 46 + record[3] : record_offset + 46 + record[3] + central_extra_len]
    local_zip64_uncompressed, local_zip64_compressed, _ = parse_zip64_extra(
        local_extra, local_compressed_field, local_uncompressed_field
    )
    central_zip64_uncompressed, central_zip64_compressed, central_zip64_offset = parse_zip64_extra(
        central_extra, u32(data, record_offset + 20), u32(data, record_offset + 24)
    )
    has_descriptor = bool(flags & 0x0008)
    descriptor_marker_present = None
    if has_descriptor:
        descriptor = local_descriptor_position(
            data,
            (record_offset, local_offset, info.compress_size, record[3], local_extra_len, flags),
        )
        descriptor_marker_present = u32(data, descriptor) == DESCRIPTOR_SIGNATURE
    local_zip64 = local_compressed_field == ZIP32_MAX and local_uncompressed_field == ZIP32_MAX
    central_zip64 = u32(data, record_offset + 20) == ZIP32_MAX or u32(data, record_offset + 24) == ZIP32_MAX
    descriptor_width = 8 if has_descriptor and (local_zip64 or central_zip64) else (4 if has_descriptor else None)
    return {
        "name": info.filename,
        "method": "store" if info.compress_type == zipfile.ZIP_STORED else "deflate",
        "flags": flags,
        "local_version_needed": local_version,
        "central_version_needed": u16(data, record_offset + 6),
        "local_zip64": local_zip64,
        "central_zip64": central_zip64,
        "local_zip64_placeholder": (local_zip64_compressed, local_zip64_uncompressed) == (0, 0),
        "central_zip64_values": {
            "compressed": central_zip64_compressed,
            "uncompressed": central_zip64_uncompressed,
            "offset": central_zip64_offset,
        },
        "descriptor_encoding": expected_encoding,
        "descriptor_signature_present": descriptor_marker_present,
        "descriptor_width": descriptor_width,
        "local_name_bytes": local_name_len,
        "local_extra_bytes": local_extra_len,
        "central_extra_bytes": central_extra_len,
    }


def verify_archive(path: Path) -> dict[str, object]:
    data = path.read_bytes()
    archive_sha256 = hashlib.sha256(data).hexdigest()
    entries = []
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            crc = 0
            count = 0
            digest = hashlib.sha256()
            with archive.open(info) as source:
                while chunk := source.read(CHUNK_BYTES):
                    count += len(chunk)
                    crc = zlib.crc32(chunk, crc) & ZIP32_MAX
                    digest.update(chunk)
            if count != info.file_size or crc != info.CRC:
                raise ValueError(f"zipfile verification failed for {path.name}:{info.filename}")
            expected_encoding = EXPECTED_DESCRIPTOR_ENCODING[path.name]
            if bool(info.flag_bits & 0x0008) != (expected_encoding is not None):
                raise ValueError(f"unexpected descriptor flag for {path.name}:{info.filename}")
            shape = descriptor_shape(data, info, expected_encoding)
            shape.update(
                {
                    "uncompressed_bytes": count,
                    "compressed_bytes": info.compress_size,
                    "crc32": crc,
                    "sha256": digest.hexdigest(),
                }
            )
            entries.append(shape)
    return {
        "archive": path.name,
        "archive_bytes": len(data),
        "archive_sha256": archive_sha256,
        "entry_count": len(entries),
        "entries": entries,
        "zipfile_verified": True,
    }


CONTENT_TYPES = (
    b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Default Extension="xml" ContentType="application/xml"/>'
    b'<Default Extension="bin" ContentType="application/octet-stream"/>'
    b'<Override PartName="/word/document.xml" ContentType="application/xml"/>'
    b"</Types>"
)
ROOT_RELS = (
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    b'Target="word/document.xml"/></Relationships>'
)
DOCUMENT = b"<document>source descriptor fixture</document>"
OPAQUE = b"untouched opaque payload with a source-only ZIP metadata tail"


OPC_MEMBERS = [
    ("[Content_Types].xml", zipfile.ZIP_STORED, CONTENT_TYPES),
    ("_rels/.rels", zipfile.ZIP_STORED, ROOT_RELS),
    ("custom/opaque.bin", zipfile.ZIP_STORED, OPAQUE),
    ("word/document.xml", zipfile.ZIP_DEFLATED, DOCUMENT),
]

def build_fixtures(output_dir: Path) -> list[Path]:
    corpus = output_dir / "corpus"
    corpus.mkdir(parents=True, exist_ok=True)
    paths = []

    ordinary = [
        ("store.bin", zipfile.ZIP_STORED, STORE_PAYLOAD),
        ("deflate.bin", zipfile.ZIP_DEFLATED, DEFLATE_PAYLOAD),
    ]
    path = corpus / "zip32-signed-store-deflate.zip"
    write_archive(path, ordinary, force_zip64=False)
    paths.append(path)

    local_only = [
        ("store.bin", zipfile.ZIP_STORED, STORE_PAYLOAD),
        ("deflate.bin", zipfile.ZIP_DEFLATED, DEFLATE_PAYLOAD),
    ]
    path = corpus / "zip64-local-only-signed.zip"
    write_archive(path, local_only, force_zip64=True)
    paths.append(path)

    path = corpus / "zip64-local-only-unsigned-crc-marker.zip"
    unsigned = [
        ("crc-marker.bin", zipfile.ZIP_STORED, CRC_SIGNATURE_PAYLOAD),
        ("deflate.bin", zipfile.ZIP_DEFLATED, DEFLATE_PAYLOAD),
    ]
    write_archive(path, unsigned, force_zip64=True)
    path.write_bytes(remove_descriptor_signatures(path.read_bytes()))
    paths.append(path)

    path = corpus / "zip64-central-local-signed.zip"
    central_local = [
        ("store.bin", zipfile.ZIP_STORED, STORE_PAYLOAD),
        ("deflate.bin", zipfile.ZIP_DEFLATED, DEFLATE_PAYLOAD),
    ]
    write_archive(path, central_local, force_zip64=True)
    path.write_bytes(promote_central_and_archive_to_zip64(path.read_bytes(), central_local))
    paths.append(path)

    path = corpus / "opc-local-only-signed.zip"
    write_archive(path, OPC_MEMBERS, force_zip64=True)
    paths.append(path)

    many = [
        (f"parts/p{index:03d}.bin", zipfile.ZIP_STORED, bytes([index & 0xFF, 0xA5, 0x15, 0x41]))
        for index in range(256)
    ]
    path = corpus / "many-small-local-only-signed.zip"
    write_archive(path, many, force_zip64=True)
    paths.append(path)

    path = corpus / "zip64-local-only-seekable-no-descriptor.zip"
    seekable = [
        ("seekable-store.bin", zipfile.ZIP_STORED, STORE_PAYLOAD),
        ("seekable-deflate.bin", zipfile.ZIP_DEFLATED, DEFLATE_PAYLOAD),
    ]
    write_seekable_archive(path, seekable, force_zip64=True)
    paths.append(path)

    # Matched 256-member controls keep the local framing and the central
    # framing independently measurable.  They are appended after the first
    # seven fixtures so existing probe references remain byte-stable.
    path = corpus / "many-small-zip32-signed.zip"
    write_archive(path, many, force_zip64=False)
    paths.append(path)

    path = corpus / "many-small-central-local-signed.zip"
    write_archive(path, many, force_zip64=True)
    path.write_bytes(promote_central_and_archive_to_zip64(path.read_bytes(), many))
    paths.append(path)

    # Empty stored members exercise the zero-size boundary.  In particular,
    # an unsigned ZIP64 descriptor is 20 bytes even when all three fields are
    # zero; a parser that falls back to the 12-byte ZIP32 framing can otherwise
    # mistake the following central signature for valid data.
    empty = [("empty.bin", zipfile.ZIP_STORED, b"")]
    path = corpus / "zip64-local-only-empty-signed.zip"
    write_archive(path, empty, force_zip64=True)
    paths.append(path)

    path = corpus / "zip64-local-only-empty-unsigned.zip"
    write_archive(path, empty, force_zip64=True)
    path.write_bytes(remove_descriptor_signatures(path.read_bytes()))
    paths.append(path)

    path = corpus / "zip64-central-local-zip32-tail-signed.zip"
    write_archive(path, central_local, force_zip64=True)
    path.write_bytes(promote_central_records_to_zip64(path.read_bytes(), central_local))
    paths.append(path)

    path = corpus / "many-small-central-local-zip32-tail-signed.zip"
    write_archive(path, many, force_zip64=True)
    path.write_bytes(promote_central_records_to_zip64(path.read_bytes(), many))
    paths.append(path)
    return paths


def generate(output_dir: Path) -> None:
    paths = build_fixtures(output_dir)
    manifest = {
        "schema_version": 1,
        "generator": "change-0416/interop.py",
        "python": sys.version.split()[0],
        "zlib": zlib.ZLIB_RUNTIME_VERSION,
        "buffer_bytes": CHUNK_BYTES,
        "zipfile_verified_independently": True,
        "crc_signature": f"0x{DESCRIPTOR_SIGNATURE:08x}",
        "fixtures": [verify_archive(path) for path in paths],
    }
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "manifest.json").write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n")
    print(json.dumps(manifest, indent=2, sort_keys=True))


def verify(output_dir: Path) -> None:
    manifest = json.loads((output_dir / "manifest.json").read_text())
    for fixture in manifest["fixtures"]:
        path = output_dir / "corpus" / fixture["archive"]
        actual = verify_archive(path)
        if actual != fixture:
            raise SystemExit(f"manifest mismatch for {path}")
    print(f"OK: independently verified {len(manifest['fixtures'])} ZIP fixtures")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("generate", "verify"))
    parser.add_argument("--output-dir", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    if args.mode == "generate":
        generate(args.output_dir)
    else:
        verify(args.output_dir)


if __name__ == "__main__":
    main()
