#!/usr/bin/env python3
"""Make the two bounded XLSX column-promotion guard fixtures.

The pinned workbook is copied at the ZIP-record level.  Only the local and
central payload metadata for ``xl/worksheets/sheet1.xml`` and that member's
deflated bytes are replaced; every other member payload is retained byte for
byte.  This script is run once before comparative capture and leaves a
manifest that makes that property replayable without regenerating anything.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
from pathlib import Path
import struct
import sys
import zlib
import zipfile


P = Path(__file__).resolve().parent
ROOT = P.parents[4]
ORIGINAL = ROOT / "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx"
SHEET_MEMBER = "xl/worksheets/sheet1.xml"
ORIGINAL_BYTES = 8_435
ORIGINAL_SHA256 = "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4"
ORIGINAL_SHEET_BYTES = 1_496
ORIGINAL_SHEET_SHA256 = "d73849d2d2d96ac61289e88224be3f6933ab5e671c2e29d55e36b1839b3e3ba7"
BASE = "7eeaab48c0281b06f53527d4c4f4ea79050d27e7"
EOCD = b"PK\x05\x06"
CENTRAL = b"PK\x01\x02"
LOCAL = b"PK\x03\x04"
CENTRAL_HEADER = struct.Struct("<4s6H3I5H2I")
LOCAL_HEADER = struct.Struct("<4s5H3I2H")


class FixtureError(ValueError):
    """The pinned source archive is not the expected bounded ZIP input."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise FixtureError(message)


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def write_once(path: Path, value: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write(value)


def _decode_name(raw: bytes, flags: int) -> str:
    return raw.decode("utf-8" if flags & 0x800 else "cp437")


def central_records(data: bytes) -> tuple[list[dict[str, object]], int, int, bytes]:
    """Return central-directory records and the original EOCD template."""
    eocd_offset = data.rfind(EOCD)
    require(eocd_offset >= 0, "ZIP end-of-central-directory record is missing")
    require(eocd_offset + 22 <= len(data), "ZIP end-of-central-directory record is truncated")
    (
        _signature,
        disk,
        central_disk,
        entries_on_disk,
        entries,
        central_size,
        central_offset,
        comment_size,
    ) = struct.unpack_from("<4s4H2IH", data, eocd_offset)
    require(disk == central_disk == 0, "multi-disk ZIP is outside the guard")
    require(entries_on_disk == entries, "split ZIP central-directory count is inconsistent")
    require(entries < 0xFFFF and central_size < 0xFFFFFFFF, "ZIP64 is outside the guard")
    require(central_offset + central_size == eocd_offset, "ZIP central-directory bounds changed")
    require(eocd_offset + 22 + comment_size == len(data), "ZIP trailing bytes are not a comment")
    cursor = central_offset
    result: list[dict[str, object]] = []
    seen: set[str] = set()
    for _ in range(entries):
        require(data[cursor:cursor + 4] == CENTRAL, "ZIP central-directory signature is invalid")
        header = data[cursor:cursor + CENTRAL_HEADER.size]
        require(len(header) == CENTRAL_HEADER.size, "ZIP central-directory header is truncated")
        (
            _signature,
            _version_made,
            version_needed,
            flags,
            method,
            _mtime,
            _mdate,
            crc,
            compressed_size,
            uncompressed_size,
            name_size,
            extra_size,
            comment_len,
            disk_start,
            _internal_attributes,
            _external_attributes,
            local_offset,
        ) = CENTRAL_HEADER.unpack(header)
        end = cursor + CENTRAL_HEADER.size + name_size + extra_size + comment_len
        require(end <= eocd_offset, "ZIP central-directory entry is truncated")
        name_bytes = data[cursor + CENTRAL_HEADER.size:cursor + CENTRAL_HEADER.size + name_size]
        name = _decode_name(name_bytes, flags)
        require(name not in seen, f"duplicate ZIP member {name!r}")
        seen.add(name)
        require(disk_start == 0, "multi-disk ZIP member is outside the guard")
        result.append({
            "name": name,
            "name_bytes": name_bytes,
            "central_offset": cursor,
            "central_bytes": data[cursor:end],
            "local_offset": local_offset,
            "flags": flags,
            "method": method,
            "crc": crc,
            "compressed_size": compressed_size,
            "uncompressed_size": uncompressed_size,
        })
        cursor = end
    require(cursor == eocd_offset, "ZIP central-directory size does not match its entries")
    eocd = data[eocd_offset:eocd_offset + 22 + comment_size]
    return result, central_offset, central_size, eocd


def local_payload(data: bytes, record: dict[str, object]) -> tuple[bytes, bytes]:
    """Return local-header bytes and the exact stored member payload."""
    offset = int(record["local_offset"])
    require(data[offset:offset + 4] == LOCAL, f"local header missing for {record['name']!r}")
    header = data[offset:offset + LOCAL_HEADER.size]
    require(len(header) == LOCAL_HEADER.size, "ZIP local header is truncated")
    (
        _signature,
        _version_needed,
        flags,
        method,
        _mtime,
        _mdate,
        crc,
        compressed_size,
        uncompressed_size,
        name_size,
        extra_size,
    ) = LOCAL_HEADER.unpack(header)
    require(flags == int(record["flags"]), f"ZIP flags disagree for {record['name']!r}")
    require(method == int(record["method"]), f"ZIP method disagrees for {record['name']!r}")
    require(not flags & 0x08, "ZIP data descriptors are outside the guard")
    name_start = offset + LOCAL_HEADER.size
    name_end = name_start + name_size
    extra_end = name_end + extra_size
    payload_end = extra_end + int(record["compressed_size"])
    require(payload_end <= len(data), f"ZIP local payload is truncated for {record['name']!r}")
    require(data[name_start:name_end] == bytes(record["name_bytes"]),
            f"local and central names disagree for {record['name']!r}")
    require(crc == int(record["crc"])
            and compressed_size == int(record["compressed_size"])
            and uncompressed_size == int(record["uncompressed_size"]),
            f"local and central sizes disagree for {record['name']!r}")
    return data[offset:extra_end], data[extra_end:payload_end]


def member_facts(data: bytes) -> list[dict[str, object]]:
    records, _central_offset, _central_size, _eocd = central_records(data)
    with zipfile.ZipFile(io.BytesIO(data), "r") as archive:
        infos = archive.infolist()
        require([info.filename for info in infos] == [str(record["name"]) for record in records],
                "ZIP reader order does not match the central directory")
        facts = []
        for record, info in zip(records, infos):
            _header, compressed = local_payload(data, record)
            payload = archive.read(info)
            require(len(payload) == int(record["uncompressed_size"]),
                    f"ZIP uncompressed size disagrees for {record['name']!r}")
            facts.append({
                "name": str(record["name"]),
                "method": int(record["method"]),
                "compressed_bytes": len(compressed),
                "compressed_sha256": sha_bytes(compressed),
                "payload_bytes": len(payload),
                "payload_sha256": sha_bytes(payload),
            })
        return facts


def raw_deflate(payload: bytes) -> bytes:
    compressor = zlib.compressobj(level=9, method=zlib.DEFLATED, wbits=-15)
    return compressor.compress(payload) + compressor.flush()


def replace_member(data: bytes, target: str, payload: bytes) -> bytes:
    """Copy a ZIP while replacing one member's compressed payload."""
    records, _central_offset, _central_size, eocd = central_records(data)
    output = bytearray()
    rewritten: list[tuple[dict[str, object], int, int, int, int]] = []
    replaced = False
    for record in records:
        local_header, compressed = local_payload(data, record)
        name = str(record["name"])
        new_compressed = compressed
        new_payload = None
        if name == target:
            require(not replaced, f"duplicate replacement target {target!r}")
            replaced = True
            new_payload = payload
            new_compressed = raw_deflate(payload)
            local_header = bytearray(local_header)
            struct.pack_into(
                "<III",
                local_header,
                14,
                zlib.crc32(payload) & 0xFFFFFFFF,
                len(new_compressed),
                len(payload),
            )
            local_header = bytes(local_header)
        local_offset = len(output)
        output.extend(local_header)
        output.extend(new_compressed)
        rewritten.append((record, local_offset, len(new_compressed),
                          len(payload) if new_payload is not None else int(record["uncompressed_size"]),
                          zlib.crc32(payload) & 0xFFFFFFFF if new_payload is not None else int(record["crc"])))
    central_offset = len(output)
    for record, local_offset, compressed_size, uncompressed_size, crc in rewritten:
        central = bytearray(bytes(record["central_bytes"]))
        if str(record["name"]) == target:
            struct.pack_into("<III", central, 16, crc, compressed_size, uncompressed_size)
        struct.pack_into("<I", central, 42, local_offset)
        output.extend(central)
    central_size = len(output) - central_offset
    updated_eocd = bytearray(eocd)
    struct.pack_into("<II", updated_eocd, 12, central_size, central_offset)
    output.extend(updated_eocd)
    result = bytes(output)
    require(replaced, f"ZIP replacement target is missing: {target!r}")
    # The only changed member is checked at the payload level by the manifest;
    # this additional ZIP-open check catches malformed offsets immediately.
    with zipfile.ZipFile(io.BytesIO(result), "r") as archive:
        require(archive.namelist() == [str(record["name"]) for record in records],
                "rewritten ZIP member order changed")
        require(archive.read(target) == payload, "rewritten target payload changed")
    return result


def sheet_fragment(xml: bytes) -> tuple[bytes, int, int]:
    start = xml.find(b"<cols>")
    end_marker = b"</cols>"
    require(start >= 0, "pinned worksheet has no cols element")
    end = xml.find(end_marker, start)
    require(end >= 0, "pinned worksheet cols element is not closed")
    end += len(end_marker)
    return xml[start:end], start, end


def replacement_xml(original: bytes, fragment: bytes) -> bytes:
    old, start, end = sheet_fragment(original)
    result = original[:start] + fragment + original[end:]
    require(result[:start] == original[:start] and result[start + len(fragment):] == original[end:],
            "worksheet replacement changed bytes outside cols")
    require(result.count(b"<cols>") == 1 and result.count(b"</cols>") == 1,
            "worksheet replacement has an unexpected cols count")
    # These markers are the preservation boundary for the producer fixture.
    require(result.count(b"<c ") == 8, "worksheet cell count changed")
    for marker in (b"mc:Ignorable=\"x14ac\"", b"<sheetData>", b"<autoFilter ref=\"A1:B4\">"):
        require(marker in result, f"worksheet preservation marker disappeared: {marker!r}")
    require(old == b'<cols><col min="2" max="2" width="19.5703125" customWidth="1"/></cols>',
            "pinned worksheet cols fragment changed")
    return result


def variant_specs() -> list[dict[str, object]]:
    disjoint = (b'<cols><col min="2" max="2" width="19.5703125" customWidth="1"/>'
                b'<col min="4" max="4" width="12" customWidth="1"/></cols>')
    records = []
    for index in range(1, 129):
        width = 9 + index
        records.append(
            f'<col min="{index}" max="16384" width="{width}" style="1" hidden="1" '
            'bestFit="1" customWidth="1" phonetic="1" outlineLevel="1" collapsed="1"/>'
        )
    wide = ("<cols>" + "".join(records) + "</cols>").encode("ascii")
    return [
        {
            "id": "disjoint-two-records",
            "filename": "disjoint-two-records.xlsx",
            "records": 2,
            "description": "the pinned record plus one disjoint complete record at column D",
            "fragment": disjoint,
        },
        {
            "id": "overlap-128-wide-complete",
            "filename": "overlap-128-wide-complete.xlsx",
            "records": 128,
            "description": "128 valid complete records with nested wide overlap through column XFD",
            "fragment": wide,
        },
    ]


def descriptor(path: Path, *, relative_to: Path = P) -> dict[str, object]:
    return {
        "path": str(path.relative_to(relative_to)),
        "bytes": path.stat().st_size,
        "sha256": sha_file(path),
    }


def fixture_manifest() -> dict[str, object]:
    require(ORIGINAL.is_file() and not ORIGINAL.is_symlink(), f"missing original fixture: {ORIGINAL}")
    original = ORIGINAL.read_bytes()
    require(len(original) == ORIGINAL_BYTES and sha_bytes(original) == ORIGINAL_SHA256,
            "pinned original fixture changed")
    original_facts = member_facts(original)
    names = [str(row["name"]) for row in original_facts]
    require(names.count(SHEET_MEMBER) == 1, "pinned fixture has an unexpected sheet1 member count")
    with zipfile.ZipFile(io.BytesIO(original), "r") as archive:
        original_sheet = archive.read(SHEET_MEMBER)
    require(len(original_sheet) == ORIGINAL_SHEET_BYTES
            and sha_bytes(original_sheet) == ORIGINAL_SHEET_SHA256,
            "pinned sheet1 payload changed")
    old_fragment, _start, _end = sheet_fragment(original_sheet)
    require(old_fragment == b'<cols><col min="2" max="2" width="19.5703125" customWidth="1"/></cols>',
            "pinned sheet1 cols fragment changed")
    variants = []
    for spec in variant_specs():
        fragment = bytes(spec["fragment"])
        sheet = replacement_xml(original_sheet, fragment)
        output = replace_member(original, SHEET_MEMBER, sheet)
        output_path = P / "fixtures" / str(spec["filename"])
        write_once(output_path, output)
        output_facts = member_facts(output)
        require([row["name"] for row in output_facts] == names,
                f"{spec['id']} changed ZIP member names or order")
        by_name = {str(row["name"]): row for row in output_facts}
        for name in names:
            if name == SHEET_MEMBER:
                continue
            before = next(row for row in original_facts if row["name"] == name)
            after = by_name[name]
            require(before["compressed_sha256"] == after["compressed_sha256"]
                    and before["payload_sha256"] == after["payload_sha256"]
                    and before["compressed_bytes"] == after["compressed_bytes"]
                    and before["payload_bytes"] == after["payload_bytes"],
                    f"untouched ZIP member payload changed: {name}")
        with zipfile.ZipFile(io.BytesIO(output), "r") as archive:
            output_sheet = archive.read(SHEET_MEMBER)
        require(output_sheet == sheet, f"{spec['id']} sheet payload changed while writing")
        variants.append({
            "id": spec["id"],
            "description": spec["description"],
            "records": spec["records"],
            "cols_fragment_bytes": len(fragment),
            "cols_fragment_sha256": sha_bytes(fragment),
            "sheet_payload_bytes": len(sheet),
            "sheet_payload_sha256": sha_bytes(sheet),
            "archive": descriptor(output_path),
            "members": output_facts,
            "untouched_members": [name for name in names if name != SHEET_MEMBER],
        })
    return {
        "schema": "litchi.performance.0832.column-promotion-fixtures.v1",
        "base": BASE,
        "generator": descriptor(P / "generate_fixtures.py"),
        "original": {
            **descriptor(ORIGINAL, relative_to=ROOT),
            "member": SHEET_MEMBER,
            "sheet_payload_bytes": len(original_sheet),
            "sheet_payload_sha256": sha_bytes(original_sheet),
            "cols_fragment_bytes": len(old_fragment),
            "cols_fragment_sha256": sha_bytes(old_fragment),
            "members": original_facts,
        },
        "preservation_contract": {
            "archive_copy": "raw ZIP local and central records are copied; only sheet1.xml payload and its CRC/sizes are replaced",
            "untouched_member_payloads": "compressed and uncompressed payload bytes must match the original for every member except xl/worksheets/sheet1.xml",
            "worksheet_outside_cols": "sheet1.xml bytes before and after the cols element remain identical",
            "cells": "the pinned eight cell elements, MCE declaration, shared strings, date styles, and autofilter markers remain present",
        },
        "variants": variants,
    }


def verify_manifest() -> dict[str, object]:
    path = P / "fixture-manifest.json"
    require(path.is_file() and not path.is_symlink(), "fixture manifest is missing")
    value = json.loads(path.read_text(encoding="utf-8"))
    require(value.get("schema") == "litchi.performance.0832.column-promotion-fixtures.v1",
            "fixture manifest schema changed")
    require(value.get("base") == BASE, "fixture manifest base changed")
    generator = value.get("generator")
    require(isinstance(generator, dict) and generator.get("sha256") == sha_file(P / "generate_fixtures.py"),
            "fixture generator changed after generation")
    original = value.get("original")
    require(isinstance(original, dict) and original.get("sha256") == ORIGINAL_SHA256,
            "fixture original descriptor changed")
    for variant in value.get("variants", []):
        archive = variant.get("archive")
        require(isinstance(archive, dict), "fixture variant archive descriptor is malformed")
        output = P / str(archive.get("path"))
        require(output.is_file() and not output.is_symlink(), f"missing fixture archive: {output}")
        require(output.stat().st_size == archive.get("bytes") and sha_file(output) == archive.get("sha256"),
                f"fixture archive changed: {output}")
    return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="verify the retained fixture manifest")
    args = parser.parse_args(argv)
    try:
        if args.check:
            value = verify_manifest()
        else:
            require(not (P / "fixture-manifest.json").exists(), "fixture manifest already exists")
            value = fixture_manifest()
            write_once(P / "fixture-manifest.json",
                       (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8"))
        print(json.dumps({"status": "pass", "variants": len(value["variants"])}, sort_keys=True))
        return 0
    except (FixtureError, OSError, UnicodeError, ValueError, KeyError, zipfile.BadZipFile) as error:
        print(f"promotion fixture generation failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
