#!/usr/bin/env python3
"""Independent raw ZIP census for the 0909 PPTX owning-input probe.

The helper deliberately has no dependency on Litchi or on the probe that
produced either archive.  ``zipfile`` is used only for the decoded readback;
local headers, central-directory records, data descriptors, compressed byte
ranges, and the EOCD comment are parsed from the archive bytes here.

The report is create-only.  A non-zero result means either that an archive is
malformed for this bounded parser or that a difference was found outside the
explicitly allowed member names.  Equality of compressed bytes is evidence of
the resulting archive bytes; it does not identify the implementation path
that produced them.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import io
import json
from pathlib import Path
import struct
import sys
import zipfile
from typing import Any, Iterable, Sequence


SCHEMA = "litchi.performance.zip-census.v1"
EXPECTED_MEMBER_COUNT = 445
EXPECTED_SLIDE_POSITIONS = (0, 28, 57, 85, 114, 142, 171, 199)
EXPECTED_SLIDE_NAMES = tuple(
    f"ppt/slides/slide{position + 1}.xml" for position in EXPECTED_SLIDE_POSITIONS
)

EOCD = b"PK\x05\x06"
ZIP64_EOCD = b"PK\x06\x06"
ZIP64_LOCATOR = b"PK\x06\x07"
CENTRAL = b"PK\x01\x02"
LOCAL = b"PK\x03\x04"
DESCRIPTOR = b"PK\x07\x08"

LOCAL_STRUCT = struct.Struct("<4s5H3I2H")
CENTRAL_STRUCT = struct.Struct("<4s6H3I5H2I")
EOCD_STRUCT = struct.Struct("<4s4H2IH")
ZIP64_LOCATOR_STRUCT = struct.Struct("<4sIQI")
ZIP64_EOCD_FIXED = struct.Struct("<4sQ2H2I2Q")

COMPONENTS = (
    "local_header_metadata",
    "local_name",
    "local_extra",
    "local_data_descriptor",
    "raw_compressed_payload",
    "central_metadata",
    "central_name",
    "central_extra",
    "central_comment",
    "decoded_payload",
)


class CensusError(ValueError):
    """The archive or the comparison is outside the helper's contract."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CensusError(message)


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def decode_name(raw: bytes, flags: int) -> str:
    codec = "utf-8" if flags & 0x800 else "cp437"
    try:
        return raw.decode(codec)
    except UnicodeDecodeError as error:
        raise CensusError(f"invalid {codec} member name: {raw!r}") from error


def hex_digest(value: bytes) -> dict[str, Any]:
    return {"bytes": len(value), "sha256": sha256_bytes(value)}


def find_eocd(data: bytes) -> tuple[int, tuple[Any, ...]]:
    # The EOCD comment is at most 65,535 bytes.  Requiring its declared end to
    # be EOF avoids mistaking a byte sequence in a comment for the record.
    first = max(0, len(data) - 0xFFFF - EOCD_STRUCT.size)
    for position in range(len(data) - EOCD_STRUCT.size, first - 1, -1):
        if data[position : position + 4] != EOCD:
            continue
        values = EOCD_STRUCT.unpack_from(data, position)
        comment_length = values[-1]
        if position + EOCD_STRUCT.size + comment_length == len(data):
            return position, values
    raise CensusError("cannot locate a terminal EOCD record")


def parse_zip64_extra(
    extra: bytes,
    *,
    need_uncompressed: bool,
    need_compressed: bool,
    need_offset: bool,
    need_disk: bool,
) -> tuple[int | None, int | None, int | None, int | None]:
    """Read the ordered ZIP64 extra fields required by one central record."""

    cursor = 0
    payload: bytes | None = None
    while cursor + 4 <= len(extra):
        field_id, field_length = struct.unpack_from("<HH", extra, cursor)
        end = cursor + 4 + field_length
        require(end <= len(extra), "truncated extra field")
        if field_id == 0x0001:
            payload = extra[cursor + 4 : end]
            break
        cursor = end
    required = need_uncompressed or need_compressed or need_offset or need_disk
    require(not required or payload is not None, "record needs a ZIP64 extra field")
    if payload is None:
        return None, None, None, None

    cursor = 0

    def take(width: int, label: str) -> int:
        nonlocal cursor
        require(cursor + width <= len(payload), f"truncated ZIP64 {label}")
        if width == 8:
            value = struct.unpack_from("<Q", payload, cursor)[0]
        else:
            value = struct.unpack_from("<I", payload, cursor)[0]
        cursor += width
        return value

    uncompressed = take(8, "uncompressed size") if need_uncompressed else None
    compressed = take(8, "compressed size") if need_compressed else None
    offset = take(8, "local-header offset") if need_offset else None
    disk = take(4, "disk number") if need_disk else None
    return uncompressed, compressed, offset, disk


def central_location(data: bytes, eocd_position: int, eocd: tuple[Any, ...]) -> tuple[int, int, int, bool]:
    _, disk, central_disk, entries_disk, entries, central_size, central_offset, _ = eocd
    require(disk == 0 and central_disk == 0, "multi-disk ZIP archives are unsupported")
    zip64 = (
        entries_disk == 0xFFFF
        or entries == 0xFFFF
        or central_size == 0xFFFFFFFF
        or central_offset == 0xFFFFFFFF
    )
    if not zip64:
        require(entries_disk == entries, "EOCD entry counts disagree")
        return central_offset, central_size, entries, False

    locator_position = eocd_position - ZIP64_LOCATOR_STRUCT.size
    require(locator_position >= 0, "ZIP64 EOCD locator is missing")
    locator = ZIP64_LOCATOR_STRUCT.unpack_from(data, locator_position)
    require(locator[0] == ZIP64_LOCATOR, "invalid ZIP64 EOCD locator")
    _, zip64_disk, zip64_offset, total_disks = locator
    require(zip64_disk == 0 and total_disks == 1, "multi-disk ZIP64 archives are unsupported")
    require(zip64_offset + ZIP64_EOCD_FIXED.size <= len(data), "truncated ZIP64 EOCD")
    fixed = ZIP64_EOCD_FIXED.unpack_from(data, zip64_offset)
    require(fixed[0] == ZIP64_EOCD, "invalid ZIP64 EOCD signature")
    record_size = fixed[1]
    require(zip64_offset + 12 + record_size <= len(data), "truncated ZIP64 EOCD record")
    _, _, _, _, disk_number, central_disk_number, entries_disk64, entries64, size64, offset64 = fixed
    require(disk_number == 0 and central_disk_number == 0, "multi-disk ZIP64 records are unsupported")
    require(entries_disk64 == entries64, "ZIP64 entry counts disagree")
    return offset64, size64, entries64, True


@dataclass(frozen=True)
class CentralRecord:
    index: int
    name: str
    name_raw: bytes
    flags: int
    method: int
    crc32: int
    compressed_size: int
    uncompressed_size: int
    header_offset: int
    fixed: bytes
    metadata: bytes
    extra: bytes
    comment: bytes
    zip64: bool


@dataclass(frozen=True)
class LocalRecord:
    fixed: bytes
    metadata: bytes
    name_raw: bytes
    extra: bytes
    compressed: bytes
    descriptor: bytes


@dataclass(frozen=True)
class MemberObservation:
    central: CentralRecord
    local: LocalRecord
    decoded_hash: str
    decoded_size: int


@dataclass(frozen=True)
class ParsedArchive:
    path: Path
    data: bytes
    comment: bytes
    members: tuple[MemberObservation, ...]
    zip64: bool
    has_data_descriptors: bool


def parse_central_records(data: bytes, start: int, size: int, count: int) -> tuple[CentralRecord, ...]:
    require(start >= 0 and size >= 0 and start + size <= len(data), "central directory is out of bounds")
    records: list[CentralRecord] = []
    cursor = start
    end = start + size
    for index in range(count):
        require(cursor + CENTRAL_STRUCT.size <= end, f"central record {index} is truncated")
        values = CENTRAL_STRUCT.unpack_from(data, cursor)
        (
            signature,
            version_made,
            version_needed,
            flags,
            method,
            mtime,
            mdate,
            crc32,
            compressed32,
            uncompressed32,
            name_length,
            extra_length,
            comment_length,
            disk,
            internal_attributes,
            external_attributes,
            offset32,
        ) = values
        require(signature == CENTRAL, f"central record {index} has an invalid signature")
        record_end = cursor + CENTRAL_STRUCT.size + name_length + extra_length + comment_length
        require(record_end <= end, f"central record {index} extends past the directory")
        name_start = cursor + CENTRAL_STRUCT.size
        name_raw = data[name_start : name_start + name_length]
        extra_start = name_start + name_length
        extra = data[extra_start : extra_start + extra_length]
        comment_start = extra_start + extra_length
        comment = data[comment_start : comment_start + comment_length]
        zip64_needed = (
            compressed32 == 0xFFFFFFFF
            or uncompressed32 == 0xFFFFFFFF
            or offset32 == 0xFFFFFFFF
            or disk == 0xFFFF
        )
        zip64_uncompressed, zip64_compressed, zip64_offset, zip64_disk = parse_zip64_extra(
            extra,
            need_uncompressed=uncompressed32 == 0xFFFFFFFF,
            need_compressed=compressed32 == 0xFFFFFFFF,
            need_offset=offset32 == 0xFFFFFFFF,
            need_disk=disk == 0xFFFF,
        )
        uncompressed = zip64_uncompressed if uncompressed32 == 0xFFFFFFFF else uncompressed32
        compressed = zip64_compressed if compressed32 == 0xFFFFFFFF else compressed32
        header_offset = zip64_offset if offset32 == 0xFFFFFFFF else offset32
        disk_number = zip64_disk if disk == 0xFFFF else disk
        require(uncompressed is not None and compressed is not None and header_offset is not None, "incomplete ZIP64 central record")
        require(disk_number == 0, "members on a non-zero disk are unsupported")
        fixed = data[cursor : cursor + CENTRAL_STRUCT.size]
        # The final four bytes are the physical local-header offset.  It is
        # expected to move when preceding members change, so it is the sole
        # central fixed-field excluded from central metadata equality.
        metadata = fixed[:-4]
        records.append(
            CentralRecord(
                index=index,
                name=decode_name(name_raw, flags),
                name_raw=name_raw,
                flags=flags,
                method=method,
                crc32=crc32,
                compressed_size=int(compressed),
                uncompressed_size=int(uncompressed),
                header_offset=int(header_offset),
                fixed=fixed,
                metadata=metadata,
                extra=extra,
                comment=comment,
                zip64=zip64_needed,
            )
        )
        cursor = record_end
    require(cursor == end, "central directory has trailing or unaccounted bytes")
    return tuple(records)


def descriptor_bytes(data: bytes, offset: int, central: CentralRecord, local_fixed: bytes) -> bytes:
    """Return the exact descriptor, accepting signed and unsigned forms."""

    zip64_hint = (
        central.zip64
        or struct.unpack_from("<I", local_fixed, 14)[0] == 0xFFFFFFFF
        or struct.unpack_from("<I", local_fixed, 18)[0] == 0xFFFFFFFF
    )
    candidates: list[tuple[bool, int, bytes]] = []
    for has_signature in (True, False):
        beginning = offset
        if has_signature:
            if data[beginning : beginning + 4] != DESCRIPTOR:
                continue
            beginning += 4
        for width in ((8 if zip64_hint else 4), (4 if zip64_hint else 8)):
            length = 4 + width * 2
            end = beginning + length
            if end > len(data):
                continue
            crc32 = struct.unpack_from("<I", data, beginning)[0]
            if width == 4:
                compressed, uncompressed = struct.unpack_from("<2I", data, beginning + 4)
            else:
                compressed, uncompressed = struct.unpack_from("<2Q", data, beginning + 4)
            if (crc32, compressed, uncompressed) != (
                central.crc32,
                central.compressed_size,
                central.uncompressed_size,
            ):
                continue
            candidates.append((has_signature, width, data[offset:end]))
    require(candidates, f"data descriptor does not match central record {central.index}")
    # A normal 32-bit descriptor is preferred unless the record requires
    # ZIP64.  If malformed bytes make two candidates possible, fail closed.
    preferred = [candidate for candidate in candidates if candidate[1] == (8 if zip64_hint else 4)]
    if len(preferred) == 1:
        return preferred[0][2]
    require(len(candidates) == 1, f"ambiguous data descriptor for central record {central.index}")
    return candidates[0][2]


def parse_local(data: bytes, central: CentralRecord) -> LocalRecord:
    offset = central.header_offset
    require(offset + LOCAL_STRUCT.size <= len(data), f"local header {central.index} is truncated")
    values = LOCAL_STRUCT.unpack_from(data, offset)
    signature, version, flags, method, mtime, mdate, crc32, compressed32, uncompressed32, name_length, extra_length = values
    require(signature == LOCAL, f"local header {central.index} has an invalid signature")
    header_end = offset + LOCAL_STRUCT.size + name_length + extra_length
    require(header_end <= len(data), f"local header {central.index} name/extra is truncated")
    name_start = offset + LOCAL_STRUCT.size
    name_raw = data[name_start : name_start + name_length]
    extra = data[name_start + name_length : header_end]
    compressed_end = header_end + central.compressed_size
    require(compressed_end <= len(data), f"compressed payload {central.index} is truncated")
    descriptor = b""
    if flags & 0x08:
        descriptor = descriptor_bytes(data, compressed_end, central, data[offset:offset + LOCAL_STRUCT.size])
    return LocalRecord(
        fixed=data[offset : offset + LOCAL_STRUCT.size],
        metadata=data[offset + 4 : offset + LOCAL_STRUCT.size],
        name_raw=name_raw,
        extra=extra,
        compressed=data[header_end:compressed_end],
        descriptor=descriptor,
    )


def decoded_hashes(data: bytes, records: Sequence[CentralRecord]) -> tuple[tuple[str, int], ...]:
    result: list[tuple[str, int]] = []
    try:
        archive = zipfile.ZipFile(io.BytesIO(data), "r")
    except (OSError, zipfile.BadZipFile) as error:
        raise CensusError(f"zipfile cannot open archive: {error}") from error
    with archive:
        infos = archive.infolist()
        require(len(infos) == len(records), "zipfile and raw central-directory counts disagree")
        for index, (info, record) in enumerate(zip(infos, records)):
            require(info.filename == record.name, f"decoded name mismatch at member {index}")
            require(info.flag_bits == record.flags, f"flag mismatch at member {index}")
            require(info.compress_type == record.method, f"method mismatch at member {index}")
            require(info.CRC == record.crc32, f"CRC mismatch at member {index}")
            require(info.compress_size == record.compressed_size, f"compressed-size mismatch at member {index}")
            require(info.file_size == record.uncompressed_size, f"uncompressed-size mismatch at member {index}")
            digest = hashlib.sha256()
            size = 0
            try:
                with archive.open(info, "r") as stream:
                    for block in iter(lambda: stream.read(1 << 20), b""):
                        digest.update(block)
                        size += len(block)
            except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
                raise CensusError(f"cannot decode member {index} ({record.name}): {error}") from error
            require(size == record.uncompressed_size, f"decoded byte count mismatch at member {index}")
            result.append((digest.hexdigest(), size))
    return tuple(result)


def parse_archive(path: Path) -> ParsedArchive:
    require(path.is_file() and not path.is_symlink(), f"archive is not a regular file: {path}")
    try:
        data = path.read_bytes()
    except OSError as error:
        raise CensusError(f"cannot read archive {path}: {error}") from error
    eocd_position, eocd = find_eocd(data)
    comment = data[eocd_position + EOCD_STRUCT.size :]
    central_start, central_size, count, zip64 = central_location(data, eocd_position, eocd)
    records = parse_central_records(data, central_start, central_size, count)
    decoded = decoded_hashes(data, records)
    members: list[MemberObservation] = []
    for record, (decoded_hash, decoded_size) in zip(records, decoded):
        local = parse_local(data, record)
        members.append(MemberObservation(record, local, decoded_hash, decoded_size))
    return ParsedArchive(
        path=path,
        data=data,
        comment=comment,
        members=tuple(members),
        zip64=zip64 or any(member.central.zip64 for member in members),
        has_data_descriptors=any(bool(member.local.descriptor) for member in members),
    )


def side(member: MemberObservation) -> dict[str, Any]:
    central = member.central
    local = member.local
    return {
        "index": central.index,
        "name": central.name,
        "name_raw_hex": central.name_raw.hex(),
        "local_header": {"raw_fixed": hex_digest(local.fixed), "metadata": hex_digest(local.metadata)},
        "local_name": hex_digest(local.name_raw),
        "local_extra": hex_digest(local.extra),
        "local_data_descriptor": hex_digest(local.descriptor),
        "raw_compressed_payload": {
            **hex_digest(local.compressed),
            "crc32": f"{central.crc32:08x}",
            "method": central.method,
        },
        "central_metadata": hex_digest(central.metadata),
        "central_name": hex_digest(central.name_raw),
        "central_extra": hex_digest(central.extra),
        "central_comment": hex_digest(central.comment),
        "decoded_payload": {"bytes": member.decoded_size, "sha256": member.decoded_hash},
        "physical": {
            "local_header_offset": central.header_offset,
            "compressed_data_offset": central.header_offset + len(local.fixed) + len(local.name_raw) + len(local.extra),
            "compressed_size": central.compressed_size,
        },
    }


def compare_component(source: bytes, output: bytes) -> dict[str, Any]:
    return {
        "equal": source == output,
        "source": hex_digest(source),
        "output": hex_digest(output),
    }


def compare_decoded(source: MemberObservation, output: MemberObservation) -> dict[str, Any]:
    source_value = {"bytes": source.decoded_size, "sha256": source.decoded_hash}
    output_value = {"bytes": output.decoded_size, "sha256": output.decoded_hash}
    return {"equal": source_value == output_value, "source": source_value, "output": output_value}


def compare_member(source: MemberObservation, output: MemberObservation) -> dict[str, Any]:
    source_c = source.central
    output_c = output.central
    source_l = source.local
    output_l = output.local
    components = {
        "local_header_metadata": compare_component(source_l.metadata, output_l.metadata),
        "local_name": compare_component(source_l.name_raw, output_l.name_raw),
        "local_extra": compare_component(source_l.extra, output_l.extra),
        "local_data_descriptor": compare_component(source_l.descriptor, output_l.descriptor),
        "raw_compressed_payload": compare_component(source_l.compressed, output_l.compressed),
        "central_metadata": compare_component(source_c.metadata, output_c.metadata),
        "central_name": compare_component(source_c.name_raw, output_c.name_raw),
        "central_extra": compare_component(source_c.extra, output_c.extra),
        "central_comment": compare_component(source_c.comment, output_c.comment),
        "decoded_payload": compare_decoded(source, output),
    }
    changed = [name for name, result in components.items() if not result["equal"]]
    return {
        "index": source_c.index,
        "name": source_c.name,
        "output_name": output_c.name,
        "allowed_change": False,
        "changed_components": changed,
        "components": components,
        "source": side(source),
        "output": side(output),
        "central_relocation_offset": {
            "ignored_for_equality": True,
            "source": source_c.header_offset,
            "output": output_c.header_offset,
        },
    }


def fixture_expectation(enabled: bool, source: ParsedArchive, output: ParsedArchive) -> dict[str, Any]:
    result: dict[str, Any] = {
        "mode": "expected_generated_pptx_slides" if enabled else "explicit_allowlist",
        "generated_fixture_scope": {
            "expected_source_member_count": EXPECTED_MEMBER_COUNT,
            "expected_selected_slide_positions": list(EXPECTED_SLIDE_POSITIONS),
            "expected_selected_slide_names": list(EXPECTED_SLIDE_NAMES),
            "source_member_count": len(source.members),
            "output_member_count": len(output.members),
        },
        "adversarial_zip64_data_descriptor_scope": {
            "zip64_parser_observed_source": source.zip64,
            "zip64_parser_observed_output": output.zip64,
            "data_descriptor_parser_observed_source": source.has_data_descriptors,
            "data_descriptor_parser_observed_output": output.has_data_descriptors,
            "claim": "parser support is recorded; generated-fixture equality is not adversarial coverage",
        },
    }
    if enabled:
        result["generated_fixture_scope"]["member_count_matches_expected"] = (
            len(source.members) == EXPECTED_MEMBER_COUNT and len(output.members) == EXPECTED_MEMBER_COUNT
        )
    return result


def build_report(
    source: ParsedArchive,
    output: ParsedArchive,
    allowed: set[str],
    expected_generated_slides: bool,
) -> dict[str, Any]:
    source_names = [member.central.name for member in source.members]
    output_names = [member.central.name for member in output.members]
    ordered_equal = source_names == output_names
    archive_comment = compare_component(source.comment, output.comment)
    archive_comment["source"]["raw_hex"] = source.comment.hex()
    archive_comment["output"]["raw_hex"] = output.comment.hex()
    reports: list[dict[str, Any]] = []
    violations: list[dict[str, Any]] = []
    changed_by_component = {component: [] for component in COMPONENTS}
    changed_members: list[str] = []
    effective_allowed = set(allowed)
    if expected_generated_slides:
        effective_allowed.update(EXPECTED_SLIDE_NAMES)

    for index in range(min(len(source.members), len(output.members))):
        member = compare_member(source.members[index], output.members[index])
        name = member["name"]
        member["allowed_change"] = name in effective_allowed
        reports.append(member)
        changed = member["changed_components"]
        if changed:
            changed_members.append(name)
        for component in changed:
            changed_by_component[component].append(name)
            if name not in effective_allowed:
                violations.append({"member": name, "component": component, "reason": "changed_untouched_member"})
        if member["output_name"] != name:
            violations.append({"index": index, "reason": "ordered_filename_changed", "source": name, "output": member["output_name"]})

    if len(source.members) != len(output.members):
        violations.append(
            {
                "reason": "member_count_changed",
                "source": len(source.members),
                "output": len(output.members),
            }
        )
    if not ordered_equal:
        violations.append({"reason": "ordered_filenames_changed"})
    if not archive_comment["equal"]:
        violations.append({"reason": "archive_comment_changed"})
    if expected_generated_slides:
        source_set = set(source_names)
        output_set = set(output_names)
        for position, name in zip(EXPECTED_SLIDE_POSITIONS, EXPECTED_SLIDE_NAMES):
            if name not in source_set or name not in output_set:
                violations.append({"reason": "expected_slide_member_missing", "member": name})
            if position >= len(source_names) or source_names[position] != name:
                violations.append(
                    {
                        "reason": "expected_slide_position_mismatch",
                        "side": "source",
                        "position": position,
                        "expected": name,
                        "actual": source_names[position] if position < len(source_names) else None,
                    }
                )
            if position >= len(output_names) or output_names[position] != name:
                violations.append(
                    {
                        "reason": "expected_slide_position_mismatch",
                        "side": "output",
                        "position": position,
                        "expected": name,
                        "actual": output_names[position] if position < len(output_names) else None,
                    }
                )
        if len(source.members) != EXPECTED_MEMBER_COUNT or len(output.members) != EXPECTED_MEMBER_COUNT:
            violations.append({"reason": "generated_fixture_member_count_mismatch", "expected": EXPECTED_MEMBER_COUNT})

    return {
        "schema": SCHEMA,
        "status": "pass" if not violations else "fail",
        "inputs": {
            "source": {"path": str(source.path), "bytes": len(source.data), "sha256": sha256_bytes(source.data)},
            "output": {"path": str(output.path), "bytes": len(output.data), "sha256": sha256_bytes(output.data)},
        },
        "archive": {
            "source": {"comment": hex_digest(source.comment), "member_count": len(source.members), "zip64": source.zip64, "data_descriptors": source.has_data_descriptors},
            "output": {"comment": hex_digest(output.comment), "member_count": len(output.members), "zip64": output.zip64, "data_descriptors": output.has_data_descriptors},
            "comment_comparison": archive_comment,
        },
        "allow_changed": sorted(allowed),
        "effective_allow_changed": sorted(effective_allowed),
        "ordered_filenames": {
            "equal": ordered_equal,
            "source": source_names,
            "output": output_names,
        },
        "changed_members": {"names": changed_members, "by_component": changed_by_component},
        "unchanged_members": [member["name"] for member in reports if not member["changed_components"]],
        "violations": violations,
        "members": reports,
        "coverage": fixture_expectation(expected_generated_slides, source, output),
        "claims": {
            "unchanged_compressed_bytes": "exact raw compressed payload byte equality is reported per member",
            "implementation_path": "not inferred from byte equality",
            "central_relocation_offset": "ignored only for central metadata equality because local-header offsets relocate when archive layout changes",
            "decoded_hashes": "computed independently by stdlib zipfile readback",
        },
    }


def write_create_only(path: Path, value: dict[str, Any]) -> None:
    require(not path.exists(), f"refusing to overwrite existing report: {path}")
    require(path.parent.is_dir(), f"report parent does not exist: {path.parent}")
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    try:
        with path.open("x", encoding="utf-8") as stream:
            stream.write(encoded)
    except FileExistsError as error:
        raise CensusError(f"refusing to overwrite existing report: {path}") from error
    except OSError as error:
        raise CensusError(f"cannot create report {path}: {error}") from error


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--source", type=Path, required=True, help="source ZIP archive")
    result.add_argument("--output", type=Path, required=True, help="candidate/output ZIP archive")
    result.add_argument("--report", type=Path, required=True, help="new JSON report path")
    result.add_argument(
        "--allow-changed",
        action="append",
        default=[],
        metavar="MEMBER",
        help="member name whose component differences are allowed; repeat as needed",
    )
    result.add_argument(
        "--expected-generated-slides",
        action="store_true",
        help="also allow the eight selected generated fixture slide XML members and enforce the 445-member fixture shape",
    )
    return result


def main(argv: Sequence[str] | None = None) -> int:
    arguments = parser().parse_args(argv)
    report_path = arguments.report
    if report_path.exists():
        print(f"error: refusing to overwrite existing report: {report_path}", file=sys.stderr)
        return 2
    try:
        source = parse_archive(arguments.source)
        output = parse_archive(arguments.output)
        report = build_report(source, output, set(arguments.allow_changed), arguments.expected_generated_slides)
        write_create_only(report_path, report)
    except (CensusError, OSError) as error:
        # A malformed archive has no trustworthy member comparison to emit.
        # If the caller supplied a fresh report path, retain a small failure
        # receipt so an automated qualification cannot mistake no output for a
        # successful run.
        failure = {"schema": SCHEMA, "status": "error", "error": str(error)}
        try:
            if not report_path.exists():
                write_create_only(report_path, failure)
        except (CensusError, OSError):
            pass
        print(f"error: {error}", file=sys.stderr)
        return 2
    if report["status"] != "pass":
        print(f"census failed: {len(report['violations'])} violation(s)", file=sys.stderr)
        return 1
    print(f"census passed: {len(report['members'])} members; report={report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
