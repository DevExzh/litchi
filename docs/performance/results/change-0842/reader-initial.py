"""Offline, fail-closed readers for the 0842 compression dispatch packet.

The capture driver and the Rust probe own process creation and measurement.  This
module only reads retained JSON/ZIP evidence.  In particular, the ZIP checker
below does not trust a digest emitted by the probe: it opens the retained bytes
with :mod:`zipfile`, recomputes every logical member digest and CRC, and also
extracts the raw compressed byte range from each local file header.
"""

from __future__ import annotations

import binascii
import hashlib
import io
import json
import math
import struct
import statistics
import zipfile
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


PACKET = Path(__file__).resolve().parent
REPORT_SCHEMA = "litchi.compression-probe-0842.v1"
REPORT_TOOL = "compression-probe-0842"
GENERATOR = "probe/main.rs::payload_for-0842.v1"
SHAPES = ("tiny", "xml", "random", "repeated")
MODES = ("write", "read", "changed")
LEGS = ("before", "after")
NATIVE_SAMPLES = 30
NATIVE_WARMUP = 3
QUALIFICATION_SAMPLES = 1
QUALIFICATION_WARMUP = 0
NATIVE_BLOCKS = 6
NATIVE_REPORTS = NATIVE_BLOCKS * len(SHAPES) * len(MODES) * len(LEGS)
QUALIFICATION_REPORTS = len(SHAPES) * len(MODES) * len(LEGS)
MAX_SAMPLES = 100_000
MAX_WARMUP = 100_000
HEX = frozenset("0123456789abcdefABCDEF")

PARTS = (
    {
        "name": "/word/document.xml",
        "archive_name": "word/document.xml",
        "content_type": "application/xml",
        "role": "DOCX-like XML part",
    },
    {
        "name": "/xl/workbook.xml",
        "archive_name": "xl/workbook.xml",
        "content_type": "application/xml",
        "role": "XLSX-like XML part",
    },
    {
        "name": "/ppt/presentation.xml",
        "archive_name": "ppt/presentation.xml",
        "content_type": "application/xml",
        "role": "PPTX-like XML part",
    },
    {
        "name": "/custom/payload.bin",
        "archive_name": "custom/payload.bin",
        "content_type": "application/octet-stream",
        "role": "opaque payload part",
    },
)
PART_ARCHIVE_NAMES = tuple(item["archive_name"] for item in PARTS)
METADATA_ARCHIVE_NAMES = ("[Content_Types].xml", "_rels/.rels")


class ReaderError(ValueError):
    """Raised when retained evidence is missing, malformed, or contradictory."""


def fail(message: str) -> None:
    raise ReaderError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _json(path: Path) -> Any:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing JSON report: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON report {path}: {error}")


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> str:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    try:
        with path.open("rb") as stream:
            digest = hashlib.sha256()
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
            return digest.hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")


def _sha(value: Any, label: str) -> str:
    require(
        isinstance(value, str) and len(value) == 64 and all(char in HEX for char in value),
        f"{label} is not a SHA-256 hexadecimal string",
    )
    return value.lower()


def _uint(value: Any, label: str, *, positive: bool = False) -> int:
    require(
        isinstance(value, int) and not isinstance(value, bool) and value >= 0,
        f"{label} is not a non-negative integer",
    )
    if positive:
        require(value > 0, f"{label} is not positive")
    return value


def _finite(value: Any, label: str) -> float:
    require(
        isinstance(value, (int, float)) and not isinstance(value, bool),
        f"{label} is not numeric",
    )
    result = float(value)
    require(math.isfinite(result), f"{label} is not finite")
    return result


def _bytes(value: bytes | bytearray | memoryview | Path) -> bytes:
    if isinstance(value, Path):
        require(value.is_file() and not value.is_symlink(), f"missing archive: {value}")
        try:
            return value.read_bytes()
        except OSError as error:
            fail(f"cannot read archive {value}: {error}")
    require(isinstance(value, (bytes, bytearray, memoryview)), "archive is not bytes")
    return bytes(value)


def _safe_member_name(name: Any, label: str) -> str:
    require(isinstance(name, str) and name and "\x00" not in name, f"{label}: invalid name")
    path = Path(name)
    require(not path.is_absolute() and ".." not in path.parts, f"{label}: unsafe name")
    require("\\" not in name, f"{label}: backslash in name")
    return name


def _raw_member(data: bytes, info: zipfile.ZipInfo, label: str) -> tuple[bytes, dict[str, Any]]:
    """Extract and independently bind one entry's raw compressed byte range."""

    offset = info.header_offset
    require(offset >= 0 and offset + 30 <= len(data), f"{label}: local header is out of range")
    fields = struct.unpack_from("<4s5H3I2H", data, offset)
    signature, version, flags, method, mtime, mdate, crc, compressed, logical, name_len, extra_len = fields
    require(signature == b"PK\x03\x04", f"{label}: local header signature changed")
    require(method == info.compress_type, f"{label}: local/central compression method differs")
    require(flags == info.flag_bits, f"{label}: local/central flags differ")
    name_start = offset + 30
    name_end = name_start + name_len
    extra_end = name_end + extra_len
    require(extra_end <= len(data), f"{label}: local header fields are out of range")
    name_bytes = data[name_start:name_end]
    try:
        decoded_name = name_bytes.decode("utf-8" if flags & 0x800 else "cp437")
    except UnicodeDecodeError as error:
        fail(f"{label}: local name is not decodable: {error}")
    require(decoded_name == info.filename, f"{label}: local/central names differ")
    # A data descriptor legitimately leaves these fields zero in the local
    # header.  Otherwise bind them to the central directory values exactly.
    if not flags & 0x08:
        require(crc == info.CRC, f"{label}: local/central CRC differs")
        if compressed != 0xFFFFFFFF:
            require(compressed == info.compress_size, f"{label}: local compressed size differs")
        if logical != 0xFFFFFFFF:
            require(logical == info.file_size, f"{label}: local logical size differs")
    start = extra_end
    end = start + info.compress_size
    require(start <= end <= len(data), f"{label}: compressed byte range is out of bounds")
    return data[start:end], {
        "offset": offset,
        "version": version,
        "flags": flags,
        "method": method,
        "mtime": mtime,
        "mdate": mdate,
        "name_bytes": name_len,
        "extra_bytes": extra_len,
    }


def _compression_name(method: int) -> str:
    return {
        zipfile.ZIP_STORED: "stored",
        zipfile.ZIP_DEFLATED: "deflated",
        zipfile.ZIP_BZIP2: "bzip2",
        zipfile.ZIP_LZMA: "lzma",
    }.get(method, f"method-{method}")


def archive_manifest(value: bytes | bytearray | memoryview | Path) -> dict[str, Any]:
    """Return a strict independent manifest for one retained ZIP archive.

    The ``logical`` digest is over decompressed member bytes.  The ``raw``
    digest is over precisely the compressed data between the local header and
    the central directory.  Central-directory fields and local-header fields
    are retained so a later comparison cannot silently collapse metadata or
    framing differences into a logical-content-only result.
    """

    data = _bytes(value)
    require(data, "archive is empty")
    try:
        archive = zipfile.ZipFile(io.BytesIO(data), "r")
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        fail(f"invalid ZIP archive: {error}")
    try:
        infos = archive.infolist()
        require(infos, "archive has no members")
        members: list[dict[str, Any]] = []
        names: set[str] = set()
        for index, info in enumerate(infos):
            label = f"archive.members[{index}]"
            name = _safe_member_name(info.filename, f"{label}.name")
            require(name not in names, f"{label}: duplicate member name {name!r}")
            names.add(name)
            require(not (info.flag_bits & 0x1), f"{label}: encrypted members are unsupported")
            try:
                logical = archive.read(info)
            except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
                fail(f"{label}: cannot decompress member: {error}")
            require(len(logical) == info.file_size, f"{label}: logical size differs")
            crc = binascii.crc32(logical) & 0xFFFFFFFF
            require(crc == info.CRC, f"{label}: CRC32 differs from central directory")
            raw, local = _raw_member(data, info, label)
            central = {
                "compress_type": info.compress_type,
                "flag_bits": info.flag_bits,
                "create_system": info.create_system,
                "create_version": info.create_version,
                "extract_version": info.extract_version,
                "reserved": info.reserved,
                "volume": info.volume,
                "internal_attr": info.internal_attr,
                "external_attr": info.external_attr,
                "date_time": list(info.date_time),
                "comment": info.comment.hex(),
                "extra": info.extra.hex(),
            }
            members.append({
                "index": index,
                "name": name,
                "compression": _compression_name(info.compress_type),
                "central": central,
                "local": local,
                "crc32": crc,
                "crc32_hex": f"{crc:08x}",
                "logical": {"bytes": len(logical), "sha256": _sha256(logical)},
                "raw": {"bytes": len(raw), "sha256": _sha256(raw)},
            })
    finally:
        archive.close()
    return {
        "archive": {"bytes": len(data), "sha256": _sha256(data)},
        "member_count": len(members),
        "members": members,
    }


def _member_projection(member: Mapping[str, Any], label: str) -> dict[str, Any]:
    """Normalize a report member identity for exact archive comparison."""

    require(isinstance(member, Mapping), f"{label}: member is not an object")
    name = _safe_member_name(member.get("name"), f"{label}.name")
    logical = member.get("logical")
    raw = member.get("raw")
    require(isinstance(logical, Mapping), f"{label}.logical is missing")
    require(isinstance(raw, Mapping), f"{label}.raw is missing")
    logical_bytes = _uint(logical.get("bytes"), f"{label}.logical.bytes")
    raw_bytes = _uint(raw.get("bytes"), f"{label}.raw.bytes")
    logical_sha = _sha(logical.get("sha256"), f"{label}.logical.sha256")
    raw_sha = _sha(raw.get("sha256"), f"{label}.raw.sha256")
    crc = _uint(member.get("crc32"), f"{label}.crc32")
    require(crc <= 0xFFFFFFFF, f"{label}.crc32 exceeds u32")
    compression = member.get("compression")
    require(isinstance(compression, str) and compression, f"{label}.compression is missing")
    central = member.get("central")
    local = member.get("local")
    require(isinstance(central, Mapping), f"{label}.central is missing")
    require(isinstance(local, Mapping), f"{label}.local is missing")
    return {
        "name": name,
        "compression": compression,
        "central": dict(central),
        "local": dict(local),
        "crc32": crc,
        "crc32_hex": f"{crc:08x}",
        "logical": {"bytes": logical_bytes, "sha256": logical_sha},
        "raw": {"bytes": raw_bytes, "sha256": raw_sha},
    }


def compare_archive_manifest(actual: Mapping[str, Any], expected: Mapping[str, Any],
                             label: str = "archive") -> dict[str, Any]:
    """Compare every member's logical/raw identity and all retained metadata."""

    require(isinstance(actual, Mapping), f"{label}: actual manifest is not an object")
    require(isinstance(expected, Mapping), f"{label}: expected manifest is not an object")
    actual_members = actual.get("members")
    expected_members = expected.get("members")
    require(isinstance(actual_members, list) and isinstance(expected_members, list),
            f"{label}: member lists are missing")
    require(len(actual_members) == len(expected_members), f"{label}: member count differs")
    for index, (left, right) in enumerate(zip(actual_members, expected_members)):
        a = _member_projection(left, f"{label}.actual.members[{index}]")
        e = _member_projection(right, f"{label}.expected.members[{index}]")
        require(a == e, f"{label}.members[{index}] identity or metadata differs")
    if "member_count" in expected:
        require(expected.get("member_count") == len(expected_members),
                f"{label}: expected member_count is inconsistent")
        require(actual.get("member_count") == len(actual_members),
                f"{label}: actual member_count is inconsistent")
    return {"member_count": len(actual_members), "members": actual_members}


def _u32(value: int) -> int:
    return value & 0xFFFFFFFF


def _xorshift_payload(size: int, part_index: int, changed: bool) -> bytes:
    state = _u32(
        0x9E3779B9
        + part_index * 0x6D2B79F5
        + (0x13579BDF if changed else 0)
    )
    output = bytearray(size)
    for index in range(size):
        state = _u32(state ^ _u32(state << 13))
        state = _u32(state ^ (state >> 17))
        state = _u32(state ^ _u32(state << 5))
        output[index] = state >> 24
    return bytes(output)


def _repeated_payload(size: int, part_index: int, changed: bool) -> bytes:
    pattern = b"litchi-0842-repeated-deflate-pattern-0123456789abcdef"
    marker = 0x5A if changed else 0xA5
    return bytes(
        pattern[(offset + part_index * 7) % len(pattern)] ^ marker
        for offset in range(size)
    )


def _xml_payload(size: int, shape: str, part_index: int, changed: bool) -> bytes:
    shape_name = {
        "tiny": "tiny",
        "xml": "xml",
        "random": "random-control",
        "repeated": "repeated-control",
    }[shape]
    root = f'<root shape="{shape_name}" part="{part_index}" changed="{str(changed).lower()}">'
    tail_open = b"<tail>"
    tail_close = b"</tail>"
    close = b"</root>"
    fixed = len(root.encode("ascii")) + len(tail_open) + len(tail_close) + len(close)
    require(size >= fixed, "XML fixture size is too small for its framing")
    output = bytearray(root.encode("ascii"))
    item = 0
    while True:
        item_bytes = f'<item id="{item}">value-{shape_name}-{part_index}-{item}</item>'.encode("ascii")
        if len(output) + len(item_bytes) + fixed > size:
            break
        output.extend(item_bytes)
        item += 1
    output.extend(tail_open)
    filler = size - len(output) - len(tail_close) - len(close)
    require(filler >= 0, "XML fixture filler underflow")
    output.extend(bytes(ord("a") + ((offset + part_index) % 26) for offset in range(filler)))
    output.extend(tail_close)
    output.extend(close)
    require(len(output) == size, "XML fixture length differs")
    return bytes(output)


def payload_for(shape: str, part_index: int, changed: bool = False) -> bytes:
    """Rebuild the probe's deterministic payload without importing the probe."""

    require(shape in SHAPES, f"unknown fixture shape {shape!r}")
    require(0 <= part_index < len(PARTS), f"invalid fixture part index {part_index}")
    size = TINY_BYTES if shape == "tiny" else LARGE_BYTES
    if part_index < len(PARTS) - 1:
        return _xml_payload(size, shape, part_index, changed)
    if shape == "random":
        return _xorshift_payload(size, part_index, changed)
    return _repeated_payload(size, part_index, changed)


def _logical_identity_from_members(members: Sequence[tuple[str, bytes]]) -> dict[str, Any]:
    digest = hashlib.sha256()
    total = 0
    for name, payload in members:
        encoded_name = name.encode("utf-8")
        digest.update(len(encoded_name).to_bytes(8, "little"))
        digest.update(encoded_name)
        digest.update(len(payload).to_bytes(8, "little"))
        digest.update(payload)
        total += len(payload)
    return {"bytes": total, "sha256": digest.hexdigest()}


def _metadata_payloads() -> dict[str, bytes]:
    """Rebuild the authored OPC content-types member for this fixture."""

    content_types = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" '
        b'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b'<Override PartName="/custom/payload.bin" '
        b'ContentType="application/octet-stream"/>'
        b'</Types>'
    )
    return {
        "[Content_Types].xml": content_types,
        "_rels/.rels": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    }


def _generator_payload_check(manifest: Mapping[str, Any], data: bytes, shape: str,
                            changed: bool, label: str) -> dict[str, Any]:
    """Check every generated member's bytes, SHA-256, and CRC independently."""

    members = manifest.get("members")
    require(isinstance(members, list), f"{label}: ZIP member list is missing")
    by_name = {item.get("name"): item for item in members if isinstance(item, Mapping)}
    require(len(by_name) == len(members), f"{label}: ZIP member names are not unique")
    expected_names = (*METADATA_ARCHIVE_NAMES, *sorted(PART_ARCHIVE_NAMES))
    require(tuple(item.get("name") for item in members) == expected_names,
            f"{label}: ZIP member set differs from the four-part generator")
    for index, member in enumerate(members):
        central = member.get("central")
        local = member.get("local")
        require(isinstance(central, Mapping) and isinstance(local, Mapping),
                f"{label}.members[{index}]: ZIP metadata is missing")
        require(central.get("compress_type") == zipfile.ZIP_DEFLATED,
                f"{label}.members[{index}]: compression method is not Deflate")
        require(local.get("method") == zipfile.ZIP_DEFLATED,
                f"{label}.members[{index}]: local compression method is not Deflate")
    expected_payloads = _metadata_payloads()
    for index, name in enumerate(PART_ARCHIVE_NAMES):
        expected_payloads[name] = payload_for(shape, index, changed and index == len(PARTS) - 1)
    try:
        archive = zipfile.ZipFile(io.BytesIO(data), "r")
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        fail(f"{label}: cannot reopen ZIP for generator comparison: {error}")
    try:
      for name, expected in expected_payloads.items():
        member = by_name[name]
        logical = member.get("logical")
        require(isinstance(logical, Mapping), f"{label}.{name}: logical identity missing")
        require(logical.get("bytes") == len(expected), f"{label}.{name}: logical size differs")
        require(logical.get("sha256") == _sha256(expected),
                f"{label}.{name}: logical SHA differs from independent generator")
        try:
            actual = archive.read(name)
        except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
            fail(f"{label}.{name}: cannot read generated member: {error}")
        require(actual == expected, f"{label}.{name}: logical bytes differ from generator")
        crc = binascii.crc32(expected) & 0xFFFFFFFF
        require(member.get("crc32") == crc, f"{label}.{name}: CRC differs from generator")
    finally:
        archive.close()
    return {
        "generated_members": list(expected_names),
        "shape": shape,
        "changed": changed,
    }


def _archive_logical_identity(manifest: Mapping[str, Any], data: bytes,
                             label: str) -> dict[str, Any]:
    """Recompute the probe's ordered logical identity from Python ZIP reads."""

    try:
        archive = zipfile.ZipFile(io.BytesIO(data), "r")
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        fail(f"{label}: cannot reopen ZIP for logical identity: {error}")
    try:
        values = []
        for member in manifest["members"]:
            name = member["name"]
            values.append((name, archive.read(name)))
    finally:
        archive.close()
    return _logical_identity_from_members(values)


def _identity(value: Any, label: str, *, positive: bool = True) -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label} is not an identity object")
    require(set(value) == {"bytes", "sha256"}, f"{label} fields changed")
    return {
        "bytes": _uint(value.get("bytes"), f"{label}.bytes", positive=positive),
        "sha256": _sha(value.get("sha256"), f"{label}.sha256"),
    }


def _check_report_finite(value: Any, label: str) -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{label}: non-finite number")
    elif isinstance(value, list):
        for index, item in enumerate(value):
            _check_report_finite(item, f"{label}[{index}]")
    elif isinstance(value, dict):
        for key, item in value.items():
            require(isinstance(key, str), f"{label}: non-string key")
            _check_report_finite(item, f"{label}.{key}")


def _sample_elapsed(sample: Mapping[str, Any], label: str) -> int:
    value = sample.get("elapsed_ns")
    return _uint(value, f"{label}.elapsed_ns", positive=True)


def _nearest_rank(values: Sequence[float], quantile: float) -> float:
    require(values, "nearest-rank received no samples")
    ordered = sorted(_finite(value, "nearest-rank input") for value in values)
    rank = max(1, math.ceil(float(quantile) * len(ordered)))
    return ordered[min(rank, len(ordered)) - 1]


def _report_path(value: Path | Mapping[str, Any]) -> tuple[Any, str, Path | None]:
    if isinstance(value, Mapping):
        return value, "<report>", None
    path = Path(value)
    return _json(path), str(path), path


def validate_report(path: Path | Mapping[str, Any], shape: str, mode: str,
                    samples: int, warmup: int, artifact_path: Path | None = None,
                    expected: Mapping[str, Any] | None = None) -> dict[str, Any]:
    """Validate one report against the fixed 0842 schedule.

    ``expected`` is an optional independent qualification identity.  When
    supplied, the complete source/output/archive identity is compared against
    it; report-provided hashes are never accepted as their own oracle.
    """

    require(shape in SHAPES, f"unknown shape {shape!r}")
    require(mode in MODES, f"unknown mode {mode!r}")
    require(type(samples) is int and 0 < samples <= MAX_SAMPLES,
            "samples must be a positive bounded integer")
    require(type(warmup) is int and 0 <= warmup <= MAX_WARMUP,
            "warmup must be a bounded non-negative integer")
    report, label, report_file = _report_path(path)
    _check_report_finite(report, label)
    require(isinstance(report, Mapping), f"{label}: report is not an object")
    require(report.get("schema") == REPORT_SCHEMA, f"{label}: schema changed")
    require(report.get("tool") == REPORT_TOOL, f"{label}: tool changed")
    require(report.get("shape") == shape and report.get("mode") == mode,
            f"{label}: shape/mode changed")
    descriptions = {
        "tiny": "four 64-byte members",
        "xml": "four 1 MiB members with varying valid XML or textual payloads",
        "random": "three 1 MiB valid XML members and one xorshift payload",
        "repeated": "three 1 MiB valid XML members and one repeated deterministic byte pattern",
    }
    require(report.get("shape_description") == descriptions[shape],
            f"{label}: shape description changed")
    timing_scopes = {
        "write": "litchi_opc::PackageWriter::to_bytes on a prepared OpcPackage",
        "read": "ArchiveReader::new plus indexed reads and Deflate decompression of every member",
        "changed": "SourceBackedPackage::from_vec, one replacement plan, and source-backed publication",
    }
    require(report.get("timing_scope") == timing_scopes[mode],
            f"{label}: timing scope changed")
    require(_uint(report.get("part_count"), f"{label}.part_count") == len(PARTS),
            f"{label}: part count changed")
    report_parts = report.get("parts")
    require(isinstance(report_parts, list) and len(report_parts) == len(PARTS),
            f"{label}: parts are missing or cardinality changed")
    for index, (actual_part, expected_part) in enumerate(zip(report_parts, PARTS)):
        part_label = f"{label}.parts[{index}]"
        require(isinstance(actual_part, Mapping) and set(actual_part) == set(expected_part),
                f"{part_label}: fields changed")
        require(dict(actual_part) == expected_part, f"{part_label}: part identity changed")
    require(report.get("target_part") == "/custom/payload.bin",
            f"{label}: target part changed")
    expected_payload_size = TINY_BYTES if shape == "tiny" else LARGE_BYTES
    require(_uint(report.get("member_payload_bytes"), f"{label}.member_payload_bytes") == expected_payload_size,
            f"{label}: member payload dimension changed")
    require(_uint(report.get("samples_requested"), f"{label}.samples_requested") == samples,
            f"{label}: samples_requested differs from schedule")
    require(_uint(report.get("warmup"), f"{label}.warmup") == warmup,
            f"{label}: warmup differs from schedule")
    expected_report_keys = {
        "schema", "tool", "mode", "shape", "shape_description", "timing_scope",
        "part_count", "parts", "target_part", "member_payload_bytes", "warmup",
        "samples_requested", "source", "expected_archive", "expected_logical",
        "expected_changed_archive", "expected_changed_logical", "samples",
        "artifact",
    }
    require(set(report) == expected_report_keys or set(report) == expected_report_keys - {"artifact"},
            f"{label}: report fields changed")
    source = _identity(report.get("source"), f"{label}.source")
    expected_archive = _identity(report.get("expected_archive"), f"{label}.expected_archive")
    expected_logical = _identity(report.get("expected_logical"), f"{label}.expected_logical")
    expected_changed_archive = _identity(
        report.get("expected_changed_archive"), f"{label}.expected_changed_archive"
    )
    expected_changed_logical = _identity(
        report.get("expected_changed_logical"), f"{label}.expected_changed_logical"
    )
    operation_archive = source if mode in ("write", "read") else expected_changed_archive
    operation_logical = expected_logical if mode in ("write", "read") else expected_changed_logical
    require(expected_archive == operation_archive, f"{label}: expected archive contradicts operation")
    require(expected_logical == operation_logical, f"{label}: expected logical identity contradicts operation")
    rows = report.get("samples")
    require(isinstance(rows, list) and len(rows) == samples,
            f"{label}: sample count differs from schedule")
    elapsed: list[int] = []
    normalized: list[dict[str, Any]] = []
    for index, sample in enumerate(rows):
        sample_label = f"{label}.samples[{index}]"
        require(isinstance(sample, Mapping), f"{sample_label}: sample is not an object")
        require(set(sample) == {
            "index", "elapsed_ns", "archive", "logical", "archive_exact_match",
            "logical_exact_match", "archive_role", "input_archive_sha256",
            "expected_logical_sha256", "uncompressed_bytes",
        },
                f"{sample_label}: fields changed")
        require(_uint(sample.get("index"), f"{sample_label}.index") == index,
                f"{sample_label}.index is not contiguous")
        value = _sample_elapsed(sample, sample_label)
        elapsed.append(value)
        sample_archive = _identity(sample.get("archive"), f"{sample_label}.archive")
        sample_logical = _identity(sample.get("logical"), f"{sample_label}.logical")
        require(sample_archive == operation_archive,
                f"{sample_label}: archive identity differs from expected operation output")
        require(sample_logical == operation_logical,
                f"{sample_label}: logical identity differs from expected operation output")
        require(sample.get("archive_exact_match") is True,
                f"{sample_label}: archive exact-match witness is false")
        require(sample.get("logical_exact_match") is True,
                f"{sample_label}: logical exact-match witness is false")
        role = "input-source-archive" if mode == "read" else "operation-output-archive"
        require(sample.get("archive_role") == role,
                f"{sample_label}: archive role changed")
        require(sample.get("input_archive_sha256") == source["sha256"],
                f"{sample_label}: input archive identity changed")
        require(sample.get("expected_logical_sha256") == operation_logical["sha256"],
                f"{sample_label}: expected logical identity changed")
        require(sample.get("uncompressed_bytes") == operation_logical["bytes"],
                f"{sample_label}: uncompressed byte count changed")
        normalized.append({"index": index, "elapsed_ns": value,
                           "archive": dict(sample_archive), "logical": dict(sample_logical),
                           "archive_exact_match": True, "logical_exact_match": True})
    archive = None
    artifact = None
    if artifact_path is not None:
        artifact_path = Path(artifact_path)
        require(artifact_path.is_file() and not artifact_path.is_symlink(),
                f"{label}: artifact is missing: {artifact_path}")
        artifact = {"bytes": artifact_path.stat().st_size,
                    "sha256": _sha256_file(artifact_path)}
        require(artifact == operation_archive,
                f"{label}: artifact identity differs from expected operation archive")
        require("artifact" in report, f"{label}: report artifact identity is missing")
        require(_identity(report.get("artifact"), f"{label}.artifact") == artifact,
                f"{label}: report artifact identity differs")
        artifact_bytes = _bytes(artifact_path)
        archive = archive_manifest(artifact_bytes)
        _generator_payload_check(archive, artifact_bytes, shape, mode == "changed", label)
        actual_logical = _archive_logical_identity(archive, artifact_bytes, label)
        require(actual_logical == operation_logical,
                f"{label}: independently recomputed logical identity differs")
        require(archive["archive"] == artifact, f"{label}: archive byte identity differs")
    elif "artifact" in report:
        fail(f"{label}: artifact identity requires artifact_path")
    if expected is not None:
        for field in ["expected_archive", "expected_logical", "expected_changed_archive", "expected_changed_logical"]:
            require(report[field] == expected[field], f"{label}: {field} differs from qualification")
        expected_source = _identity(expected.get("source"), f"{label}.expected.source")
        expected_output_value = expected.get("output", expected.get("archive_identity"))
        if expected_output_value is None:
            expected_output_value = (
                expected.get("expected_changed_archive")
                if mode == "changed" else expected.get("expected_archive")
            )
        expected_output = _identity(expected_output_value, f"{label}.expected.output")
        require(source == expected_source, f"{label}: source differs from independent identity")
        require(operation_archive == expected_output,
                f"{label}: output archive differs from independent identity")
        expected_manifest = expected.get("archive_manifest", expected.get("archive"))
        if archive is not None and isinstance(expected_manifest, Mapping) and "members" in expected_manifest:
            compare_archive_manifest(archive, expected_manifest, label=f"{label}.expected_archive")
    return {
        "report": {"path": str(report_file) if report_file is not None else label,
                   "bytes": report_file.stat().st_size if report_file is not None else None,
                   "sha256": _sha256_file(report_file) if report_file is not None else None},
        "schema": REPORT_SCHEMA,
        "tool": REPORT_TOOL,
        "generator": GENERATOR,
        "shape": shape,
        "mode": mode,
        "samples_requested": samples,
        "warmup": warmup,
        "source": source,
        "output": operation_archive,
        "archive_identity": operation_archive,
        "logical": operation_logical,
        "expected_archive": expected_archive,
        "expected_logical": expected_logical,
        "expected_changed_archive": expected_changed_archive,
        "expected_changed_logical": expected_changed_logical,
        "archive": archive,
        "archive_manifest": archive,
        "artifact": artifact,
        "wall_ns": elapsed,
        "p50_ns": _nearest_rank(elapsed, 0.50),
        "p95_ns": _nearest_rank(elapsed, 0.95),
        "p99_ns": _nearest_rank(elapsed, 0.99),
        "mean_ns": statistics.fmean(elapsed),
        "samples": normalized,
        "verification_ok": True,
    }


def parse_report(*args: Any, **kwargs: Any) -> dict[str, Any]:
    """Compatibility alias used by analyzer drafts."""

    return validate_report(*args, **kwargs)


def validate_archive_pair(before: bytes | Path, after: bytes | Path) -> dict[str, Any]:
    """Require exact whole-archive equality before comparing member evidence."""

    left = _bytes(before)
    right = _bytes(after)
    require(left == right, "before/after whole-archive bytes differ")
    manifest = archive_manifest(left)
    return {"equal": True, "archive": manifest["archive"],
            "member_count": manifest["member_count"], "members": manifest["members"]}


__all__ = [
    "ReaderError",
    "archive_manifest",
    "compare_archive_manifest",
    "parse_report",
    "validate_archive_pair",
    "validate_report",
]
