"""Offline, fail-closed readers for the 0844 OPC preservation packet.

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
import re
import struct
import statistics
import zipfile
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


PACKET = Path(__file__).resolve().parent
REPORT_SCHEMA = "litchi.opc-preservation-profile-0844.v1"
REPORT_TOOL = "opc-preservation-probe-0844"
GENERATOR = "probe/src/main.rs::source_parts-0844.v1"
PART_COUNTS = (4, 64, 1024, 4096)
SHAPES = ("shared", "root", "deep-shared", "deep-unique")
MODES = ("publish", "lifecycle")
PROFILE_OWNER = "opc_preservation_probe::profile_publish"
PROFILE_OPTIONS = (
    "--tool=callgrind",
    "--collect-atstart=no",
    "--toggle-collect=" + PROFILE_OWNER,
    "--zero-before=" + PROFILE_OWNER,
    "--dump-after=" + PROFILE_OWNER,
)
PART_BYTES = 256
NATIVE_SAMPLES = 30
NATIVE_WARMUP = 3
QUALIFICATION_SAMPLES = 1
QUALIFICATION_WARMUP = 0
NATIVE_BLOCKS = 12
NATIVE_REPORTS = NATIVE_BLOCKS * len(SHAPES) * len(PART_COUNTS) * len(MODES)
QUALIFICATION_REPORTS = 2 * len(SHAPES) * len(PART_COUNTS) * len(MODES)
PROFILE_REPORTS = 0
HEAP_REPORTS = QUALIFICATION_REPORTS
EXPECTED_TOTAL_REPORTS = NATIVE_REPORTS + QUALIFICATION_REPORTS + PROFILE_REPORTS + HEAP_REPORTS
EXPECTED_TOTAL_SAMPLES = (NATIVE_REPORTS * NATIVE_SAMPLES
                          + QUALIFICATION_REPORTS * QUALIFICATION_SAMPLES
                          + PROFILE_REPORTS * QUALIFICATION_SAMPLES
                          + HEAP_REPORTS * QUALIFICATION_SAMPLES)
MAX_SAMPLES = 100_000
MAX_WARMUP = 100_000
HEX = frozenset("0123456789abcdefABCDEF")
TIMING_SCOPE = {
    "publish": "SourceBackedPackage::prepare_topology validation and write_topology_to_stream via profile_publish on a prepared source package, replacement plan, and empty output Vec",
    "lifecycle": "SourceBackedPackage::from_vec, SourceTopologyPlan::try_replace_part, prepare_topology validation, and write_topology_to_stream via profile_publish; source/replacement/target clones are prepared before the clock",
}
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


def _raw_member(data: bytes, info: zipfile.ZipInfo, label: str,
                span_end: int | None = None) -> tuple[bytes, dict[str, Any]]:
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
    if span_end is None:
        span_end = end
    require(end <= span_end <= len(data), f"{label}: local span is out of bounds")
    return data[start:end], {
        "offset": offset,
        "version": version,
        "flags": flags,
        "method": method,
        "mtime": mtime,
        "mdate": mdate,
        "name_bytes": name_len,
        "extra_bytes": extra_len,
        "span": {"bytes": span_end - offset,
                 "sha256": _sha256(data[offset:span_end])},
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
            next_offset = infos[index + 1].header_offset if index + 1 < len(infos) else archive.start_dir
            require(next_offset > info.header_offset,
                    f"{label}: member offsets are not increasing")
            raw, local = _raw_member(data, info, label, next_offset)
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


def _validate_shape(shape: Any) -> str:
    require(isinstance(shape, str) and shape in SHAPES,
            f"unsupported shape {shape!r}")
    return shape


def _shape_index(first: Any, second: Any | None, label: str) -> tuple[str, int]:
    """Accept the packet API's ``(index, shape)`` and ``(shape, index)`` forms.

    The former keeps index-first helper call sites source-compatible while the
    latter reads naturally in shape-oriented analyzers.  Both forms resolve
    to one strict shape and integer index before any archive oracle runs.
    """

    if isinstance(first, str):
        shape = _validate_shape(first)
        index = second
    else:
        index = first
        shape = "shared" if second is None else second
        shape = _validate_shape(shape)
    require(type(index) is int and index >= 0, f"{label}: invalid part index {index!r}")
    return shape, index


def part_name(index: int | str, shape: str | int = "shared") -> str:
    """Return the exact 0844 OPC URI for one shape and part index."""

    resolved_shape, resolved_index = _shape_index(index, shape, "part name")
    if resolved_shape == "shared":
        return f"/custom/p{resolved_index:05d}.bin"
    if resolved_shape == "root":
        return f"/p{resolved_index:05d}.bin"
    if resolved_shape == "deep-shared":
        return f"/a/b/c/d/e/f/g/h/p{resolved_index:05d}.bin"
    return f"/u{resolved_index:05d}/b/c/d/e/f/g/h/p{resolved_index:05d}.bin"


def part_archive_name(index: int | str, shape: str | int = "shared") -> str:
    return part_name(index, shape).lstrip("/")


def shape_depth(shape: str) -> int:
    return {"shared": 2, "root": 1, "deep-shared": 9, "deep-unique": 9}[_validate_shape(shape)]


def part_payload(index: int, changed: bool = False, part_count: int | None = None) -> bytes:
    """Rebuild one independent 0844 payload.

    Each 256-byte part is ``(index * 17 + offset * 13) % 251``.  A changed
    package flips byte zero of the final part only; ``part_count`` binds that
    mutation to the requested package size.
    """

    require(type(index) is int and index >= 0, f"invalid part index {index!r}")
    if changed:
        require(part_count in PART_COUNTS, "changed payload requires a supported part count")
        require(index == part_count - 1, "only the final part may be changed")
    payload = bytes((index * 17 + offset * 13) % 251 for offset in range(PART_BYTES))
    if changed:
        mutable = bytearray(payload)
        mutable[0] ^= 1
        return bytes(mutable)
    return payload


def payload_for(parts: int, part_index: int, changed: bool = False) -> bytes:
    """Rebuild the fixed-size payload for a supported package."""

    require(parts in PART_COUNTS, f"unsupported part count {parts!r}")
    require(type(part_index) is int and 0 <= part_index < parts,
            f"invalid part index {part_index}")
    return part_payload(part_index, changed=changed, part_count=parts)


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


def _metadata_payloads(parts: int | None = None, shape: str = "shared") -> dict[str, bytes]:
    """Rebuild the fixed metadata members independently for one shape."""

    require(parts is None or parts in PART_COUNTS, "unsupported metadata part count")
    shape = _validate_shape(shape)
    overrides = b"".join(
        f'<Override PartName="{part_name(index, shape)}" '
        'ContentType="application/octet-stream"/>'.encode("ascii")
        for index in range(parts or 0)
    )
    content_types = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" '
        b'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        + overrides
        + b'</Types>'
    )
    return {
        "[Content_Types].xml": content_types,
        "_rels/.rels": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    }


def _generator_payload_check(manifest: Mapping[str, Any], data: bytes, parts: int,
                            changed: bool, label: str, shape: str = "shared") -> dict[str, Any]:
    """Check every generated member's bytes, SHA-256, and CRC independently."""

    members = manifest.get("members")
    require(isinstance(members, list), f"{label}: ZIP member list is missing")
    by_name = {item.get("name"): item for item in members if isinstance(item, Mapping)}
    require(len(by_name) == len(members), f"{label}: ZIP member names are not unique")
    require(parts in PART_COUNTS, f"{label}: unsupported part count")
    shape = _validate_shape(shape)
    expected_part_names = tuple(part_archive_name(index, shape) for index in range(parts))
    expected_names = (*METADATA_ARCHIVE_NAMES, *expected_part_names)
    require(tuple(item.get("name") for item in members) == expected_names,
            f"{label}: ZIP member set differs from the independent generator")
    for index, member in enumerate(members):
        central = member.get("central")
        local = member.get("local")
        require(isinstance(central, Mapping) and isinstance(local, Mapping),
                f"{label}.members[{index}]: ZIP metadata is missing")
        require(central.get("compress_type") == zipfile.ZIP_DEFLATED,
                f"{label}.members[{index}]: compression method is not Deflate")
        require(local.get("method") == zipfile.ZIP_DEFLATED,
                f"{label}.members[{index}]: local compression method is not Deflate")
    expected_payloads = _metadata_payloads(parts, shape)
    for index, name in enumerate(expected_part_names):
        expected_payloads[name] = payload_for(parts, index, changed and index == parts - 1)
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
        "parts": parts,
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


# The implementation below is the strict 0844 report adapter.
def _part_spec_0844(index: int, shape: str = "shared") -> dict[str, Any]:
    shape = _validate_shape(shape)
    return {
        "name": part_name(index, shape),
        "archive_name": part_archive_name(index, shape),
        "content_type": "application/octet-stream",
        "role": "opaque payload part",
    }


def _validate_part_specs_0844(value: Any, parts: int, label: str,
                              shape: str = "shared") -> None:
    if value is None:
        return
    require(isinstance(value, list) and len(value) == parts,
            f"{label}: part specification cardinality changed")
    for index, actual in enumerate(value):
        require(isinstance(actual, Mapping), f"{label}[{index}] is not an object")
        require(dict(actual) == _part_spec_0844(index, shape),
                f"{label}[{index}] part identity changed")


def _member_span_0844(member: Mapping[str, Any]) -> dict[str, Any]:
    local = member.get("local")
    central = member.get("central")
    require(isinstance(local, Mapping) and isinstance(central, Mapping),
            "archive member metadata is missing")
    local_copy = dict(local)
    local_copy.pop("offset", None)
    return {
        "name": member.get("name"),
        "raw": dict(member.get("raw", {})),
        "logical": dict(member.get("logical", {})),
        "crc32": member.get("crc32"),
        "central": dict(central),
        "local": local_copy,
    }


def compare_untouched_spans(source: Mapping[str, Any], output: Mapping[str, Any],
                            parts: int, label: str = "archive",
                            shape: str = "shared") -> dict[str, Any]:
    """Prove raw compressed spans and metadata for every untouched member."""

    require(parts in PART_COUNTS, f"{label}: unsupported part count")
    shape = _validate_shape(shape)
    left = source.get("members")
    right = output.get("members")
    require(isinstance(left, list) and isinstance(right, list),
            f"{label}: archive member lists are missing")
    require(len(left) == len(right) == parts + len(METADATA_ARCHIVE_NAMES),
            f"{label}: member count differs")
    target = part_archive_name(parts - 1, shape)
    preserved: list[str] = []
    for index, (left_member, right_member) in enumerate(zip(left, right)):
        require(isinstance(left_member, Mapping) and isinstance(right_member, Mapping),
                f"{label}: member {index} is not an object")
        name = left_member.get("name")
        require(name == right_member.get("name"), f"{label}: member ordering changed")
        if name == target:
            continue
        require(_member_span_0844(left_member) == _member_span_0844(right_member),
                f"{label}: untouched member {name!r} changed")
        preserved.append(str(name))
    require(target in {item.get("name") for item in left},
            f"{label}: target part is missing")
    return {"target": target, "count": len(preserved), "members": preserved}


_LOCAL_STABLE_FIELDS_0844 = (
    "version", "flags", "method", "mtime", "mdate", "name_bytes", "extra_bytes",
)
_CENTRAL_STABLE_FIELDS_0844 = (
    "compress_type", "flag_bits", "create_system", "create_version",
    "extract_version", "reserved", "volume", "internal_attr", "external_attr",
    "date_time", "comment", "extra",
)


def _stable_fields(value: Any, fields: Sequence[str], label: str) -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label}: metadata is missing")
    return {field: value.get(field) for field in fields}


def compare_changed_target_metadata(source: Mapping[str, Any], output: Mapping[str, Any],
                                    parts: int, label: str = "archive",
                                    shape: str = "shared") -> dict[str, Any]:
    """Require stable local/central target framing across the replacement.

    The target's CRC, compressed length, and logical length are allowed to
    change because the replacement payload is deliberately different.  Every
    other retained local-header and central-directory field must remain equal.
    """

    require(parts in PART_COUNTS, f"{label}: unsupported part count")
    shape = _validate_shape(shape)
    left = source.get("members")
    right = output.get("members")
    require(isinstance(left, list) and isinstance(right, list),
            f"{label}: archive member lists are missing")
    target = part_archive_name(parts - 1, shape)
    source_target = next((item for item in left if isinstance(item, Mapping)
                          and item.get("name") == target), None)
    output_target = next((item for item in right if isinstance(item, Mapping)
                          and item.get("name") == target), None)
    require(source_target is not None and output_target is not None,
            f"{label}: target part is missing")
    source_local = _stable_fields(source_target.get("local"), _LOCAL_STABLE_FIELDS_0844,
                                  f"{label}.source.target.local")
    output_local = _stable_fields(output_target.get("local"), _LOCAL_STABLE_FIELDS_0844,
                                  f"{label}.output.target.local")
    require(source_local == output_local,
            f"{label}: changed target local-header stable metadata differs")
    source_central = _stable_fields(source_target.get("central"), _CENTRAL_STABLE_FIELDS_0844,
                                    f"{label}.source.target.central")
    output_central = _stable_fields(output_target.get("central"), _CENTRAL_STABLE_FIELDS_0844,
                                    f"{label}.output.target.central")
    require(source_central == output_central,
            f"{label}: changed target central-directory stable metadata differs")
    return {
        "target": target,
        "local_stable": source_local,
        "central_stable": source_central,
        "crc_and_lengths_may_change": True,
    }


def _output_identities_0844(report: Mapping[str, Any], label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    output = report.get("expected_output")
    logical = report.get("expected_output_logical")
    return (_identity(output, f"{label}.expected_output"),
            _identity(logical, f"{label}.expected_output_logical"))


def _validate_generator_0844(value: Any, parts: int, shape: str, label: str) -> None:
    require(isinstance(value, Mapping), f"{label}: generator is missing")
    shape = _validate_shape(shape)
    expected = {
        "shape": shape,
        "part_count": parts,
        "member_count": parts + len(METADATA_ARCHIVE_NAMES),
        "payload_bytes": PART_BYTES,
        "part_name_pattern": {
            "shared": "/custom/p{index:05}.bin",
            "root": "/p{index:05}.bin",
            "deep-shared": "/a/b/c/d/e/f/g/h/p{index:05}.bin",
            "deep-unique": "/u{index:05}/b/c/d/e/f/g/h/p{index:05}.bin",
        }[shape],
        "depth": shape_depth(shape),
        "payload_formula": "byte[i] = (part_index * 17 + i * 13) % 251",
        "target_part_index": parts - 1,
        "target_part_name": part_name(parts - 1, shape),
        "replacement": "target payload with byte[0] XOR 1",
        "metadata_members": list(METADATA_ARCHIVE_NAMES),
    }
    require(dict(value) == expected, f"{label}: generator contract changed")


def _sample_0844(sample: Mapping[str, Any], label: str, source: Mapping[str, Any],
                 output: Mapping[str, Any], logical: Mapping[str, Any],
                 index: int, parts: int) -> dict[str, Any]:
    required = {"index", "elapsed_ns", "output", "logical", "member_count", "hash_witness"}
    require(set(sample) == required, f"{label}: sample fields changed")
    require(_uint(sample.get("index"), f"{label}.index") == index,
            f"{label}.index is not contiguous")
    elapsed = _sample_elapsed(sample, label)
    archive = _identity(sample.get("output"), f"{label}.output")
    sample_logical = _identity(sample.get("logical"), f"{label}.logical")
    require(archive == output, f"{label}: output identity differs from expected output")
    require(sample_logical == logical, f"{label}: logical identity differs from expected output")
    require(_uint(sample.get("member_count"), f"{label}.member_count") == parts + 2,
            f"{label}: member count is invalid")
    witness = sample.get("hash_witness")
    require(isinstance(witness, Mapping), f"{label}: hash witness is missing")
    require(set(witness) == {"output_sha256", "logical_sha256",
                             "output_exact_expected", "logical_exact_expected"},
            f"{label}: hash witness fields changed")
    require(witness.get("output_sha256") == output["sha256"],
            f"{label}: output witness differs")
    require(witness.get("logical_sha256") == logical["sha256"],
            f"{label}: logical witness differs")
    require(witness.get("output_exact_expected") is True
            and witness.get("logical_exact_expected") is True,
            f"{label}: exact output witness is false")
    return {"index": index, "elapsed_ns": elapsed, "archive": archive,
            "logical": sample_logical, "archive_exact_match": True,
            "logical_exact_match": True}


def validate_report(path: Path | Mapping[str, Any], parts: int, mode: str,
                    samples: int, warmup: int, artifact_path: Path | None = None,
                    source_path: Path | None = None,
                    expected: Mapping[str, Any] | None = None,
                    source_artifact_path: Path | None = None,
                    lane: str | None = None,
                    shape: str = "shared") -> dict[str, Any]:
    """Validate one 0844 report and independently inspect retained ZIPs.

    A retained output archive is accepted only with a retained source archive.
    This makes the unchanged compressed-span claim auditable rather than
    accepting a source digest emitted by the probe as its own oracle.
    """

    if source_artifact_path is None:
        source_artifact_path = source_path
    require(parts in PART_COUNTS, f"unsupported part count {parts!r}")
    shape = _validate_shape(shape)
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
    require(report.get("shape") == shape, f"{label}: shape changed")
    require(report.get("mode") == mode, f"{label}: mode changed")
    common_keys = {
        "schema", "tool", "shape", "mode", "timing_description", "generator", "warmup",
        "samples_requested", "source", "source_logical", "expected_output",
        "expected_output_logical", "samples",
    }
    qualification_keys = common_keys | {"artifact", "artifact_role", "source_artifact"}
    if "artifact" in report or "artifact_role" in report or "source_artifact" in report:
        require(set(report) == qualification_keys, f"{label}: report fields changed")
        if lane is not None:
            require(lane == "qualification", f"{label}: artifact report lane changed")
    else:
        require(set(report) == common_keys, f"{label}: report fields changed")
        if lane is not None:
            require(lane in {"native", "profiles", "heap"},
                    f"{label}: report lane differs from schedule")
    require(report.get("timing_description") == TIMING_SCOPE[mode],
            f"{label}: timing description changed")
    _validate_generator_0844(report.get("generator"), parts, shape, f"{label}.generator")
    require(_uint(report.get("samples_requested"), f"{label}.samples_requested") == samples,
            f"{label}: samples_requested differs from schedule")
    require(_uint(report.get("warmup"), f"{label}.warmup") == warmup,
            f"{label}: warmup differs from schedule")
    source = _identity(report.get("source"), f"{label}.source")
    source_logical = _identity(report.get("source_logical"), f"{label}.source_logical")
    output, logical = _output_identities_0844(report, label)
    rows = report.get("samples")
    require(isinstance(rows, list) and len(rows) == samples,
            f"{label}: sample count differs from schedule")
    elapsed: list[int] = []
    normalized: list[dict[str, Any]] = []
    for index, sample in enumerate(rows):
        require(isinstance(sample, Mapping),
                f"{label}.samples[{index}] is not an object")
        row = _sample_0844(sample, f"{label}.samples[{index}]",
                           source, output, logical, index, parts)
        normalized.append(row)
        elapsed.append(row["elapsed_ns"])
    source_manifest = None
    output_manifest = None
    source_artifact = None
    artifact = None
    report_has_artifacts = "artifact" in report
    if report_has_artifacts:
        require(artifact_path is not None and source_artifact_path is not None,
                f"{label}: qualification report requires output and source ZIPs")
    if source_artifact_path is not None:
        source_bytes = _bytes(Path(source_artifact_path))
        source_manifest = archive_manifest(source_bytes)
        require(source_manifest["archive"] == source,
                f"{label}: source ZIP identity differs")
        _generator_payload_check(source_manifest, source_bytes, parts, False,
                                 label + ".source", shape)
        require(_archive_logical_identity(source_manifest, source_bytes, label + ".source")
                == source_logical, f"{label}: source logical identity differs")
        source_artifact = source_manifest["archive"]
    if artifact_path is not None:
        require(source_artifact_path is not None,
                f"{label}: output artifact requires retained source ZIP")
        output_bytes = _bytes(Path(artifact_path))
        output_manifest = archive_manifest(output_bytes)
        artifact = output_manifest["archive"]
        require(artifact == output, f"{label}: output ZIP identity differs")
        if "artifact" in report:
            require(_identity(report["artifact"], f"{label}.artifact") == artifact,
                    f"{label}: report artifact identity differs")
        if "artifact_role" in report:
            require(report["artifact_role"] == "changed-output-archive",
                    f"{label}: artifact role changed")
        _generator_payload_check(output_manifest, output_bytes, parts, True,
                                 label + ".output", shape)
        require(_archive_logical_identity(output_manifest, output_bytes, label) == logical,
                f"{label}: output logical identity differs")
        require(source_manifest is not None, f"{label}: source manifest is missing")
        spans = compare_untouched_spans(source_manifest, output_manifest, parts, label, shape)
        target_metadata = compare_changed_target_metadata(
            source_manifest, output_manifest, parts, label, shape)
        source_members = {row["name"]: row for row in source_manifest["members"]}
        output_members = {row["name"]: row for row in output_manifest["members"]}
        target = part_archive_name(parts - 1, shape)
        require(source_members[target]["logical"] != output_members[target]["logical"],
                f"{label}: target part did not change")
    elif "artifact" in report:
        require(lane in {"native", "profiles", "heap"}
                or report.get("lane") in {"native", "profiles", "heap"},
                f"{label}: artifact identity requires artifact_path")
    if "source_artifact" in report and source_artifact is not None:
        require(_identity(report["source_artifact"], f"{label}.source_artifact")
                == source_artifact,
                f"{label}: source artifact identity differs")
    if expected is not None:
        require(isinstance(expected, Mapping),
                f"{label}: expected identity is not an object")
        expected_source = _identity(expected.get("source"), f"{label}.expected.source")
        expected_source_logical = _identity(
            expected.get("source_logical"), f"{label}.expected.source_logical")
        expected_output = _identity(expected.get("expected_output"),
                                    f"{label}.expected.output")
        expected_output_logical = _identity(
            expected.get("expected_output_logical"),
            f"{label}.expected.expected_output_logical")
        require(source == expected_source, f"{label}: source differs from qualification")
        require(source_logical == expected_source_logical,
                f"{label}: source logical identity differs from qualification")
        require(output == expected_output, f"{label}: output differs from qualification")
        require(logical == expected_output_logical,
                f"{label}: output logical identity differs from qualification")
    return {
        "report": {
            "path": str(report_file) if report_file is not None else label,
            "bytes": report_file.stat().st_size if report_file is not None else None,
            "sha256": _sha256_file(report_file) if report_file is not None else None,
        },
        "schema": REPORT_SCHEMA,
        "tool": REPORT_TOOL,
        "shape": shape,
        "generator": GENERATOR,
        "parts": parts,
        "mode": mode,
        "samples_requested": samples,
        "warmup": warmup,
        "source": source,
        "source_logical": source_logical,
        "output": output,
        "archive_identity": output,
        "logical": logical,
        "output_logical": logical,
        "archive": output_manifest,
        "source_archive": source_manifest,
        "artifact": artifact,
        "source_artifact": source_artifact,
        "untouched": spans if artifact is not None else None,
        "target_metadata": target_metadata if artifact is not None else None,
        "wall_ns": elapsed,
        "p50_ns": _nearest_rank(elapsed, 0.50),
        "p95_ns": _nearest_rank(elapsed, 0.95),
        "p99_ns": _nearest_rank(elapsed, 0.99),
        "mean_ns": statistics.fmean(elapsed),
        "samples": normalized,
        "verification_ok": True,
    }


__all__ = sorted(set(__all__) | {
    "PART_COUNTS", "SHAPES", "MODES", "PROFILE_OWNER", "PROFILE_OPTIONS",
    "PART_BYTES", "EXPECTED_TOTAL_REPORTS", "EXPECTED_TOTAL_SAMPLES",
    "part_name", "part_archive_name", "shape_depth", "part_payload", "payload_for",
    "compare_untouched_spans", "compare_changed_target_metadata",
})


# Packet-local Callgrind grammar.  It is intentionally independent of
# callgrind_annotate and retains file/source attribution for the ROI review.
FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
CALLS_RE = re.compile(r"^calls=([+\-]?[0-9][0-9,]*)")
EVENTS_RE = re.compile(r"^events:\s*(.*)$")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")
TOTALS_RE = re.compile(r"^totals:\s*(.*)$")
PART_RE = re.compile(r"^part:\s*(\d+)$")
TRIGGER_RE = re.compile(r"^desc:\s+Trigger:\s+(.*)$")
CMD_RE = re.compile(r"^cmd:\s?(.*)$")
FILE_RE = re.compile(r"^(fl|fi|fe|cfi|cfl)=\((\d+)\)(?:\s+(.*))?$")


def _parse_ints_0844(text: str) -> list[int]:
    values: list[int] = []
    for token in text.split():
        try:
            values.append(int(token.replace(",", "")))
        except ValueError:
            break
    return values


def _header_0844(lines: Sequence[str], label: str) -> dict[str, Any]:
    events = None
    summary = None
    totals = None
    part = None
    trigger = None
    command = None
    for line in lines:
        stripped = line.strip()
        match = EVENTS_RE.match(stripped)
        if match:
            require(events is None, f"{label}: duplicate events header")
            events = match.group(1).split()
        match = SUMMARY_RE.match(stripped)
        if match:
            require(summary is None, f"{label}: duplicate summary")
            values = _parse_ints_0844(match.group(1))
            require(values, f"{label}: summary is not numeric")
            summary = values[0]
        match = TOTALS_RE.match(stripped)
        if match:
            require(totals is None, f"{label}: duplicate totals")
            values = _parse_ints_0844(match.group(1))
            require(values, f"{label}: totals is not numeric")
            totals = values[0]
        match = PART_RE.match(stripped)
        if match:
            require(part is None, f"{label}: duplicate part")
            part = int(match.group(1))
        match = TRIGGER_RE.match(stripped)
        if match:
            require(trigger is None, f"{label}: duplicate trigger")
            trigger = match.group(1)
        if command is None:
            match = CMD_RE.match(stripped)
            if match:
                command = match.group(1)
    require(events == ["Ir"], f"{label}: event set is not exactly Ir")
    require(summary is not None and summary >= 0, f"{label}: summary is missing or negative")
    require(totals is not None and totals >= 0, f"{label}: totals is missing or negative")
    require(part is not None, f"{label}: part is missing")
    require(isinstance(trigger, str) and trigger, f"{label}: trigger is missing")
    return {"events": events, "summary_ir": summary, "totals_ir": totals,
            "part": part, "trigger": trigger, "command": command}


def _add_name_0844(names: dict[int, str], function_id: int, name: str, label: str) -> None:
    if not name:
        return
    prior = names.get(function_id)
    require(prior is None or prior == name,
            f"{label}: function id {function_id} has conflicting names")
    names[function_id] = name


def _cost_0844(line: str) -> tuple[int, str] | None:
    fields = line.split()
    if len(fields) < 2:
        return None
    position = fields[0]
    if position != "*" and not re.fullmatch(r"[+\-]?\d+", position):
        return None
    try:
        return int(fields[1].replace(",", "")), position
    except ValueError:
        return None


def _position_0844(position: str) -> str:
    if position == "*":
        return "wildcard"
    return "relative" if position.startswith(("+", "-")) else "absolute"


def parse_raw(path: Path, allow_empty: bool = False) -> dict[str, Any]:
    """Parse one raw Callgrind file without executing external tools."""

    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing Callgrind file: {path}")
    try:
        lines = path.read_text(encoding="utf-8", errors="strict").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"cannot read {path}: {error}")
    label = str(path)
    header = _header_0844(lines, label)
    file_names: dict[int, str] = {}
    for line in lines:
        match = FILE_RE.match(line.strip())
        if match and match.group(3):
            file_id = int(match.group(2))
            prior = file_names.get(file_id)
            require(prior is None or prior == match.group(3),
                    f"{label}: file id {file_id} has conflicting names")
            file_names[file_id] = match.group(3)
    names: dict[int, str] = {}
    declarations = 0
    for line in lines:
        match = FUNCTION_RE.match(line.strip())
        if match and match.group(3):
            declarations += 1
            _add_name_0844(names, int(match.group(2)), match.group(3), label)
    functions: dict[int, dict[str, Any]] = {}
    current_id: int | None = None
    current_file = ""
    pending: dict[str, Any] | None = None
    costs = 0
    edges = 0
    positions: dict[str, int] = {}
    for line_number, line in enumerate(lines, 1):
        stripped = line.strip()
        file_match = FILE_RE.match(stripped)
        if file_match:
            file_id = int(file_match.group(2))
            if file_match.group(1) in {"cfi", "cfl"}:
                if file_match.group(3):
                    current = file_names.get(file_id)
                    require(current is None or current == file_match.group(3),
                            f"{label}:{line_number}: conflicting callee source file")
                    file_names[file_id] = file_match.group(3)
                continue
            if file_match.group(3):
                current_file = file_match.group(3)
            else:
                # Callgrind may reference a file ID whose declaration was
                # emitted only in a different context. Preserve the ID as a
                # stable attribution token instead of dropping the record.
                current_file = file_names.get(file_id, f"(file:{file_id})")
            require(current_file, f"{label}:{line_number}: empty source file")
            continue
        if stripped.startswith(("fl=", "fi=", "fe=")):
            current_file = stripped[3:]
            require(current_file, f"{label}:{line_number}: empty source file")
            continue
        match = FUNCTION_RE.match(stripped)
        if match:
            kind, text_id, name = match.groups()
            function_id = int(text_id)
            if name:
                _add_name_0844(names, function_id, name, f"{label}:{line_number}")
            if kind == "fn":
                require(pending is None,
                        f"{label}:{line_number}: cfn is missing a cost")
                current_id = function_id
                function = functions.setdefault(function_id, {
                    "id": function_id, "name": "", "file": current_file,
                    "records": 0, "self_ir": 0, "edges": [],
                })
                require(not function["file"] or function["file"] == current_file,
                        f"{label}:{line_number}: function source file changed")
                function["records"] += 1
            else:
                require(current_id is not None,
                        f"{label}:{line_number}: cfn outside fn")
                require(pending is None,
                        f"{label}:{line_number}: nested cfn before cost")
                pending = {"callee_id": function_id, "calls": None,
                           "line": line_number, "file": current_file}
            continue
        match = CALLS_RE.match(stripped)
        if match:
            require(current_id is not None and pending is not None,
                    f"{label}:{line_number}: calls outside cfn")
            require(pending["calls"] is None,
                    f"{label}:{line_number}: duplicate calls")
            calls = int(match.group(1).replace(",", ""))
            require(calls >= 0, f"{label}:{line_number}: negative call count")
            pending["calls"] = calls
            continue
        if current_id is None:
            continue
        parsed = _cost_0844(stripped)
        if parsed is None:
            continue
        value, position = parsed
        require(value >= 0, f"{label}:{line_number}: negative Ir cost")
        kind = _position_0844(position)
        positions[kind] = positions.get(kind, 0) + 1
        costs += 1
        function = functions[current_id]
        if pending is not None:
            require(pending["calls"] is not None,
                    f"{label}:{line_number}: cfn has no calls")
            function["edges"].append({
                "caller_id": current_id,
                "callee_id": pending["callee_id"],
                "calls": pending["calls"],
                "inclusive_ir": value,
                "line": pending["line"],
                "position": position,
                "file": pending["file"],
            })
            edges += 1
            pending = None
        else:
            function["self_ir"] += value
    require(pending is None, f"{label}: unterminated cfn record")
    for function_id, function in functions.items():
        function["name"] = names.get(function_id, "")
        require(function["file"], f"{label}: function {function_id} lacks source attribution")
    if allow_empty and header["summary_ir"] == 0 and not names and costs == 0:
        functions = {}
    else:
        require(names, f"{label}: no compressed function names")
        require(costs > 0, f"{label}: no Ir cost records")
    return {
        "path": label,
        "sha256": _sha256_file(path),
        "bytes": path.stat().st_size,
        "header": header,
        "names": names,
        "functions": functions,
        "statistics": {
            "function_records": sum(item["records"] for item in functions.values()),
            "unique_functions": len(functions),
            "compressed_name_declarations": declarations,
            "cost_records": costs,
            "call_edge_records": edges,
            "position_kinds": dict(sorted(positions.items())),
        },
    }


def aggregate_edges(function: Mapping[str, Any], names: Mapping[int, str]) -> list[dict[str, Any]]:
    grouped: dict[int, dict[str, Any]] = {}
    for edge in function.get("edges", []):
        if edge["calls"] <= 0 or edge["inclusive_ir"] <= 0:
            continue
        item = grouped.setdefault(edge["callee_id"], {
            "callee_id": edge["callee_id"],
            "callee": names.get(edge["callee_id"], ""),
            "calls": 0,
            "edge_count": 0,
            "inclusive_ir": 0,
        })
        item["calls"] += edge["calls"]
        item["edge_count"] += 1
        item["inclusive_ir"] += edge["inclusive_ir"]
    return sorted(grouped.values(), key=lambda item: (-item["inclusive_ir"], item["callee_id"]))


def function_view(function: Mapping[str, Any], names: Mapping[int, str]) -> dict[str, Any]:
    direct = aggregate_edges(function, names)
    return {
        "id": function["id"],
        "name": function.get("name") or "<unnamed>",
        "file": function.get("file", ""),
        "source": function.get("file", ""),
        "records": function["records"],
        "self_ir": function["self_ir"],
        "direct_children": direct,
        "direct_children_ir": sum(item["inclusive_ir"] for item in direct),
    }


def incoming_edges(parsed: Mapping[str, Any], target_id: int) -> list[dict[str, Any]]:
    names = parsed["names"]
    rows: list[dict[str, Any]] = []
    for function in parsed["functions"].values():
        for edge in function["edges"]:
            if edge["callee_id"] == target_id and edge["calls"] > 0 and edge["inclusive_ir"] > 0:
                rows.append({
                    **edge,
                    "caller": names.get(edge["caller_id"], "<unnamed>"),
                    "callee": names.get(edge["callee_id"], "<unnamed>"),
                })
    return rows


def _normalize_command_0844(value: str | Sequence[str]) -> str:
    if isinstance(value, str):
        return " ".join(value.split())
    require(isinstance(value, Sequence) and not isinstance(value, (bytes, bytearray)),
            "command is not a string or sequence")
    return " ".join(str(item) for item in value)


def _launcher_target_0844(command: Sequence[str]) -> str:
    values = [str(item) for item in command]
    for index, value in enumerate(values):
        if "opc-preservation-probe" in Path(value).name:
            return " ".join(values[index:])
    return " ".join(values)


def validate_profile(numbered: Path, terminal: Path, owner: str = PROFILE_OWNER,
                    command: Sequence[str] | str | None = None,
                    expected_command: str | None = None) -> dict[str, Any]:
    """Validate one scoped dump and its empty terminal Callgrind dump."""

    require(owner == PROFILE_OWNER, "profile owner changed")
    number = parse_raw(Path(numbered))
    final = parse_raw(Path(terminal), allow_empty=True)
    label = str(numbered)
    require(number["header"]["part"] == 1, f"{label}: numbered part changed")
    require(number["header"]["trigger"] == "--dump-after=" + owner,
            f"{label}: collection trigger changed")
    require(number["header"]["summary_ir"] > 0, f"{label}: scoped summary is zero")
    require(number["header"]["totals_ir"] == number["header"]["summary_ir"],
            f"{label}: totals do not match summary")
    require(final["header"]["part"] == 2 and final["header"]["trigger"] == "Program termination",
            f"{terminal}: termination dump changed")
    require(final["header"]["summary_ir"] == 0 and final["header"]["totals_ir"] == 0,
            f"{terminal}: termination dump is not zero Ir")
    require(number["header"]["command"] and final["header"]["command"],
            f"{label}: raw command is missing")
    require(_normalize_command_0844(number["header"]["command"])
            == _normalize_command_0844(final["header"]["command"]),
            f"{label}: dump commands differ")
    if expected_command is not None:
        require(_normalize_command_0844(number["header"]["command"])
                == _normalize_command_0844(expected_command),
                f"{label}: raw command differs from expected command")
    if command is not None:
        launcher = command if not isinstance(command, str) else command.split()
        values = [str(item) for item in launcher]
        for option in PROFILE_OPTIONS:
            require(option in values, f"{label}: launcher command lacks {option}")
        require(_normalize_command_0844(number["header"]["command"])
                == _normalize_command_0844(_launcher_target_0844(values)),
                f"{label}: normalized launcher target differs from raw command")
    owner_ids = [fid for fid, function in number["functions"].items()
                 if function["name"] == owner]
    require(len(owner_ids) == 1, f"{label}: owner resolved to {owner_ids}")
    owner_id = owner_ids[0]
    incoming = incoming_edges(number, owner_id)
    require(len(incoming) == 1 and incoming[0]["calls"] == 1,
            f"{label}: owner incoming call is not exactly one")
    require(incoming[0]["inclusive_ir"] == number["header"]["summary_ir"],
            f"{label}: owner incoming Ir differs from summary")
    owner_row = function_view(number["functions"][owner_id], number["names"])
    require(owner_row["self_ir"] + owner_row["direct_children_ir"]
            == number["header"]["summary_ir"],
            f"{label}: owner partition does not reconstruct summary")
    function_rows = [function_view(function, number["names"])
                     for function in number["functions"].values()]
    all_self = sum(row["self_ir"] for row in function_rows)
    require(all_self == number["header"]["summary_ir"],
            f"{label}: function self costs do not reconstruct summary")
    for row in function_rows:
        require(row["file"], f"{label}: function lacks file attribution")
    function_rows.sort(key=lambda row: (-row["self_ir"], row["name"], row["id"]))
    file_rows: dict[str, dict[str, Any]] = {}
    for row in function_rows:
        item = file_rows.setdefault(row["file"], {"file": row["file"], "self_ir": 0, "functions": 0})
        item["self_ir"] += row["self_ir"]
        item["functions"] += 1
    require(sum(item["self_ir"] for item in file_rows.values()) == all_self,
            f"{label}: file attribution does not conserve self cost")
    unexpected = Path(numbered).with_name(Path(numbered).name.rsplit(".1", 1)[0] + ".2")
    require(not unexpected.exists(), f"{label}: unexpected second numbered dump")
    return {
        "file": label,
        "sha256": number["sha256"],
        "bytes": number["bytes"],
        "termination": {
            "file": str(terminal),
            "sha256": final["sha256"],
            "bytes": final["bytes"],
            "summary_ir": 0,
        },
        "summary_ir": number["header"]["summary_ir"],
        "totals_ir": number["header"]["totals_ir"],
        "owner": {
            "id": owner_id,
            "name": owner,
            "file": owner_row["file"],
            "source": owner_row["source"],
            "incoming": incoming[0],
            "self_ir": owner_row["self_ir"],
            "direct_children_ir": owner_row["direct_children_ir"],
            "direct_children": owner_row["direct_children"],
        },
        "function_rows": function_rows,
        "file_rows": sorted(file_rows.values(), key=lambda row: row["file"]),
        "source_attribution": sorted(file_rows.values(), key=lambda row: row["file"]),
        "top_self_functions": [row for row in function_rows if row["self_ir"] > 0][:20],
        "top_edges": owner_row["direct_children"],
        "call_rows": [dict(row) for row in incoming],
        "function_totals": {
            "summary_ir": number["header"]["summary_ir"],
            "owner_self_ir": owner_row["self_ir"],
            "owner_direct_children_ir": owner_row["direct_children_ir"],
            "all_function_self_ir": all_self,
        },
        "parser": number["statistics"],
        "validation": {
            "exact_trigger": True,
            "summary_nonzero": True,
            "totals_equal_summary": True,
            "termination_zero_ir": True,
            "exactly_one_positive_owner_call": True,
            "owner_inclusive_matches_summary": True,
            "owner_self_plus_direct_children_equals_summary": True,
            "all_function_self_ir_equals_summary": True,
            "file_source_attribution_conserved": True,
            "normalized_command_equal": True,
            "single_publication": True,
        },
    }


def validate_callgrind(*args: Any, **kwargs: Any) -> dict[str, Any]:
    return validate_profile(*args, **kwargs)


def validate_profile_pair(*args: Any, **kwargs: Any) -> dict[str, Any]:
    return validate_profile(*args, **kwargs)


__all__ = sorted(set(__all__) | {
    "parse_raw", "validate_profile", "validate_callgrind",
    "validate_profile_pair", "aggregate_edges", "incoming_edges",
    "function_view",
})
