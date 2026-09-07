#!/usr/bin/env python3
"""Independent oracle for a source-backed ODP tail append.

The producer is deliberately not imported here.  ZIP records and payloads are
decoded from the archive bytes, while content.xml is parsed with Expat and its
byte offsets are used only to prove the source-prefix/source-suffix relation.
The script never rewrites an input archive.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import struct
import sys
import zlib
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable
from xml.parsers import expat


SCHEMA = "litchi-0457-native-odp-append-oracle-v1"
MAX_INPUT_BYTES = 512 * 1024 * 1024
OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
PRESENTATION_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
CONTENT_NAME = b"content.xml"


class OracleError(Exception):
    """A validation failure in an otherwise supported archive shape."""

    kind = "validation_error"


class OracleRefusal(OracleError):
    """A safe refusal for an archive shape this independent parser will not guess."""

    kind = "unsupported_oracle_shape"


def fail(message: str) -> None:
    raise OracleError(message)


def refuse(message: str) -> None:
    raise OracleRefusal(message)


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def u16(data: bytes, offset: int) -> int:
    return struct.unpack_from("<H", data, offset)[0]


def u32(data: bytes, offset: int) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def split_xml_name(name: str) -> tuple[str, str]:
    if "}" in name:
        return tuple(name.split("}", 1))  # type: ignore[return-value]
    return "", name


def xml_name(namespace: str, local: str) -> str:
    return f"{namespace}}}{local}" if namespace else local


def tag_end(raw: bytes, start: int) -> int:
    """Return the byte just after the tag whose '<' is at start."""

    quote: int | None = None
    for index in range(start, len(raw)):
        byte = raw[index]
        if quote is not None:
            if byte == quote:
                quote = None
        elif byte in (ord('"'), ord("'")):
            quote = byte
        elif byte == ord(">"):
            return index + 1
    fail(f"XML tag at byte {start} has no closing '>'")
    raise AssertionError


@dataclass
class XmlNode:
    name: str
    attrs: dict[str, str]
    start: int
    end: int | None = None
    children: list["XmlNode"] = field(default_factory=list)
    content: list[tuple[str, Any]] = field(default_factory=list)


@dataclass
class ContentFacts:
    raw: bytes
    root: XmlNode
    presentation: XmlNode
    pages: list[XmlNode]
    direct_events: list[tuple[str, Any]]
    leading_direct_children: list[str]
    trailing_direct_children: list[str]
    pages_contiguous: bool
    insert_at: int

    @property
    def page_names(self) -> list[str | None]:
        return [page.attrs.get(xml_name(DRAW_NS, "name")) for page in self.pages]


def parse_content(raw: bytes) -> ContentFacts:
    parser = expat.ParserCreate(namespace_separator="}")
    root: XmlNode | None = None
    stack: list[XmlNode] = []

    def on_start(name: str, attrs: dict[str, str]) -> None:
        nonlocal root
        start = parser.CurrentByteIndex
        end = tag_end(raw, start)
        self_closing = raw[start : end - 1].rstrip().endswith(b"/")
        node = XmlNode(name=name, attrs=dict(attrs), start=start)
        if self_closing:
            node.end = end
        if stack:
            parent = stack[-1]
            parent.children.append(node)
            parent.content.append(("element", node))
        elif root is not None:
            fail("content.xml contains more than one root element")
        else:
            root = node
        stack.append(node)

    def on_end(name: str) -> None:
        if not stack:
            fail("content.xml element stack underflow")
        node = stack.pop()
        if node.name != name:
            fail(f"content.xml close name differs: {name!r} versus {node.name!r}")
        if node.end is None:
            node.end = tag_end(raw, parser.CurrentByteIndex)

    def on_text(text: str) -> None:
        if stack:
            stack[-1].content.append(("text", text))

    def on_comment(comment: str) -> None:
        if stack:
            stack[-1].content.append(("comment", comment))

    def on_pi(target: str, data: str) -> None:
        if stack:
            stack[-1].content.append(("pi", (target, data)))

    parser.StartElementHandler = on_start
    parser.EndElementHandler = on_end
    parser.CharacterDataHandler = on_text
    parser.CommentHandler = on_comment
    parser.ProcessingInstructionHandler = on_pi
    try:
        parser.Parse(raw, True)
    except OracleError:
        raise
    except expat.ExpatError as error:
        fail(
            "content.xml Expat parse failed at "
            f"line {error.lineno}, column {error.offset}: {error}"
        )
    if stack or root is None or root.end is None:
        fail("content.xml ended with an open element")

    root_name = split_xml_name(root.name)
    if root_name != (OFFICE_NS, "document-content"):
        fail(f"content.xml root is {root.name!r}, expected office:document-content")
    bodies = [
        child
        for child in root.children
        if split_xml_name(child.name) == (OFFICE_NS, "body")
    ]
    if len(bodies) != 1:
        fail("content.xml must contain exactly one office:body")
    presentations = [
        child
        for child in bodies[0].children
        if split_xml_name(child.name) == (OFFICE_NS, "presentation")
    ]
    if len(presentations) != 1:
        fail("content.xml must contain exactly one office:presentation")
    presentation = presentations[0]
    pages = [
        child
        for child in presentation.children
        if split_xml_name(child.name) == (DRAW_NS, "page")
    ]
    if not pages:
        fail("content.xml office:presentation has no direct draw:page")
    if any(page.end is None for page in pages):
        fail("content.xml page has no byte end")

    # Content is retained as ordered events so whitespace/comments between
    # direct pages cannot be mistaken for part of a page run.
    direct_events = list(presentation.content)
    first_page_seen = False
    trailing_seen = False
    leading: list[str] = []
    trailing: list[str] = []
    pages_contiguous = True
    for kind, value in direct_events:
        if kind == "element":
            child = value
            is_page = split_xml_name(child.name) == (DRAW_NS, "page")
            if is_page:
                if trailing_seen:
                    pages_contiguous = False
                first_page_seen = True
            elif first_page_seen:
                trailing_seen = True
                trailing.append(child.name)
            else:
                leading.append(child.name)
        elif kind == "text":
            if not value.isspace():
                if first_page_seen:
                    trailing_seen = True
                    trailing.append("#text")
                else:
                    leading.append("#text")
        elif kind in {"comment", "pi"}:
            if first_page_seen:
                trailing_seen = True
                trailing.append(f"#{kind}")
            else:
                leading.append(f"#{kind}")
    insert_at = int(pages[-1].end)
    if insert_at <= 0 or insert_at > len(raw):
        fail("content.xml last-page byte end is outside the member")
    return ContentFacts(
        raw=raw,
        root=root,
        presentation=presentation,
        pages=pages,
        direct_events=direct_events,
        leading_direct_children=leading,
        trailing_direct_children=trailing,
        pages_contiguous=pages_contiguous,
        insert_at=insert_at,
    )


def attr(node: XmlNode, namespace: str, local: str) -> str | None:
    return node.attrs.get(xml_name(namespace, local))


def plain_paragraph(node: XmlNode) -> str:
    pieces: list[str] = []
    for kind, value in node.content:
        if kind == "text":
            pieces.append(value)
            continue
        if kind != "element":
            fail("generated plain text contains a comment or processing instruction")
        child: XmlNode = value
        namespace, local = split_xml_name(child.name)
        if namespace != TEXT_NS:
            fail(f"generated plain paragraph contains {child.name!r}")
        if local == "s":
            count_text = attr(child, TEXT_NS, "c")
            try:
                count = 1 if count_text is None else int(count_text)
            except ValueError:
                fail(f"generated text:s count is not an integer: {count_text!r}")
            if count < 1 or child.children or any(k != "text" or v.strip() for k, v in child.content):
                fail("generated text:s has invalid children or text")
            pieces.append(" " * count)
        elif local == "tab":
            if child.children or child.content:
                fail("generated text:tab has children")
            pieces.append("\t")
        elif local == "line-break":
            if child.children or child.content:
                fail("generated text:line-break has children")
            pieces.append("\r")
        else:
            fail(f"generated plain paragraph contains unsupported text:{local}")
    return "".join(pieces)


def plain_text_box(node: XmlNode) -> str:
    if split_xml_name(node.name) != (DRAW_NS, "text-box") or node.attrs:
        fail("generated plain frame must contain one attribute-free draw:text-box")
    if len(node.children) != len([item for item in node.content if item[0] == "element"]):
        fail("generated text-box event order is malformed")
    if any(kind != "element" for kind, _ in node.content):
        fail("generated text-box contains direct text or non-element content")
    paragraphs: list[str] = []
    for child in node.children:
        if split_xml_name(child.name) != (TEXT_NS, "p") or child.attrs:
            fail("generated plain text-box must contain attribute-free text:p")
        paragraphs.append(plain_paragraph(child))
    if not paragraphs:
        fail("generated plain text-box has no text:p")
    return "\n".join(paragraphs)


def verify_plain_page(page: XmlNode, title: str, body: str) -> dict[str, Any]:
    if split_xml_name(page.name) != (DRAW_NS, "page"):
        fail("new output node is not draw:page")
    page_name = attr(page, DRAW_NS, "name")
    if page_name is None:
        fail("new draw:page has no draw:name")
    expected_roles = []
    if title:
        expected_roles.append(("title", title, "3.506cm", "0.962cm"))
    if body:
        expected_roles.append(("object", body, "10cm", "5.0cm" if title else "2.0cm"))
    if len(page.children) != len(expected_roles):
        fail(
            "generated draw:page direct frame count differs: "
            f"{len(page.children)} versus {len(expected_roles)}"
        )
    role_facts: list[dict[str, Any]] = []
    for child, (role, expected_text, height, y) in zip(page.children, expected_roles):
        if split_xml_name(child.name) != (DRAW_NS, "frame"):
            fail("generated draw:page has a non-frame direct child")
        expected_attrs = {
            xml_name(PRESENTATION_NS, "class"): role,
            xml_name(SVG_NS, "width"): "25.199cm",
            xml_name(SVG_NS, "height"): height,
            xml_name(SVG_NS, "x"): "1.4cm",
            xml_name(SVG_NS, "y"): y,
        }
        if child.attrs != expected_attrs:
            fail(f"generated {role} frame attributes differ: {child.attrs!r}")
        if len(child.children) != 1:
            fail(f"generated {role} frame must have one draw:text-box")
        actual_text = plain_text_box(child.children[0])
        if actual_text != expected_text:
            fail(
                f"generated {role} text differs: {actual_text!r} versus {expected_text!r}"
            )
        role_facts.append({"role": role, "text_bytes": len(expected_text.encode("utf-8"))})
    # The generated page has no direct comments, text, or processing
    # instructions. Namespace declarations are not exposed as regular attrs by
    # Expat and are therefore checked through the namespace-resolved children.
    if any(kind != "element" for kind, _ in page.content):
        fail("generated draw:page has unexpected direct text or markup")
    return {"name": page_name, "roles": role_facts}


@dataclass
class ZipEntry:
    index: int
    name_raw: bytes
    name: str
    central_raw: bytes
    central: dict[str, int]
    central_extra: bytes
    central_comment: bytes
    local_offset: int
    local_raw: bytes
    local_fixed: dict[str, int]
    local_extra: bytes
    data_start: int
    compressed_payload: bytes
    descriptor: bytes
    local_span: bytes
    decoded: bytes
    is_dir: bool


@dataclass
class ZipArchiveFacts:
    path: str
    raw: bytes
    entries: list[ZipEntry]
    central_comment: bytes
    central_start: int
    central_end: int
    prefix_len: int
    zipfile_names: list[str]

    @property
    def by_name(self) -> dict[bytes, ZipEntry]:
        return {entry.name_raw: entry for entry in self.entries}

    @property
    def central_names(self) -> list[bytes]:
        return [entry.name_raw for entry in self.entries]

    @property
    def local_order(self) -> list[bytes]:
        return [entry.name_raw for entry in sorted(self.entries, key=lambda e: e.local_offset)]


def parse_extra(extra: bytes) -> list[tuple[int, bytes]]:
    fields: list[tuple[int, bytes]] = []
    cursor = 0
    while cursor < len(extra):
        if len(extra) - cursor < 4:
            refuse("malformed ZIP extra field")
        tag, length = struct.unpack_from("<HH", extra, cursor)
        cursor += 4
        end = cursor + length
        if end > len(extra):
            refuse("ZIP extra field extends beyond its record")
        fields.append((tag, extra[cursor:end]))
        cursor = end
    return fields


def decode_zip_name(name_raw: bytes, flags: int) -> str:
    encoding = "utf-8" if flags & 0x800 else "cp437"
    try:
        return name_raw.decode(encoding)
    except UnicodeDecodeError as error:
        refuse(f"ZIP member name is not decodable as {encoding}: {error}")
        raise AssertionError


def central_fields(record: bytes) -> dict[str, int]:
    if len(record) < 46 or record[:4] != b"PK\x01\x02":
        refuse("malformed ZIP central-directory record")
    values = struct.unpack_from("<4s6H3I5H2I", record, 0)
    return {
        "version_made": values[1],
        "version_needed": values[2],
        "flags": values[3],
        "method": values[4],
        "time": values[5],
        "date": values[6],
        "crc32": values[7],
        "compressed_size": values[8],
        "uncompressed_size": values[9],
        "name_len": values[10],
        "extra_len": values[11],
        "comment_len": values[12],
        "disk_start": values[13],
        "internal_attrs": values[14],
        "external_attrs": values[15],
        "local_offset_field": values[16],
    }


def local_fields(record: bytes) -> dict[str, int]:
    if len(record) < 30 or record[:4] != b"PK\x03\x04":
        refuse("malformed ZIP local-file record")
    values = struct.unpack_from("<4s5H3I2H", record, 0)
    return {
        "version_needed": values[1],
        "flags": values[2],
        "method": values[3],
        "time": values[4],
        "date": values[5],
        "crc32": values[6],
        "compressed_size": values[7],
        "uncompressed_size": values[8],
        "name_len": values[9],
        "extra_len": values[10],
    }


def descriptor_length_and_check(
    raw: bytes, start: int, central: dict[str, int], limit: int
) -> int:
    candidates: list[tuple[int, int]] = []
    if raw[start : start + 4] == b"PK\x07\x08":
        candidates.extend(((16, start + 4), (24, start + 4)))
    else:
        candidates.extend(((12, start), (20, start)))
    for length, cursor in candidates:
        if start + length > limit:
            continue
        crc = u32(raw, cursor)
        compressed = u32(raw, cursor + 4)
        uncompressed = u32(raw, cursor + 8)
        if (
            crc == central["crc32"]
            and compressed == central["compressed_size"]
            and uncompressed == central["uncompressed_size"]
        ):
            # ZIP64 descriptors are refused before this branch in normal
            # archives. A 24-byte candidate is retained only when it exactly
            # agrees with a future explicit ZIP64 extension.
            if length == 24:
                refuse("ZIP64 data descriptor is outside this oracle envelope")
            return length
    refuse("ZIP data descriptor does not agree with central metadata")
    raise AssertionError


def manual_decode(method: int, compressed: bytes) -> bytes:
    if method == 0:
        return compressed
    if method == 8:
        try:
            return zlib.decompress(compressed, -15)
        except zlib.error as error:
            fail(f"raw Deflate payload failed independent decode: {error}")
    refuse(f"ZIP compression method {method} is outside this oracle")
    raise AssertionError


def parse_zip(path: Path) -> ZipArchiveFacts:
    raw = path.read_bytes()
    if len(raw) > MAX_INPUT_BYTES:
        refuse(f"archive exceeds oracle input ceiling: {len(raw)} bytes")
    try:
        with zipfile.ZipFile(io.BytesIO(raw), "r") as zip_reader:
            infos = zip_reader.infolist()
            zip_comment = bytes(zip_reader.comment)
            zipfile_names = [info.filename for info in infos]
    except (OSError, zipfile.BadZipFile) as error:
        fail(f"Python zipfile rejected archive: {error}")

    search_start = max(0, len(raw) - (0xFFFF + 22))
    eocd_at = raw.rfind(b"PK\x05\x06", search_start)
    if eocd_at < 0 or eocd_at + 22 > len(raw):
        fail("archive has no complete classic ZIP end record")
    disk, cd_disk, disk_count, total_count, cd_size, cd_offset, comment_len = struct.unpack_from(
        "<4s4H2IH", raw, eocd_at
    )[1:]
    if eocd_at + 22 + comment_len != len(raw):
        refuse("archive has trailing bytes after its EOCD/comment")
    if raw[eocd_at + 22 :] != zip_comment:
        fail("Python zipfile and raw EOCD archive comments differ")
    if any(value == 0xFFFF for value in (disk_count, total_count)) or cd_size == 0xFFFFFFFF or cd_offset == 0xFFFFFFFF:
        refuse("ZIP64 EOCD counters are outside this independent oracle envelope")
    if disk != 0 or cd_disk != 0 or disk_count != total_count:
        refuse("multi-disk ZIP archive is outside this oracle")
    # A ZIP64 locator immediately precedes the classic EOCD in the ordinary
    # ZIP64 layout. This check avoids silently treating such an archive as a
    # prefixed ZIP32 archive.
    if eocd_at >= 20 and raw[eocd_at - 20 : eocd_at - 16] == b"PK\x06\x07":
        refuse("ZIP64 locator is outside this independent oracle envelope")
    central_end = eocd_at
    central_start = central_end - cd_size
    if central_start < 0:
        fail("ZIP central-directory bounds underflow")
    prefix_len = central_start - cd_offset
    if prefix_len < 0:
        refuse("ZIP central-directory prefix cannot be represented")
    if central_start + cd_size != central_end:
        fail("ZIP central-directory size does not meet EOCD")

    central_records: list[tuple[bytes, dict[str, int], bytes, bytes]] = []
    cursor = central_start
    for index in range(total_count):
        if cursor + 46 > central_end or raw[cursor : cursor + 4] != b"PK\x01\x02":
            fail(f"central-directory record {index} is truncated")
        fields = central_fields(raw[cursor : cursor + 46])
        record_len = 46 + fields["name_len"] + fields["extra_len"] + fields["comment_len"]
        end = cursor + record_len
        if end > central_end:
            fail(f"central-directory record {index} exceeds the directory")
        record = raw[cursor:end]
        name_start = 46
        name_end = name_start + fields["name_len"]
        extra_end = name_end + fields["extra_len"]
        name_raw = record[name_start:name_end]
        extra = record[name_end:extra_end]
        comment = record[extra_end:]
        if any(tag == 0x0001 for tag, _ in parse_extra(extra)):
            refuse("ZIP64 extra field is outside this independent oracle envelope")
        central_records.append((record, fields, name_raw, extra + b"\0" + comment))
        cursor = end
    if cursor != central_end:
        refuse("central directory has unparsed trailing bytes")
    if len(central_records) != len(infos):
        fail("raw central entry count differs from Python zipfile")

    entries: list[ZipEntry] = []
    offsets: list[tuple[int, int]] = []
    seen_names: set[bytes] = set()
    for index, ((record, fields, name_raw, extra_comment), info) in enumerate(
        zip(central_records, infos)
    ):
        extra_len = fields["extra_len"]
        name_len = fields["name_len"]
        comment_len = fields["comment_len"]
        name = decode_zip_name(name_raw, fields["flags"])
        if name_raw in seen_names:
            refuse(f"duplicate ZIP member name {name!r}")
        seen_names.add(name_raw)
        if info.filename != name:
            fail(f"Python zipfile name differs for central member {name!r}")
        if info.flag_bits != fields["flags"] or info.compress_type != fields["method"]:
            fail(f"Python zipfile metadata differs for member {name!r}")
        if info.CRC != fields["crc32"] or info.compress_size != fields["compressed_size"] or info.file_size != fields["uncompressed_size"]:
            fail(f"Python zipfile sizes/CRC differ for member {name!r}")
        if fields["flags"] & 1:
            refuse(f"encrypted ZIP member {name!r} is outside this oracle")
        if fields["local_offset_field"] == 0xFFFFFFFF:
            refuse("ZIP64 local-header offset is outside this oracle envelope")
        local_offset = fields["local_offset_field"] + prefix_len
        offsets.append((local_offset, index))
        # Keep the extra and comment independently; the temporary separator is
        # only an internal tuple convenience and is never compared as bytes.
        central_extra = extra_comment[:extra_len]
        central_comment = extra_comment[extra_len + 1 :]
        if len(central_extra) != extra_len or len(central_comment) != comment_len:
            fail(f"central variable fields are malformed for member {name!r}")
        entries.append(
            ZipEntry(
                index=index,
                name_raw=name_raw,
                name=name,
                central_raw=record,
                central=fields,
                central_extra=central_extra,
                central_comment=central_comment,
                local_offset=local_offset,
                local_raw=b"",
                local_fixed={},
                local_extra=b"",
                data_start=0,
                compressed_payload=b"",
                descriptor=b"",
                local_span=b"",
                decoded=b"",
                is_dir=name.endswith("/"),
            )
        )

    offsets.sort()
    if len({offset for offset, _ in offsets}) != len(offsets):
        fail("ZIP members have duplicate local-header offsets")
    local_order = [index for _, index in offsets]
    for position, index in enumerate(local_order):
        entry = entries[index]
        local_start = entry.local_offset
        local_end = offsets[position + 1][0] if position + 1 < len(offsets) else central_start
        if local_start < 0 or local_start + 30 > central_start or local_end > central_start or local_start >= local_end:
            fail(f"local span bounds are invalid for member {entry.name!r}")
        fixed = local_fields(raw[local_start : local_start + 30])
        name_start = local_start + 30
        name_end = name_start + fixed["name_len"]
        extra_end = name_end + fixed["extra_len"]
        if extra_end > local_end:
            fail(f"local header variable fields exceed span for {entry.name!r}")
        local_name = raw[name_start:name_end]
        local_extra = raw[name_end:extra_end]
        if local_name != entry.name_raw:
            refuse(f"local/central member names differ for {entry.name!r}")
        # ZIP permits local and central extra fields to differ. Preserve and
        # compare each raw record independently instead of treating that
        # legal difference as an unsupported archive.
        if any(tag == 0x0001 for tag, _ in parse_extra(local_extra)):
            refuse(f"ZIP64 local extra field for {entry.name!r} is outside this oracle")
        if fixed["flags"] != entry.central["flags"] or fixed["method"] != entry.central["method"]:
            fail(f"local/central method or flags differ for {entry.name!r}")
        if fixed["compressed_size"] == 0xFFFFFFFF or fixed["uncompressed_size"] == 0xFFFFFFFF:
            refuse(f"ZIP64 local size field for {entry.name!r} is outside this oracle")
        data_start = extra_end
        data_end = data_start + entry.central["compressed_size"]
        if data_end > local_end:
            fail(f"compressed payload exceeds local span for {entry.name!r}")
        descriptor = b""
        if fixed["flags"] & 0x08:
            descriptor_len = descriptor_length_and_check(
                raw, data_end, entry.central, local_end
            )
            descriptor = raw[data_end : data_end + descriptor_len]
        else:
            if (
                fixed["crc32"] != entry.central["crc32"]
                or fixed["compressed_size"] != entry.central["compressed_size"]
                or fixed["uncompressed_size"] != entry.central["uncompressed_size"]
            ):
                fail(f"sized local header differs from central metadata for {entry.name!r}")
        compressed_payload = raw[data_start:data_end]
        decoded = manual_decode(entry.central["method"], compressed_payload)
        if len(decoded) != entry.central["uncompressed_size"]:
            fail(f"decoded size disagrees for member {entry.name!r}")
        if zlib.crc32(decoded) & 0xFFFFFFFF != entry.central["crc32"]:
            fail(f"decoded CRC disagrees for member {entry.name!r}")
        try:
            with zipfile.ZipFile(io.BytesIO(raw), "r") as zip_reader:
                # Names are unique above, so reading by the independently
                # decoded name avoids accidentally reusing the final
                # ZipInfo from the central-directory loop.
                zip_decoded = zip_reader.read(entry.name)
        except (OSError, RuntimeError, zipfile.BadZipFile) as error:
            fail(f"Python zipfile could not reopen member {entry.name!r}: {error}")
        if zip_decoded != decoded:
            fail(f"Python zipfile and independent decoder differ for {entry.name!r}")
        entry.local_raw = raw[local_start:extra_end]
        entry.local_fixed = fixed
        entry.local_extra = local_extra
        entry.data_start = data_start
        entry.compressed_payload = compressed_payload
        entry.descriptor = descriptor
        entry.local_span = raw[local_start:local_end]
        entry.decoded = decoded

    return ZipArchiveFacts(
        path=str(path),
        raw=raw,
        entries=entries,
        central_comment=zip_comment,
        central_start=central_start,
        central_end=central_end,
        prefix_len=prefix_len,
        zipfile_names=zipfile_names,
    )


def central_without_offset(entry: ZipEntry) -> bytes:
    masked = bytearray(entry.central_raw)
    masked[42:46] = b"\0\0\0\0"
    return bytes(masked)


def target_central_static(entry: ZipEntry) -> tuple[Any, ...]:
    fields = entry.central
    return (
        entry.name_raw,
        fields["version_made"],
        fields["version_needed"],
        fields["flags"],
        fields["method"],
        fields["time"],
        fields["date"],
        fields["name_len"],
        fields["extra_len"],
        fields["comment_len"],
        fields["disk_start"],
        fields["internal_attrs"],
        fields["external_attrs"],
        entry.central_extra,
        entry.central_comment,
    )


def verify_zip_relation(source: ZipArchiveFacts, output: ZipArchiveFacts) -> dict[str, Any]:
    if source.central_comment != output.central_comment:
        fail("archive comment changed")
    if source.central_names != output.central_names:
        fail("central-directory member order or names changed")
    if source.local_order != output.local_order:
        fail("local member order or names changed")
    source_by_name = source.by_name
    output_by_name = output.by_name
    if CONTENT_NAME not in source_by_name or CONTENT_NAME not in output_by_name:
        fail("source and output must both contain content.xml")
    unchanged: list[str] = []
    target_central_changes: dict[str, Any] | None = None
    for name in source.central_names:
        if name not in output_by_name:
            fail(f"output lost ZIP member {name!r}")
        before = source_by_name[name]
        after = output_by_name[name]
        if name == CONTENT_NAME:
            # The target record necessarily changes CRC and sizes. Its method,
            # name, and archive placement remain independently checked.
            if after.central["method"] != before.central["method"]:
                fail("content.xml compression method changed")
            if after.name_raw != before.name_raw:
                fail("content.xml central name changed")
            if after.is_dir or after.central["flags"] & 1:
                fail("output content.xml became a directory or encrypted member")
            if after.local_fixed["method"] != before.local_fixed["method"]:
                fail("content.xml local compression method changed")
            if after.local_fixed["flags"] & 1:
                fail("output content.xml local header is encrypted")
            target_central_changes = {
                "source_crc32": before.central["crc32"],
                "output_crc32": after.central["crc32"],
                "source_compressed_size": before.central["compressed_size"],
                "output_compressed_size": after.central["compressed_size"],
                "source_uncompressed_size": before.central["uncompressed_size"],
                "output_uncompressed_size": after.central["uncompressed_size"],
                "output_flags": after.central["flags"],
                "output_time": after.central["time"],
                "output_date": after.central["date"],
            }
            continue
        if before.local_span != after.local_span:
            fail(f"untouched member local span changed: {before.name!r}")
        if central_without_offset(before) != central_without_offset(after):
            fail(f"untouched member central record changed: {before.name!r}")
        if before.compressed_payload != after.compressed_payload:
            fail(f"untouched member compressed payload changed: {before.name!r}")
        if before.central["method"] != after.central["method"]:
            fail(f"untouched member compression method changed: {before.name!r}")
        unchanged.append(before.name)

    # The target's generated local header must be internally sized and its
    # central record must agree with the payload decoded above. The producer is
    # allowed to choose fresh timestamp/attribute fields for this regenerated
    # member; all such changes are reported rather than silently attributed to
    # untouched-member preservation.
    target = output_by_name[CONTENT_NAME]
    if target.local_fixed["flags"] & 0x08:
        refuse("generated content.xml data descriptor is outside this oracle")
    if target.local_fixed["crc32"] != target.central["crc32"] or target.local_fixed["compressed_size"] != target.central["compressed_size"] or target.local_fixed["uncompressed_size"] != target.central["uncompressed_size"]:
        fail("output content.xml local and central sizes/CRC disagree")
    if target.local_fixed["time"] != target.central["time"] or target.local_fixed["date"] != target.central["date"]:
        fail("output content.xml local and central timestamps disagree")
    return {
        "central_member_order_verified": True,
        "local_member_order_verified": True,
        "archive_comment_verified": True,
        "untouched_member_count": len(unchanged),
        "untouched_members_raw_local_spans_verified": unchanged,
        "untouched_members_central_records_offset_only_verified": unchanged,
        "untouched_members_compressed_payload_verified": unchanged,
        "target_member": CONTENT_NAME.decode(),
        "target_local_central_self_consistent": True,
        "target_source_compression_method": source_by_name[CONTENT_NAME].central["method"],
        "target_output_compression_method": target.central["method"],
        "target_output_dynamic_sizes": {
            "crc32": target.central["crc32"],
            "compressed_size": target.central["compressed_size"],
            "uncompressed_size": target.central["uncompressed_size"],
        },
        "target_dynamic_central_fields": target_central_changes,
    }


def verify_content(
    source_entry: ZipEntry,
    output_entry: ZipEntry,
    title: str,
    body: str,
    requested_name: str | None,
) -> dict[str, Any]:
    source = parse_content(source_entry.decoded)
    output = parse_content(output_entry.decoded)
    if not source.pages_contiguous:
        fail("source direct draw:page sequence is not contiguous")
    if len(output.pages) != len(source.pages) + 1:
        fail(
            f"output direct page count is {len(output.pages)}, expected {len(source.pages) + 1}"
        )
    source_names = source.page_names
    if len({name for name in source_names if name is not None}) != len(
        [name for name in source_names if name is not None]
    ):
        refuse("source draw:page names are duplicated")
    name = requested_name or f"page{len(source.pages) + 1}"
    if name in source_names:
        fail(f"requested generated page name collides with source: {name!r}")
    matches = [
        page
        for page in output.pages
        if attr(page, DRAW_NS, "name") == name
    ]
    if len(matches) != 1:
        fail(f"output has {len(matches)} direct pages named {name!r}, expected one")
    new_page = matches[0]
    if new_page.end is None:
        fail("output generated page has no byte end")
    if new_page.start != source.insert_at:
        fail(
            f"generated page starts at {new_page.start}, expected source byte offset {source.insert_at}"
        )
    source_prefix = source_entry.decoded[: source.insert_at]
    source_suffix = source_entry.decoded[source.insert_at :]
    output_prefix = output_entry.decoded[: new_page.start]
    output_suffix = output_entry.decoded[new_page.end :]
    if output_prefix != source_prefix:
        fail("output content.xml prefix differs from source prefix")
    if output_suffix != source_suffix:
        fail("output content.xml suffix differs from source suffix")
    page_facts = verify_plain_page(new_page, title, body)
    output_names = output.page_names
    if output_names.count(name) != 1:
        fail("generated page name is not unique in output")
    # Existing direct page names and order are covered by the exact prefix and
    # suffix byte proofs, but keep the explicit semantic list in the result.
    if [candidate for candidate in output_names if candidate != name] != source_names:
        fail("output existing page-name order differs from source")
    return {
        "source_content_xml_sha256": sha256(source_entry.decoded),
        "output_content_xml_sha256": sha256(output_entry.decoded),
        "source_content_xml_bytes": len(source_entry.decoded),
        "output_content_xml_bytes": len(output_entry.decoded),
        "source_page_count": len(source.pages),
        "output_page_count": len(output.pages),
        "source_page_names": source_names,
        "output_page_names": output_names,
        "source_last_page_end_byte": source.insert_at,
        "output_generated_page_start_byte": new_page.start,
        "output_generated_page_end_byte": new_page.end,
        "source_prefix_sha256": sha256(source_prefix),
        "source_suffix_sha256": sha256(source_suffix),
        "generated_page_sha256": sha256(output_entry.decoded[new_page.start : new_page.end]),
        "exact_source_prefix_verified": True,
        "exact_source_suffix_verified": True,
        "one_extra_direct_draw_page_verified": True,
        "generated_page": page_facts,
        "source_leading_direct_children": source.leading_direct_children,
        "source_trailing_direct_children": source.trailing_direct_children,
        "source_pages_contiguous": source.pages_contiguous,
    }


def inventory_binding(source_path: Path, source_raw: bytes) -> dict[str, Any] | None:
    inventory_path = Path(__file__).with_name("staticinventory.json")
    try:
        inventory = json.loads(inventory_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    # native/verify-output.py is six parent steps below the repository root:
    # native/change-0457/results/performance/docs/<repo>.
    root = Path(__file__).resolve().parents[5]
    try:
        relative = source_path.resolve().relative_to(root).as_posix()
    except ValueError:
        return None
    for item in inventory.get("fixtures", []):
        if item.get("path") == relative:
            expected_sha = item.get("archive_sha256")
            expected_bytes = item.get("archive_bytes")
            if expected_sha != sha256(source_raw) or expected_bytes != len(source_raw):
                fail("source archive differs from its static inventory SHA/size pin")
            return {
                "path": relative,
                "sha256": expected_sha,
                "bytes": expected_bytes,
                "static_only": True,
            }
    return None


def verify(
    source_path: Path,
    output_path: Path,
    title: str,
    body: str,
    requested_name: str | None,
) -> dict[str, Any]:
    if not source_path.is_file() or not output_path.is_file():
        fail("source and output must be regular files")
    source_raw = source_path.read_bytes()
    output_raw = output_path.read_bytes()
    source_binding = inventory_binding(source_path, source_raw)
    source = parse_zip(source_path)
    output = parse_zip(output_path)
    zip_facts = verify_zip_relation(source, output)
    content_facts = verify_content(
        source.by_name[CONTENT_NAME],
        output.by_name[CONTENT_NAME],
        title,
        body,
        requested_name,
    )
    return {
        "schema": SCHEMA,
        "status": "validated",
        "static_only_classification": "independent ZIP/XML append oracle; no native application or semantic Rust reopen claim",
        "source": {
            "path": str(source_path),
            "sha256": sha256(source_raw),
            "bytes": len(source_raw),
            "inventory_binding": source_binding,
        },
        "output": {
            "path": str(output_path),
            "sha256": sha256(output_raw),
            "bytes": len(output_raw),
        },
        "request": {
            "title": title,
            "body": body,
            "name": requested_name or f"page{len(parse_content(source.by_name[CONTENT_NAME].decoded).pages) + 1}",
        },
        "zip": zip_facts,
        "content": content_facts,
        "source_archive_sha256": sha256(source_raw),
        "output_archive_sha256": sha256(output_raw),
        "source_hash_verified": True,
    }


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--title", required=True)
    parser.add_argument("--body", required=True)
    parser.add_argument("--name", default=None)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        result = verify(args.source, args.output, args.title, args.body, args.name)
        print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
        return 0
    except OracleRefusal as error:
        result = {
            "schema": SCHEMA,
            "status": "refused",
            "error": {"kind": error.kind, "message": str(error)},
            "source": str(args.source),
            "output": str(args.output),
            "static_only_classification": "oracle refusal; no native application or semantic Rust reopen claim",
        }
        print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
        return 2
    except (OracleError, OSError, ValueError, KeyError) as error:
        result = {
            "schema": SCHEMA,
            "status": "error",
            "error": {"kind": getattr(error, "kind", "validation_error"), "message": str(error)},
            "source": str(args.source),
            "output": str(args.output),
            "static_only_classification": "oracle validation failure; no native application or semantic Rust reopen claim",
        }
        print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
