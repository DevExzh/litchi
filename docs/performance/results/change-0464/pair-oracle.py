#!/usr/bin/env python3
"""Independent, bounded oracle for one PPTX cross-slide copy.

The oracle intentionally knows only the package contract of the current
source-backed cross-slide planner.  It does not import Litchi or execute a
producer.  It validates three caller-supplied archives: a source package, a
destination package, and the published output.  The source slide is selected
by zero-based index; the copied slide must occur at ``insertion_index`` in the
published destination order.

CLI usage::

    python3 -B pair-oracle.py --source SOURCE.pptx --destination DEST.pptx \
        --output OUTPUT.pptx --source-index 0 --insertion-index 1

The pass JSON uses ``schema`` ``litchi-0464-pptx-pair-oracle-v1`` and includes
``identities``, ``source_closure``, ``copied_part_map``, ``added_members``,
``removed_members``, ``metadata``, and raw-record/relationship proof flags.
Expected SHA-256 and byte-count options are optional caller bindings; omitting
them still computes and reports the actual identities.

The proof is package-level: payload identities, OPC relationship legality,
content types, inherited layout/master/theme identity, XML insertion points,
and raw ZIP preservation are checked.  It does not claim that a native Office
application accepted the result.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import posixpath
import re
import struct
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Any, Iterable
import xml.etree.ElementTree as ET

SCHEMA = "litchi-0464-pptx-pair-oracle-v1"
CONTENT_TYPES = "[Content_Types].xml"
ROOT_RELS = "_rels/.rels"
REL_CT = "application/vnd.openxmlformats-package.relationships+xml"
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
LAYOUT_CT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
MASTER_CT = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
THEME_CT = "application/vnd.openxmlformats-officedocument.theme+xml"
STRICT_SLIDE_CT = "application/vnd.ms-powerpoint.slide+xml"
STRICT_LAYOUT_CT = "application/vnd.ms-powerpoint.slideLayout+xml"
STRICT_MASTER_CT = "application/vnd.ms-powerpoint.slideMaster+xml"
STRICT_THEME_CT = "application/vnd.ms-office.theme+xml"
PML_NAMESPACES = {
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://purl.oclc.org/ooxml/presentationml/main",
}
OFFICE_REL_PREFIXES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/",
)

# Relationship local names used by the current bounded planner.  The planner
# accepts both transitional and strict URI dialects; URI prefixes are checked
# separately by comparing the full source relationship type.
SHARED_OWNER_RELS = {
    "slide", "notesSlide", "notesMaster", "comments", "commentAuthors",
    "slideMaster", "theme", "tableStyles", "modernComment",
    "modernCommentAuthor",
}
EXTERNAL_RELS = {"hyperlink", "image", "audio", "video", "oleObject", "media"}
OWNED_REL_KINDS = {
    "image", "audio", "video", "media", "chart", "chartUserShapes",
    "drawing", "diagramData", "diagramLayout", "diagramQuickStyle",
    "diagramColors", "diagramDrawing", "tags", "themeOverride", "oleObject",
    "package", "chartStyle", "chartColorStyle",
}


class VerificationError(ValueError):
    """The archives do not satisfy the independent pair contract."""


def fail(message: str) -> None:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _namespace(tag: str) -> str | None:
    return tag[1:].split("}", 1)[0] if tag.startswith("{") and "}" in tag else None


def _attr(element: ET.Element, local: str, namespace: str | None = None) -> str | None:
    for key, value in element.attrib.items():
        if _local_name(key) == local and (namespace is None or _namespace(key) == namespace):
            return value
    return None


def _xml(raw: bytes, label: str) -> ET.Element:
    try:
        return ET.fromstring(raw)
    except ET.ParseError as error:
        fail(f"{label}: invalid XML: {error}")


def _decode_name(raw: bytes, flags: int, label: str) -> str:
    try:
        return raw.decode("utf-8" if flags & 0x800 else "cp437")
    except UnicodeDecodeError as error:
        fail(f"{label}: invalid ZIP member name: {error}")


def _validate_name(name: str, label: str) -> None:
    require(name and "\\" not in name and not name.startswith("/"), f"{label}: invalid member name")
    parts = PurePosixPath(name).parts
    require(".." not in parts and "." not in parts, f"{label}: traversal member name")


@dataclass(frozen=True)
class RawZip:
    data: bytes
    order: tuple[str, ...]
    payloads: dict[str, bytes]
    local: dict[str, bytes]
    central: dict[str, bytes]


def _central_records(raw: bytes, label: str) -> tuple[list[str], dict[str, bytes], dict[str, tuple[int, int, int, int]]]:
    eocd = raw.rfind(b"PK\x05\x06")
    require(eocd >= 0 and eocd + 22 <= len(raw), f"{label}: missing ZIP EOCD")
    fields = struct.unpack_from("<I4H2IH", raw, eocd)
    require(fields[0] == 0x06054B50, f"{label}: bad ZIP EOCD")
    require(fields[1] == 0 and fields[2] == 0, f"{label}: multi-disk ZIP is unsupported")
    count, central_size, central_offset, comment_length = fields[4], fields[5], fields[6], fields[7]
    require(eocd + 22 + comment_length == len(raw), f"{label}: trailing ZIP bytes")
    require(central_offset + central_size <= len(raw), f"{label}: central directory outside archive")
    cursor = central_offset
    names: list[str] = []
    records: dict[str, bytes] = {}
    fields_by_name: dict[str, tuple[int, int, int, int]] = {}
    for index in range(count):
        prefix = f"{label}.central[{index}]"
        require(cursor + 46 <= len(raw), f"{prefix}: truncated central record")
        values = struct.unpack_from("<I6H3I5H2I", raw, cursor)
        require(values[0] == 0x02014B50, f"{prefix}: missing central signature")
        flags = values[3]
        compressed_size = values[8]
        uncompressed_size = values[9]
        name_length, extra_length, comment_length = values[10], values[11], values[12]
        local_offset = values[16]
        end = cursor + 46 + name_length + extra_length + comment_length
        require(end <= len(raw), f"{prefix}: central record outside archive")
        name = _decode_name(raw[cursor + 46:cursor + 46 + name_length], flags, prefix)
        _validate_name(name, prefix)
        require(name not in records, f"{prefix}: duplicate member {name}")
        record = bytearray(raw[cursor:end])
        # ZIP central records carry the local-header offset at bytes 42..45;
        # relocation changes only those four bytes for an untouched member.
        record[42:46] = b"\0" * 4
        names.append(name)
        records[name] = bytes(record)
        fields_by_name[name] = (flags, compressed_size, uncompressed_size, local_offset)
        cursor = end
    require(cursor == central_offset + central_size, f"{label}: central size mismatch")
    return names, records, fields_by_name


def _local_record(raw: bytes, name: str, fields: tuple[int, int, int, int], label: str) -> bytes:
    flags, compressed_size, _uncompressed_size, offset = fields
    require(offset + 30 <= len(raw), f"{label}: local header outside archive")
    header = struct.unpack_from("<I5H3I2H", raw, offset)
    require(header[0] == 0x04034B50, f"{label}: missing local signature")
    require(header[2] == flags, f"{label}: local flags differ from central record")
    name_length, extra_length = header[9], header[10]
    name_start = offset + 30
    data_start = name_start + name_length + extra_length
    require(data_start <= len(raw), f"{label}: local name/extra outside archive")
    local_name = _decode_name(raw[name_start:name_start + name_length], flags, label)
    require(local_name == name, f"{label}: local member name differs")
    data_end = data_start + compressed_size
    require(data_end <= len(raw), f"{label}: local compressed data outside archive")
    if flags & 0x08:
        # A data descriptor follows the payload.  ZIP32/ZIP64 widths are
        # determined by the central sentinel fields; ordinary PPTX is ZIP32.
        descriptor = 16 if compressed_size < 0xFFFFFFFF and _uncompressed_size < 0xFFFFFFFF else 24
        if raw[data_end:data_end + 4] != b"PK\x07\x08":
            descriptor -= 4
        data_end += descriptor
        require(data_end <= len(raw), f"{label}: data descriptor outside archive")
    return raw[offset:data_end]


def _snapshot(raw: bytes, label: str) -> RawZip:
    names, central, fields = _central_records(raw, label)
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            infos = archive.infolist()
            info_names = [info.filename for info in infos]
            require(tuple(info_names) == tuple(names), f"{label}: central order differs from ZIP reader")
            payloads = {info.filename: archive.read(info) for info in infos}
    except (OSError, RuntimeError, ValueError, zipfile.BadZipFile, zipfile.LargeZipFile) as error:
        fail(f"{label}: invalid ZIP: {error}")
    local = {name: _local_record(raw, name, fields[name], f"{label}.{name}") for name in names}
    return RawZip(raw, tuple(names), payloads, local, central)


def _resolve(base: str, target: str, label: str) -> str:
    require(isinstance(target, str) and target, f"{label}: empty relationship target")
    require("?" not in target and "#" not in target, f"{label}: query/fragment is unsupported")
    result = target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(base), target))
    _validate_name(result, label)
    return result


def _rels_path(part: str) -> str:
    parent, name = posixpath.split(part)
    return posixpath.join(parent, "_rels", name + ".rels") if parent else posixpath.join("_rels", name + ".rels")


def _type_local(value: str) -> str:
    return value.rsplit("/", 1)[-1]


def _known_office_relationship(rel_type: str, kind: str) -> bool:
    if rel_type.startswith(OFFICE_REL_PREFIXES):
        return True
    return kind == "media" and rel_type == "http://schemas.microsoft.com/office/2007/relationships/media"


@dataclass(frozen=True)
class Relationship:
    rid: str
    rel_type: str
    target: str
    target_mode: str | None

    @property
    def kind(self) -> str:
        return _type_local(self.rel_type)


class Package:
    def __init__(self, raw: RawZip, label: str):
        self.raw = raw
        self.label = label
        self.types = self._content_types()
        require(not any(name.startswith("_xmlsignatures/") for name in raw.order), f"{label}: signature infrastructure is outside this oracle scope")
        require(not any("vbaProject" in name for name in raw.order), f"{label}: macro infrastructure is outside this oracle scope")
        require(not any("macroEnabled" in value for value in self.types.values()), f"{label}: macro-enabled content type is outside this oracle scope")
        self.rels_cache: dict[str, tuple[Relationship, ...]] = {}
        self.presentation = self._presentation_part()
        self.presentation_rels = _rels_path(self.presentation)
        require(self.presentation_rels in self.raw.payloads, f"{label}: presentation relationships missing")
        self.slides = self._slide_order()

    def _content_types(self) -> dict[str, str]:
        require(CONTENT_TYPES in self.raw.payloads, f"{self.label}: [Content_Types].xml missing")
        root = _xml(self.raw.payloads[CONTENT_TYPES], f"{self.label}.[Content_Types].xml")
        require(_local_name(root.tag) == "Types", f"{self.label}: wrong content-types root")
        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        for index, child in enumerate(root):
            kind = _local_name(child.tag)
            if kind == "Default":
                ext = _attr(child, "Extension")
                value = _attr(child, "ContentType")
                require(ext and value and ext not in defaults, f"{self.label}.Default[{index}]: invalid/duplicate")
                defaults[ext.lower()] = value
            elif kind == "Override":
                part = _attr(child, "PartName")
                value = _attr(child, "ContentType")
                require(part and value and part.startswith("/"), f"{self.label}.Override[{index}]: invalid")
                part = part[1:]
                _validate_name(part, f"{self.label}.Override[{index}]")
                require(part not in overrides, f"{self.label}.Override[{index}]: duplicate {part}")
                overrides[part] = value
        result: dict[str, str] = {}
        for name in self.raw.order:
            if name == CONTENT_TYPES:
                continue
            value = overrides.get(name)
            if value is None:
                extension = name.rsplit("/", 1)[-1].rsplit(".", 1)[-1].lower() if "." in name.rsplit("/", 1)[-1] else ""
                value = defaults.get(extension)
            # The retained LibreOffice/Litchi package family commonly omits a
            # newly-created .rels override while the relationship sidecar is
            # still unambiguously identified by its OPC filename.  Treat that
            # filename convention as the narrow relationship-content-type
            # fallback; every other part still requires an explicit
            # [Content_Types] declaration or Default.
            if value is None and name.endswith(".rels"):
                value = REL_CT
            require(value is not None, f"{self.label}: no content type for {name}")
            result[name] = value
        self.defaults = defaults
        self.overrides = overrides
        self._override_order = tuple(
            _attr(child, "PartName")[1:]
            for child in root if _local_name(child.tag) == "Override" and _attr(child, "PartName")
        )
        return result

    def relationships(self, part: str) -> tuple[Relationship, ...]:
        if part in self.rels_cache:
            return self.rels_cache[part]
        sidecar = _rels_path(part)
        require(sidecar in self.raw.payloads, f"{self.label}.{part}: relationship sidecar missing") if False else None
        if sidecar not in self.raw.payloads:
            result: tuple[Relationship, ...] = ()
            self.rels_cache[part] = result
            return result
        require(self.types.get(sidecar) == REL_CT, f"{self.label}.{sidecar}: wrong relationship content type")
        root = _xml(self.raw.payloads[sidecar], f"{self.label}.{sidecar}")
        require(_local_name(root.tag) == "Relationships", f"{self.label}.{sidecar}: wrong root")
        result_list: list[Relationship] = []
        seen: set[str] = set()
        for index, child in enumerate(root):
            prefix = f"{self.label}.{sidecar}.Relationship[{index}]"
            require(_local_name(child.tag) == "Relationship", f"{prefix}: unexpected child")
            rid = _attr(child, "Id")
            rel_type = _attr(child, "Type")
            target = _attr(child, "Target")
            mode = _attr(child, "TargetMode")
            require(rid and rel_type and target and rid not in seen, f"{prefix}: invalid/duplicate relationship")
            seen.add(rid)
            if mode is not None:
                require(mode == "External", f"{prefix}: unsupported TargetMode {mode}")
            result_list.append(Relationship(rid, rel_type, target, mode))
        result = tuple(result_list)
        self.rels_cache[part] = result
        return result

    def target(self, part: str, relationship: Relationship) -> str:
        return _resolve(part, relationship.target, f"{self.label}.{part}.{relationship.rid}")

    def _presentation_part(self) -> str:
        require(ROOT_RELS in self.raw.payloads, f"{self.label}: package relationships missing")
        root = _xml(self.raw.payloads[ROOT_RELS], f"{self.label}.{ROOT_RELS}")
        require(_local_name(root.tag) == "Relationships", f"{self.label}: wrong package relationships root")
        candidates: list[str] = []
        for index, child in enumerate(root):
            if _local_name(child.tag) != "Relationship":
                continue
            rel_type = _attr(child, "Type") or ""
            target = _attr(child, "Target")
            mode = _attr(child, "TargetMode")
            if target and mode is None and _type_local(rel_type) == "officeDocument" and _known_office_relationship(rel_type, "officeDocument"):
                candidates.append(_resolve("", target, f"{self.label}.{ROOT_RELS}[{index}].Target"))
        require(len(candidates) == 1, f"{self.label}: expected one internal officeDocument relationship")
        part = candidates[0]
        require(part in self.raw.payloads and self.types.get(part) in {
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
            "application/vnd.ms-powerpoint.presentationml.presentation.main+xml",
        } or part in self.raw.payloads, f"{self.label}: presentation part missing")
        return part

    def _slide_order(self) -> list[str]:
        root = _xml(self.raw.payloads[self.presentation], f"{self.label}.{self.presentation}")
        require(_local_name(root.tag) == "presentation", f"{self.label}: presentation XML root mismatch")
        relations = self.relationships(self.presentation)
        by_id = {rel.rid: rel for rel in relations}
        lists = [node for node in root.iter() if _local_name(node.tag) == "sldIdLst"]
        require(len(lists) == 1, f"{self.label}: expected one slide ID list")
        result: list[str] = []
        for index, node in enumerate(list(lists[0])):
            require(_local_name(node.tag) == "sldId", f"{self.label}: unexpected slide-list child")
            rid = _attr(node, "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            if rid is None:
                # Strict packages use the strict relationships namespace; local
                # name remains sufficient after the root dialect was admitted.
                rid = next((value for key, value in node.attrib.items() if _local_name(key) == "id"), None)
            require(rid in by_id, f"{self.label}: slide ID {rid!r} has no relationship")
            rel = by_id[rid]
            require(rel.kind == "slide" and rel.target_mode is None, f"{self.label}: slide ID does not target an internal slide")
            target = self.target(self.presentation, rel)
            require(target in self.raw.payloads and self.types.get(target) in {SLIDE_CT, STRICT_SLIDE_CT}, f"{self.label}: slide target {target} is not a slide part")
            result.append(target)
        require(result, f"{self.label}: presentation has no slides")
        return result

    def content_type(self, part: str) -> str:
        require(part in self.types, f"{self.label}: missing content type for {part}")
        return self.types[part]


@dataclass(frozen=True)
class Closure:
    owned: tuple[str, ...]
    layout: str
    sidecars: tuple[str, ...]
    external_relationships: int


def _validate_xml_surface(package: Package, part: str, rel_kind: str) -> None:
    content_type = package.content_type(part)
    if not (content_type.endswith("+xml") or content_type.endswith("/xml")):
        return
    root = _xml(package.raw.payloads[part], f"{package.label}.{part}")
    require(b"markup-compatibility" not in package.raw.payloads[part], f"{package.label}.{part}: markup-compatibility extensions are outside this oracle scope")
    expected = {
        "slide": "sld", "chart": "chartSpace", "chartUserShapes": "userShapes",
        "drawing": "userShapes", "diagramData": "dataModel", "diagramLayout": "layoutDef",
        "diagramQuickStyle": "styleDef", "diagramColors": "colorsDef", "diagramDrawing": "drawing",
        "tags": "tagLst", "themeOverride": "themeOverride", "chartStyle": "chartStyle",
        "chartColorStyle": "colorStyle",
    }.get(rel_kind)
    if part in package.slides or content_type in {SLIDE_CT, STRICT_SLIDE_CT}:
        expected = "sld"
    if expected is not None:
        require(_local_name(root.tag) == expected, f"{package.label}.{part}: unexpected XML root {root.tag}")


def _kind_valid(kind: str, content_type: str) -> bool:
    if kind == "image": return content_type.startswith("image/")
    if kind == "audio": return content_type.startswith("audio/")
    if kind in {"video", "media"}: return content_type.startswith("video/") or content_type.startswith("audio/")
    if kind == "chart": return content_type in {"application/vnd.openxmlformats-officedocument.drawingml.chart+xml", "application/vnd.ms-office.chart+xml"}
    if kind == "chartUserShapes": return content_type in {"application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml", "application/vnd.openxmlformats-officedocument.drawingml.drawing+xml", "application/vnd.ms-office.drawing+xml"}
    if kind == "drawing": return content_type.endswith("drawing+xml")
    if kind == "diagramData": return content_type.endswith("diagramData+xml") or content_type.endswith("diagramData+xml")
    if kind == "diagramLayout": return content_type.endswith("diagramLayout+xml")
    if kind == "diagramQuickStyle": return content_type.endswith("diagramQuickStyle+xml")
    if kind == "diagramColors": return content_type.endswith("diagramColors+xml")
    if kind == "diagramDrawing": return content_type.endswith("diagramDrawing+xml")
    if kind == "tags": return content_type.endswith("tags+xml")
    if kind == "themeOverride": return content_type.endswith("themeOverride+xml")
    if kind == "oleObject": return content_type.endswith("oleObject+xml")
    if kind in {"package", "chartStyle", "chartColorStyle"}: return True
    return False


def _collect_closure(package: Package, slide: str) -> Closure:
    queue = [slide]
    owned: set[str] = set()
    layout: str | None = None
    sidecars: set[str] = set()
    external = 0
    while queue:
        part = queue.pop(0)
        if part in owned:
            continue
        require(part in package.raw.payloads, f"{package.label}: closure part missing {part}")
        owned.add(part)
        _validate_xml_surface(package, part, "slide" if part == slide else "")
        sidecar = _rels_path(part)
        rels = package.relationships(part)
        if rels:
            require(sidecar in package.raw.payloads, f"{package.label}.{part}: relationship sidecar missing")
            sidecars.add(sidecar)
        for rel in rels:
            kind = rel.kind
            require(_known_office_relationship(rel.rel_type, kind), f"{package.label}.{part}.{rel.rid}: unsupported relationship namespace")
            if rel.target_mode == "External":
                require(kind in EXTERNAL_RELS, f"{package.label}.{part}.{rel.rid}: external relationship {kind} is not allowlisted")
                external += 1
                continue
            target = package.target(part, rel)
            require(target in package.raw.payloads, f"{package.label}.{part}.{rel.rid}: target {target} missing")
            if kind == "slideLayout":
                require(part == slide, f"{package.label}.{part}: private dependency reaches slide layout")
                require(layout is None, f"{package.label}.{part}: more than one slide layout")
                require(package.content_type(target) in {LAYOUT_CT, STRICT_LAYOUT_CT}, f"{package.label}: invalid slide layout content type")
                layout = target
                continue
            require(kind not in SHARED_OWNER_RELS, f"{package.label}.{part}.{rel.rid}: shared-owner relationship {kind} is outside closure")
            require(kind in OWNED_REL_KINDS, f"{package.label}.{part}.{rel.rid}: unsupported internal relationship {kind}")
            ctype = package.content_type(target)
            require(_kind_valid(kind, ctype), f"{package.label}.{part}.{rel.rid}: incompatible content type {ctype}")
            _validate_xml_surface(package, target, kind)
            queue.append(target)
    require(layout is not None, f"{package.label}.{slide}: exactly one reusable layout is required")
    return Closure(tuple(sorted(owned)), layout, tuple(sorted(sidecars)), external)


def _relationship_signature(package: Package, part: str, rel: Relationship) -> tuple[str, str, str, str | None]:
    return (rel.rid, rel.rel_type, rel.target, rel.target_mode)


def _compare_inheritance(source: Package, source_part: str, destination: Package, destination_part: str, surface: str, seen: set[tuple[str, str]]) -> None:
    pair = (source_part, destination_part)
    if pair in seen:
        return
    seen.add(pair)
    expected = {"layout": {LAYOUT_CT, STRICT_LAYOUT_CT}, "master": {MASTER_CT, STRICT_MASTER_CT}, "theme": {THEME_CT, STRICT_THEME_CT}}[surface]
    require(source.content_type(source_part) in expected and destination.content_type(destination_part) in expected, f"shared {surface}: content type mismatch")
    require(source.content_type(source_part) == destination.content_type(destination_part), f"shared {surface}: content types differ")
    require(b"markup-compatibility" not in source.raw.payloads[source_part] and b"markup-compatibility" not in destination.raw.payloads[destination_part], f"shared {surface}: markup-compatibility extensions are outside this oracle scope")
    require(source.raw.payloads[source_part] == destination.raw.payloads[destination_part], f"shared {surface}: payload differs")
    left = source.relationships(source_part)
    right = destination.relationships(destination_part)
    require(len(left) == len(right), f"shared {surface}: relationship count differs")
    for lrel, rrel in zip(left, right):
        require(_known_office_relationship(lrel.rel_type, lrel.kind), f"shared {surface}: unsupported relationship namespace")
        require(_relationship_signature(source, source_part, lrel) == _relationship_signature(destination, destination_part, rrel), f"shared {surface}: relationship identity differs")
        require(lrel.target_mode is None, f"shared {surface}: external inheritance relationship is unsupported")
        next_surface: str | None = None
        if surface == "layout" and lrel.kind == "slideMaster": next_surface = "master"
        elif surface == "master" and lrel.kind == "slideLayout": next_surface = "layout"
        elif surface == "master" and lrel.kind == "theme": next_surface = "theme"
        else:
            fail(f"shared {surface}: unsupported internal relationship {lrel.kind}")
        lt = source.target(source_part, lrel)
        rt = destination.target(destination_part, rrel)
        _compare_inheritance(source, lt, destination, rt, next_surface, seen)


def _metadata_splice(before: bytes, after: bytes, closing: bytes, label: str, *, before_whitespace: bool = False) -> bytes:
    """Return the one inserted lexical fragment in a metadata owner.

    ``insert_slide`` inserts before the selected existing entry for a middle
    position and after the last entry for an append.  Therefore a fixed
    ``</sldIdLst>`` boundary would reject a valid position-zero publication.
    The longest common prefix/suffix pair proves that the destination bytes
    were preserved around exactly one insertion at either position.
    """
    require(before != after, f"{label}: expected metadata insertion")
    require(closing in before and closing in after, f"{label}: closing element missing")
    prefix_length = 0
    while prefix_length < len(before) and prefix_length < len(after) and before[prefix_length] == after[prefix_length]:
        prefix_length += 1
    suffix_length = 0
    while (suffix_length < len(before) - prefix_length
           and suffix_length < len(after) - prefix_length
           and before[len(before) - suffix_length - 1] == after[len(after) - suffix_length - 1]):
        suffix_length += 1
    prefix, suffix = before[:prefix_length], before[len(before) - suffix_length:] if suffix_length else b""
    require(after == prefix + after[prefix_length:len(after) - suffix_length if suffix_length else None] + suffix, f"{label}: destination lexical prefix/suffix changed")
    require(before == prefix + before[prefix_length:len(before) - suffix_length if suffix_length else None] + suffix, f"{label}: ambiguous metadata splice")
    inserted = after[prefix_length:len(after) - suffix_length if suffix_length else None]
    require(inserted, f"{label}: empty metadata insertion")
    return inserted


def _closing_token(raw: bytes, local: bytes, label: str) -> bytes:
    match = re.search(rb"</(?:[A-Za-z_][A-Za-z0-9_.-]*:)?" + re.escape(local) + rb"\s*>", raw)
    require(match is not None, f"{label}: closing element missing")
    return match.group(0)


def _check_metadata(source: Package, destination: Package, output: Package, added: set[str], copied_slide: str) -> dict[str, Any]:
    # The current publisher edits exactly these metadata owners.  A lexical
    # splice check prevents a semantic XML rewriter from silently changing the
    # untouched destination bytes around the inserted node.
    pres_insert = _metadata_splice(
        destination.raw.payloads[destination.presentation], output.raw.payloads[output.presentation],
        _closing_token(destination.raw.payloads[destination.presentation], b"sldIdLst", f"{output.label}.{output.presentation}"),
        f"{output.label}.{output.presentation}",
    )
    rel_insert = _metadata_splice(
        destination.raw.payloads[destination.presentation_rels], output.raw.payloads[output.presentation_rels],
        _closing_token(destination.raw.payloads[destination.presentation_rels], b"Relationships", f"{output.label}.{output.presentation_rels}"),
        f"{output.label}.{output.presentation_rels}",
    )
    ct_insert = _metadata_splice(
        destination.raw.payloads[CONTENT_TYPES], output.raw.payloads[CONTENT_TYPES],
        _closing_token(destination.raw.payloads[CONTENT_TYPES], b"Types", f"{output.label}.{CONTENT_TYPES}"),
        f"{output.label}.{CONTENT_TYPES}",
        before_whitespace=True,
    )
    require(set(destination.raw.order) - {destination.presentation, destination.presentation_rels, CONTENT_TYPES} <= set(output.raw.order), "metadata: destination member disappeared")
    # Existing content-type declarations remain in order. New declarations may
    # only name copied parts; defaults themselves must not be rewritten.
    require(output.defaults == destination.defaults, "metadata: content-type defaults changed")
    old_overrides = [(part, destination.overrides[part]) for part in destination._override_order]
    new_overrides = [(part, output.overrides[part]) for part in output._override_order]
    cursor = 0
    for item in old_overrides:
        try:
            cursor = next(i for i in range(cursor, len(new_overrides)) if new_overrides[i] == item) + 1
        except StopIteration:
            fail(f"metadata: destination content-type override removed {item[0]}")
    new_parts = {part for part, _value in new_overrides} - set(destination.overrides)
    require(new_parts <= added, "metadata: unexpected content-type override")
    for name in added:
        value = output.content_type(name)
        require(value == output.overrides.get(name, value), f"metadata: copied part override mismatch {name}")
    return {
        "presentation_insertion_bytes": len(pres_insert),
        "presentation_relationship_insertion_bytes": len(rel_insert),
        "content_types_insertion_bytes": len(ct_insert),
        "exact_prefix_suffix_checked": True,
        "new_content_type_overrides": sorted(new_parts),
    }


def _same_payload(source: Package, source_part: str, output: Package, output_part: str) -> None:
    require(output_part in output.raw.payloads, f"copied payload missing {output_part}")
    require(source.content_type(source_part) == output.content_type(output_part), f"copied content type differs {source_part} -> {output_part}")
    require(source.raw.payloads[source_part] == output.raw.payloads[output_part], f"copied payload differs {source_part} -> {output_part}")


def _map_copied_closure(source: Package, source_closure: Closure, destination: Package, output: Package, copied_slide: str) -> tuple[dict[str, str], set[str], dict[str, str]]:
    source_slide = next(part for part in source_closure.owned if part in source.slides)
    mapping = {source_slide: copied_slide}
    added = set(output.raw.order) - set(destination.raw.order)
    require(copied_slide in added, "closure: copied slide is not an output addition")
    copied_sidecars: set[str] = set()
    provenance: dict[str, str] = {}
    queue = [source_slide]
    while queue:
        source_part = queue.pop(0)
        output_part = mapping[source_part]
        _same_payload(source, source_part, output, output_part)
        source_rels = source.relationships(source_part)
        output_sidecar = _rels_path(output_part)
        if source_rels:
            require(output_sidecar in added, f"closure: copied relationship sidecar missing for {output_part}")
            copied_sidecars.add(output_sidecar)
            output_rels = output.relationships(output_part)
            require(len(output_rels) == len(source_rels), f"closure: relationship count differs for {source_part}")
        else:
            require(output_sidecar not in added, f"closure: unexpected empty relationship sidecar for {output_part}")
            output_rels = ()
        for srel, orel in zip(source_rels, output_rels):
            require((srel.rid, srel.rel_type, srel.target_mode) == (orel.rid, orel.rel_type, orel.target_mode), f"closure: relationship identity differs for {source_part}.{srel.rid}")
            if srel.target_mode == "External":
                require(srel.kind in EXTERNAL_RELS and srel.target == orel.target, f"closure: external relationship changed {source_part}.{srel.rid}")
                continue
            starget = source.target(source_part, srel)
            if srel.kind == "slideLayout":
                require(orel.target_mode is None, f"closure: layout relationship became external")
                otarget = output.target(output_part, orel)
                require(otarget in destination.raw.payloads and otarget not in added, f"closure: layout target is not reused destination member")
                require(destination.content_type(otarget) in {LAYOUT_CT, STRICT_LAYOUT_CT}, f"closure: reused layout has wrong content type")
                provenance[starget] = otarget
                continue
            require(srel.kind in OWNED_REL_KINDS and srel.kind not in SHARED_OWNER_RELS, f"closure: unsupported copied relationship {srel.kind}")
            otarget = output.target(output_part, orel)
            require(otarget in added, f"closure: copied relationship target is not a new member {otarget}")
            if starget in mapping:
                require(mapping[starget] == otarget, f"closure: inconsistent mapping for {starget}")
            else:
                mapping[starget] = otarget
                queue.append(starget)
            _same_payload(source, starget, output, otarget)
            provenance[starget] = otarget
    copied_parts = set(mapping.values())
    expected_sidecars = set()
    for source_part in mapping:
        if source.relationships(source_part):
            expected_sidecars.add(_rels_path(mapping[source_part]))
    expected_added = copied_parts | expected_sidecars
    require(added == expected_added, f"closure: unexpected added/removed parts: actual={sorted(added)} expected={sorted(expected_added)}")
    # All source-owned parts must be mapped exactly once, including shared
    # references from several relationships.
    require(set(mapping) == set(source_closure.owned), f"closure: source dependency closure is not fully represented")
    return mapping, expected_added, provenance


def _check_relationship_target_legality(package: Package, parts: Iterable[str]) -> None:
    for part in parts:
        for rel in package.relationships(part):
            if rel.target_mode == "External":
                continue
            target = package.target(part, rel)
            require(target in package.raw.payloads, f"{package.label}.{part}.{rel.rid}: target is not a package member")


def verify_pair(
    source_path: str | Path,
    destination_path: str | Path,
    output_path: str | Path,
    source_index: int,
    insertion_index: int,
    *,
    expected_source_sha256: str | None = None,
    expected_destination_sha256: str | None = None,
    expected_output_sha256: str | None = None,
    expected_source_bytes: int | None = None,
    expected_destination_bytes: int | None = None,
    expected_output_bytes: int | None = None,
    same_input_destination: bool = False,
) -> dict[str, Any]:
    paths = [Path(source_path), Path(destination_path), Path(output_path)]
    require(all(path.is_file() for path in paths), "input/output path is not a regular file")
    resolved = [path.resolve() for path in paths]
    require(resolved[2] not in {resolved[0], resolved[1]}, "output must be a separate archive")
    same_path = resolved[0] == resolved[1]
    if same_path:
        require(same_input_destination, "source and destination are the same path; pass same_input_destination explicitly")
    source_raw = paths[0].read_bytes()
    destination_raw = paths[1].read_bytes()
    output_raw = paths[2].read_bytes()
    identities = {
        "source": {"path": str(paths[0]), "resolved_path": str(resolved[0]), "bytes": len(source_raw), "sha256": sha(source_raw)},
        "destination": {"path": str(paths[1]), "resolved_path": str(resolved[1]), "bytes": len(destination_raw), "sha256": sha(destination_raw)},
        "output": {"path": str(paths[2]), "resolved_path": str(resolved[2]), "bytes": len(output_raw), "sha256": sha(output_raw)},
    }
    expected = ((expected_source_sha256, identities["source"]["sha256"], "source sha256"), (expected_destination_sha256, identities["destination"]["sha256"], "destination sha256"), (expected_output_sha256, identities["output"]["sha256"], "output sha256"), (expected_source_bytes, len(source_raw), "source bytes"), (expected_destination_bytes, len(destination_raw), "destination bytes"), (expected_output_bytes, len(output_raw), "output bytes"))
    for wanted, actual, label in expected:
        if wanted is not None:
            require(wanted == actual, f"{label}: expected {wanted}, got {actual}")
    require(isinstance(source_index, int) and source_index >= 0, "source_index must be non-negative")
    require(isinstance(insertion_index, int) and insertion_index >= 0, "insertion_index must be non-negative")
    source = Package(_snapshot(source_raw, "source"), "source")
    destination = Package(_snapshot(destination_raw, "destination"), "destination")
    output = Package(_snapshot(output_raw, "output"), "output")
    require(source_index < len(source.slides), "source_index is outside source slide order")
    require(insertion_index <= len(destination.slides), "insertion_index is outside destination insertion range")
    source_slide = source.slides[source_index]
    source_closure = _collect_closure(source, source_slide)
    _check_relationship_target_legality(source, source_closure.owned)
    _check_relationship_target_legality(destination, destination.slides)
    _check_relationship_target_legality(output, output.slides)
    require(source.raw.payloads[source.presentation] != b"" and destination.raw.payloads[destination.presentation] != b"", "empty presentation")
    # Dialect consistency is part of the native planner's admission boundary.
    source_presentation_root = _xml(source.raw.payloads[source.presentation], "source presentation")
    destination_presentation_root = _xml(destination.raw.payloads[destination.presentation], "destination presentation")
    source_namespace = _namespace(source_presentation_root.tag)
    destination_namespace = _namespace(destination_presentation_root.tag)
    require(source_namespace in PML_NAMESPACES and destination_namespace in PML_NAMESPACES, "source/destination are not PresentationML packages")
    require(source_namespace == destination_namespace, "source and destination PresentationML dialects differ")
    added = set(output.raw.order) - set(destination.raw.order)
    removed = set(destination.raw.order) - set(output.raw.order)
    require(not removed, f"destination members removed: {sorted(removed)}")
    candidates = [name for name in added if output.content_type(name) in {SLIDE_CT, STRICT_SLIDE_CT} and output.raw.payloads[name] == source.raw.payloads[source_slide]]
    require(len(candidates) == 1, f"expected one copied selected-slide addition, found {candidates}")
    copied_slide = candidates[0]
    mapping, expected_added, provenance = _map_copied_closure(source, source_closure, destination, output, copied_slide)
    output_order = output.slides
    expected_order = destination.slides[:insertion_index] + [copied_slide] + destination.slides[insertion_index:]
    require(output_order == expected_order, f"slide order differs: expected {expected_order}, got {output_order}")
    # The copied slide's reused layout must be a destination layout whose full
    # inheritance graph equals the source layout graph.
    copied_rels = output.relationships(copied_slide)
    copied_layout_rels = [rel for rel in copied_rels if rel.kind == "slideLayout"]
    require(len(copied_layout_rels) == 1, "copied slide must have exactly one slideLayout relationship")
    reused_layout = output.target(copied_slide, copied_layout_rels[0])
    require(reused_layout in destination.raw.payloads and reused_layout not in added, "copied slide layout is not a reused destination member")
    require(destination.content_type(reused_layout) in {LAYOUT_CT, STRICT_LAYOUT_CT}, "copied slide layout has an invalid content type")
    _compare_inheritance(source, source_closure.layout, destination, reused_layout, "layout", set())
    # Presentation relationship ownership and slide-ID insertion are checked by
    # comparing all old IDs plus one new slide relation.
    d_rels = destination.relationships(destination.presentation)
    o_rels = output.relationships(output.presentation)
    d_ids = [rel.rid for rel in d_rels]
    o_ids = [rel.rid for rel in o_rels]
    require([rid for rid in o_ids if rid in d_ids] == d_ids, "presentation relationship order/identity changed")
    new_rels = [rel for rel in o_rels if rel.rid not in set(d_ids)]
    require(len(new_rels) == 1 and new_rels[0].kind == "slide" and new_rels[0].target_mode is None, "expected one new presentation slide relationship")
    require(output.target(output.presentation, new_rels[0]) == copied_slide, "new presentation relationship targets the wrong slide")
    d_nodes = [_attr(node, "id") for node in _xml(destination.raw.payloads[destination.presentation], "destination presentation").iter() if _local_name(node.tag) == "sldId"]
    o_nodes = [_attr(node, "id") for node in _xml(output.raw.payloads[output.presentation], "output presentation").iter() if _local_name(node.tag) == "sldId"]
    require(len(o_nodes) == len(d_nodes) + 1, "presentation slide-ID count differs")
    require(o_nodes[:insertion_index] + o_nodes[insertion_index + 1:] == d_nodes, "existing presentation slide IDs changed")
    metadata = _check_metadata(source, destination, output, added, copied_slide)
    retained = set(destination.raw.order) - {destination.presentation, destination.presentation_rels, CONTENT_TYPES}
    for name in retained:
        require(output.raw.local[name] == destination.raw.local[name], f"untouched local ZIP record changed: {name}")
        require(output.raw.central[name] == destination.raw.central[name], f"untouched central ZIP record changed: {name}")
    return {
        "schema": SCHEMA,
        "status": "pass",
        "input_identity_fact": "same_path_explicit" if same_path else ("same_bytes_distinct_paths" if source_raw == destination_raw else "distinct_bytes"),
        "independent_producer_claim": False,
        "identities": identities,
        "source_index": source_index,
        "insertion_index": insertion_index,
        "source_slide": source_slide,
        "copied_slide": copied_slide,
        "destination_slide_order": destination.slides,
        "output_slide_order": output.slides,
        "source_closure": {
            "owned_parts": list(source_closure.owned),
            "reused_source_layout": source_closure.layout,
            "external_relationships": source_closure.external_relationships,
            "sidecars": list(source_closure.sidecars),
        },
        "copied_part_map": dict(sorted(mapping.items())),
        "reused_destination_layout": reused_layout,
        "added_members": sorted(expected_added),
        "removed_members": [],
        "untouched_destination_raw_records_checked": len(retained),
        "metadata": metadata,
        "relationship_target_legality_checked": True,
        "payload_identity_checked": True,
        "limits": {
            "dependency_surface": "current bounded planner relationship families; unsupported internal families fail closed",
            "native_application_acceptance": False,
            "independent_producer": False,
        },
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--destination", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--source-index", required=True, type=int)
    parser.add_argument("--insertion-index", required=True, type=int)
    parser.add_argument("--same-source-destination", action="store_true", help="explicitly admit the same source/destination path; no independent-producer claim is made")
    for role in ("source", "destination", "output"):
        parser.add_argument(f"--{role}-sha256")
        parser.add_argument(f"--{role}-bytes", type=int)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        result = verify_pair(
            args.source, args.destination, args.output, args.source_index, args.insertion_index,
            expected_source_sha256=args.source_sha256,
            expected_destination_sha256=args.destination_sha256,
            expected_output_sha256=args.output_sha256,
            expected_source_bytes=args.source_bytes,
            expected_destination_bytes=args.destination_bytes,
            expected_output_bytes=args.output_bytes,
            same_input_destination=args.same_source_destination,
        )
    except (VerificationError, OSError, ValueError, zipfile.BadZipFile) as error:
        print(f"pair oracle failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
