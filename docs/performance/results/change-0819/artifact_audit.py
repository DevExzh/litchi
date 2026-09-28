#!/usr/bin/env python3
"""Independent oracle for the ordinary-save artifact export.

The Rust exporter proves that its five publication policies agree and that the
format readers can reopen their outputs.  This file deliberately uses only the
Python standard library and ZIP/XML parsing so that the export has a second,
independent check.  It is intended to run before any timed selector captures:

    python3 -B docs/performance/results/change-0819/artifact_audit.py \
      --artifacts docs/performance/results/change-0819/artifacts \
      --report docs/performance/results/change-0819/artifact-audit.json

The audit is strict about semantic preservation.  Compression, ZIP offsets,
and other serialization details may change, but untouched XML must have the
same canonical meaning and untouched binary members must have the same bytes.
An admitted edit is permitted only at the target described by the manifest;
an edit refusal must leave the package semantically unchanged.  A real input
is additionally checked against the caller-named source path in the manifest.
"""

from __future__ import annotations

import argparse
import copy
import difflib
import hashlib
import json
import posixpath
import struct
import urllib.parse
import zipfile
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET
from io import BytesIO


SCHEMA = "litchi.performance.0819.artifact-audit.v1"
EXPECTED_KIND = "ordinary-save-artifact-export"
EXPECTED_GENERATOR = "litchi-perf-ordinary-save-artifacts-v1"
MARKER = "litchi-perf-0638-ordinary-save"
POLICIES = ("default", "full", "file-only", "no-sync", "stream")
FORMATS = ("DOCX", "XLSX", "PPTX")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P_STRICT_NS = "http://purl.oclc.org/ooxml/presentationml/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"

XML_SUFFIXES = (".xml", ".rels", ".vml")
PML_SHAPE_LOCALS = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp", "contentPart"}

# These are deliberately finite independent-oracle limits.  The 32 MiB
# package ceiling is the ordinary-save input contract; the decompressed and
# XML ceilings keep a malformed or highly compressed artifact from turning
# this reader into an unbounded allocation path.
MAX_ARCHIVE_BYTES = 32 * 1024 * 1024
MAX_ZIP_ENTRIES = 100_000
MAX_MEMBER_COMPRESSED_BYTES = 32 * 1024 * 1024
MAX_MEMBER_DECOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_TOTAL_DECOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_XML_BYTES = 32 * 1024 * 1024
MAX_XML_DEPTH = 256
MAX_XML_EVENTS = 1_000_000
MAX_XML_ATTRIBUTE_BYTES = 64 * 1024
MAX_RELATIONSHIPS_PER_PART = 100_000
MAX_TOTAL_RELATIONSHIPS = 1_000_000

XLSX_DIRTY_CALC_PR = {
    "calcId": "0",
    "fullCalcOnLoad": "true",
    "calcCompleted": "false",
    "calcOnSave": "true",
    "forceFullCalc": "true",
}
XLSX_KNOWN_CALC_PR_ATTRIBUTES = set(XLSX_DIRTY_CALC_PR) | {
    "calcMode",
    "iterate",
    "iterateCount",
    "iterateDelta",
    "refMode",
}
GENERATED_XLSX_MEDIA_PREFIX = "xl/media/litchi-cell-crud-"
GENERATED_XLSX_MEDIA_COUNT = 8
GENERATED_XLSX_MEDIA_BYTES = 512 * 1024
GENERATED_XLSX_ENTRY_BYTES = 4


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def local_name(tag: str) -> str:
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def namespace_uri(tag: str) -> str:
    if not isinstance(tag, str):
        return ""
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


def is_xml_member(name: str) -> bool:
    return name.lower().endswith(XML_SUFFIXES)


def _remove_formatting_whitespace(element: ET.Element) -> None:
    # Whitespace between elements is formatting.  Whitespace in a leaf text
    # node is retained because it can be a real document/shape string.
    if len(element) and element.text is not None and element.text.isspace():
        element.text = None
    for child in element:
        _remove_formatting_whitespace(child)
        if child.tail is not None and child.tail.isspace():
            child.tail = None


def parse_xml(data: bytes) -> ET.Element:
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True))
    root = ET.fromstring(data, parser=parser)
    _remove_formatting_whitespace(root)
    return root


def canonical_xml(root_or_bytes: ET.Element | bytes) -> bytes:
    root = parse_xml(root_or_bytes) if isinstance(root_or_bytes, bytes) else root_or_bytes
    raw = ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    try:
        value = ET.canonicalize(xml_data=raw, with_comments=True)
    except TypeError:  # Python versions before the keyword was accepted.
        value = ET.canonicalize(xml_data=raw)
    return value.encode("utf-8") if isinstance(value, str) else value


def canonical_hash(data: bytes) -> str | None:
    try:
        return sha256(canonical_xml(data))
    except (ET.ParseError, ValueError, UnicodeError):
        return None


def namespace_census(data: bytes) -> list[list[str]]:
    """Return lexical namespace declarations in document order.

    Canonical XML intentionally normalizes prefixes.  The census keeps a
    separate record of prefix/declaration changes so that normalization is
    never mistaken for byte preservation.
    """
    result: list[list[str]] = []
    for _event, (prefix, uri) in ET.iterparse(
        BytesIO(data), events=("start-ns",)
    ):
        result.append([prefix or "", uri])
    return result


def xml_bounds(root: ET.Element, data: bytes, name: str) -> None:
    if len(data) > MAX_XML_BYTES:
        raise ValueError(f"{name}: XML member exceeds {MAX_XML_BYTES} bytes")
    events = 0
    attributes = 0

    def visit(node: ET.Element, depth: int) -> None:
        nonlocal events, attributes
        events += 1
        if events > MAX_XML_EVENTS:
            raise ValueError(f"{name}: XML event bound exceeded")
        if depth > MAX_XML_DEPTH:
            raise ValueError(f"{name}: XML depth bound exceeded")
        attributes += sum(
            len(str(key).encode("utf-8")) + len(str(value).encode("utf-8"))
            for key, value in node.attrib.items()
        )
        if attributes > MAX_XML_ATTRIBUTE_BYTES:
            raise ValueError(f"{name}: XML attribute-byte bound exceeded")
        for child in node:
            if isinstance(child.tag, str):
                visit(child, depth + 1)

    visit(root, 1)


def limited_diff(left: bytes, right: bytes, limit: int = 120) -> list[str]:
    left_lines = left.decode("utf-8", "replace").splitlines()
    right_lines = right.decode("utf-8", "replace").splitlines()
    result = list(
        difflib.unified_diff(
            left_lines,
            right_lines,
            fromfile="source-canonical",
            tofile="output-canonical",
            lineterm="",
            n=2,
        )
    )
    return result[:limit]


def _compressed_payload(archive: bytes, info: zipfile.ZipInfo) -> bytes:
    offset = info.header_offset
    if archive[offset : offset + 4] != b"PK\x03\x04":
        raise ValueError(f"{info.filename}: local ZIP header is absent")
    if offset + 30 > len(archive):
        raise ValueError(f"{info.filename}: truncated local ZIP header")
    name_len, extra_len = struct.unpack_from("<HH", archive, offset + 26)
    start = offset + 30 + name_len + extra_len
    end = start + info.compress_size
    if end > len(archive):
        raise ValueError(f"{info.filename}: truncated compressed payload")
    return archive[start:end]


def zip_metadata(info: zipfile.ZipInfo) -> dict[str, Any]:
    # Offsets and timestamps are serialization details; the remaining fields
    # describe the member's ZIP identity and are useful preservation evidence.
    return {
        "compress_type": info.compress_type,
        "flag_bits": info.flag_bits,
        "create_system": info.create_system,
        "create_version": info.create_version,
        "extract_version": info.extract_version,
        "internal_attr": info.internal_attr,
        "external_attr": info.external_attr,
        "CRC": info.CRC,
        "file_size": info.file_size,
        "compress_size": info.compress_size,
        "extra": info.extra.hex(),
    }


def read_archive(path: Path) -> dict[str, Any]:
    archive = path.read_bytes()
    if len(archive) > MAX_ARCHIVE_BYTES:
        raise ValueError(f"{path}: archive exceeds {MAX_ARCHIVE_BYTES} bytes")
    members: dict[str, dict[str, Any]] = {}
    with zipfile.ZipFile(path) as zf:
        infos = zf.infolist()
        if len(infos) > MAX_ZIP_ENTRIES:
            raise ValueError(f"{path}: ZIP entry bound exceeded")
        total_decompressed = 0
        for info in infos:
            name = info.filename
            normalized_name = name.rstrip("/")
            name_parts = normalized_name.split("/") if normalized_name else []
            if (
                not name
                or name.startswith(("/", "\\"))
                or "\\" in name
                or ":" in name.split("/", 1)[0]
                or any(part in ("", ".", "..") for part in name_parts)
                or posixpath.normpath(normalized_name) != normalized_name
            ):
                raise ValueError(f"{path}: unsafe ZIP member name {name!r}")
            if info.flag_bits & 0x1:
                raise ValueError(f"{path}: encrypted ZIP member {name!r}")
            if info.compress_size > MAX_MEMBER_COMPRESSED_BYTES:
                raise ValueError(f"{path}: compressed member bound exceeded for {name!r}")
            if info.file_size > MAX_MEMBER_DECOMPRESSED_BYTES:
                raise ValueError(f"{path}: decompressed member bound exceeded for {name!r}")
            total_decompressed += info.file_size
            if total_decompressed > MAX_TOTAL_DECOMPRESSED_BYTES:
                raise ValueError(f"{path}: aggregate decompressed member bound exceeded")
            if info.filename in members:
                raise ValueError(f"duplicate ZIP member {info.filename!r} in {path}")
            data = zf.read(info)
            compressed = _compressed_payload(archive, info)
            if is_xml_member(info.filename):
                try:
                    parsed = parse_xml(data)
                    xml_bounds(parsed, data, info.filename)
                    semantic = sha256(canonical_xml(parsed))
                    namespaces = namespace_census(data)
                except (ET.ParseError, ValueError, UnicodeError) as exc:
                    raise ValueError(f"{info.filename}: XML parse failed: {exc}") from exc
            else:
                semantic = None
                namespaces = None
            members[info.filename] = {
                "data": data,
                "compressed": compressed,
                "compressed_sha256": sha256(compressed),
                "metadata": zip_metadata(info),
                "xml": is_xml_member(info.filename),
                "canonical_sha256": semantic,
                "namespace_census": namespaces,
            }
    return {"path": str(path), "bytes": archive, "members": members}


def rels_targets(archive: dict[str, Any], rels_name: str) -> dict[str, tuple[str, str]]:
    member = archive["members"].get(rels_name)
    if not member:
        return {}
    root = parse_xml(member["data"])
    result: dict[str, tuple[str, str]] = {}
    for rel in root:
        if local_name(rel.tag) != "Relationship":
            continue
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rid or target is None:
            continue
        result[rid] = (target, rel.attrib.get("Type", ""))
    return result


def resolve_target(source_member: str, target: str) -> str:
    target = urllib.parse.unquote(target)
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_member), target))


def workbook_sheets(archive: dict[str, Any]) -> dict[str, tuple[str, str]]:
    workbook = archive["members"].get("xl/workbook.xml")
    if not workbook:
        return {}
    rels = rels_targets(archive, "xl/_rels/workbook.xml.rels")
    root = parse_xml(workbook["data"])
    result: dict[str, tuple[str, str]] = {}
    for sheet in root.iter():
        if local_name(sheet.tag) != "sheet":
            continue
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get(f"{{{R_NS}}}id")
        if not name or not rid or rid not in rels:
            continue
        target, rel_type = rels[rid]
        if rel_type.endswith("/worksheet"):
            result[name] = (resolve_target("xl/workbook.xml", target), rid)
    return result


def presentation_slides(archive: dict[str, Any]) -> list[str]:
    presentation = archive["members"].get("ppt/presentation.xml")
    if not presentation:
        return []
    rels = rels_targets(archive, "ppt/_rels/presentation.xml.rels")
    root = parse_xml(presentation["data"])
    result = []
    for item in root.iter():
        if local_name(item.tag) != "sldId":
            continue
        rid = item.attrib.get(f"{{{R_NS}}}id")
        if rid in rels:
            target, rel_type = rels[rid]
            if rel_type.endswith("/slide"):
                result.append(resolve_target("ppt/presentation.xml", target))
    return result


def content_type_graph(archive: dict[str, Any]) -> dict[str, Any]:
    member = archive["members"].get("[Content_Types].xml")
    if not member:
        return {"valid": False, "errors": ["[Content_Types].xml is absent"]}
    root = parse_xml(member["data"])
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for item in root:
        kind = local_name(item.tag)
        if kind == "Default" and item.attrib.get("Extension"):
            defaults[item.attrib["Extension"].lower()] = item.attrib.get("ContentType", "")
        elif kind == "Override" and item.attrib.get("PartName"):
            overrides[item.attrib["PartName"].lstrip("/")] = item.attrib.get("ContentType", "")
    errors: list[str] = []
    if len(defaults) + len(overrides) > MAX_ZIP_ENTRIES:
        errors.append("content-type mapping bound exceeded")
    resolved: dict[str, str] = {}
    for name in sorted(archive["members"]):
        if name == "[Content_Types].xml" or name.endswith("/"):
            continue
        content_type = overrides.get(name)
        if content_type is None:
            extension = name.rsplit("/", 1)[-1].rsplit(".", 1)[-1].lower() if "." in name else ""
            content_type = defaults.get(extension)
        if not content_type:
            errors.append(f"no content type for {name}")
        else:
            resolved[name] = content_type
    return {
        "valid": not errors,
        "errors": errors,
        "defaults": defaults,
        "overrides": overrides,
        "resolved": resolved,
    }


def relationship_graph(archive: dict[str, Any]) -> dict[str, Any]:
    edges: list[dict[str, str]] = []
    errors: list[str] = []
    names = set(archive["members"])
    for rel_name, member in sorted(archive["members"].items()):
        if not rel_name.endswith(".rels"):
            continue
        if rel_name == "_rels/.rels":
            owner = ""
        else:
            marker = "/_rels/"
            if marker not in rel_name:
                errors.append(f"relationship part {rel_name} is not under _rels")
                continue
            parent, child = rel_name.split(marker, 1)
            owner = f"{parent}/{child[:-5]}"
        root = parse_xml(member["data"])
        relation_count = sum(1 for item in root if local_name(item.tag) == "Relationship")
        if relation_count > MAX_RELATIONSHIPS_PER_PART:
            errors.append(f"{rel_name}: relationship-count bound exceeded")
        if len(edges) + relation_count > MAX_TOTAL_RELATIONSHIPS:
            errors.append("aggregate relationship-count bound exceeded")
        for item in root:
            if local_name(item.tag) != "Relationship":
                continue
            target = item.attrib.get("Target", "")
            mode = item.attrib.get("TargetMode", "Internal")
            if mode.lower() == "external":
                resolved = "external"
            else:
                resolved = resolve_target(owner, target)
                if resolved not in names:
                    errors.append(f"{rel_name}: internal target {resolved!r} is absent")
            edges.append(
                {
                    "part": rel_name,
                    "owner": owner,
                    "id": item.attrib.get("Id", ""),
                    "type": item.attrib.get("Type", ""),
                    "target": target,
                    "resolved": resolved,
                    "mode": mode,
                }
            )
    return {"valid": not errors, "errors": errors, "edges": edges}


def text_content(element: ET.Element, text_namespace: str) -> str:
    return "".join(
        node.text or ""
        for node in element.iter()
        if node.tag == f"{{{text_namespace}}}t"
    )


def docx_paragraphs(data: bytes) -> list[tuple[str, ET.Element]]:
    root = parse_xml(data)
    return [
        (text_content(node, W_NS), node)
        for node in root.iter()
        if node.tag == f"{{{W_NS}}}p"
    ]


def worksheet_cell_semantics(
    archive: dict[str, Any],
) -> dict[str, dict[str, tuple[str, str, str | None, str | None]]]:
    strings: list[str] = []
    shared = archive["members"].get("xl/sharedStrings.xml")
    if shared:
        root = parse_xml(shared["data"])
        strings = [text_content(item, S_NS) for item in root if local_name(item.tag) == "si"]
    result: dict[str, dict[str, tuple[str, str, str | None, str | None]]] = {}
    for sheet_name, (member_name, _rid) in workbook_sheets(archive).items():
        member = archive["members"].get(member_name)
        if not member:
            result[sheet_name] = {}
            continue
        root = parse_xml(member["data"])
        cells: dict[str, tuple[str, str, str | None, str | None]] = {}
        for cell in root.iter():
            if local_name(cell.tag) != "c":
                continue
            ref = cell.attrib.get("r")
            if not ref:
                continue
            kind = cell.attrib.get("t", "n")
            value_node = next((x for x in cell if local_name(x.tag) == "v"), None)
            value = value_node.text if value_node is not None and value_node.text is not None else ""
            if kind == "s":
                try:
                    value = strings[int(value)]
                except (ValueError, IndexError):
                    value = f"<invalid-shared-string:{value}>"
                semantic_kind = "text"
            elif kind == "inlineStr":
                inline = next((x for x in cell if local_name(x.tag) == "is"), None)
                value = text_content(inline, S_NS) if inline is not None else ""
                semantic_kind = "text"
            elif kind == "str":
                semantic_kind = "text"
            elif kind == "b":
                semantic_kind = "bool"
            elif kind == "e":
                semantic_kind = "error"
            else:
                semantic_kind = "number"
            formula = next((x for x in cell if local_name(x.tag) == "f"), None)
            formula_text = formula.text if formula is not None else None
            cells[ref] = (
                semantic_kind,
                value,
                formula_text,
                cell.attrib.get("s"),
            )
        result[sheet_name] = cells
    return result


def pptx_shape_nodes(root: ET.Element) -> list[ET.Element]:
    nodes: list[ET.Element] = []
    for node in root.iter():
        if namespace_uri(node.tag) in (P_NS, P_STRICT_NS) and local_name(node.tag) in PML_SHAPE_LOCALS:
            nodes.append(node)
    return nodes


def pptx_shape_texts(data: bytes) -> list[str]:
    root = parse_xml(data)
    return [text_content(node, A_NS) for node in pptx_shape_nodes(root)]


def replace_target_texts(root: ET.Element, target: ET.Element, namespace: str) -> bool:
    found = False
    for index, node in enumerate(target.iter()):
        if node.tag == f"{{{namespace}}}t":
            node.text = f"__LITCHI_TARGET_TEXT_{index}__"
            found = True
    return found


def remove_target_paragraph(root: ET.Element, target: ET.Element) -> bool:
    def visit(parent: ET.Element) -> bool:
        for index, child in enumerate(list(parent)):
            if child is target:
                parent.remove(child)
                return True
            if visit(child):
                return True
        return False

    return visit(root)


def remove_all_cells(root: ET.Element) -> None:
    for parent in root.iter():
        for child in list(parent):
            if local_name(child.tag) == "c" and namespace_uri(child.tag) == S_NS:
                parent.remove(child)


def excel_column_name(column: int) -> str:
    """Return the bounded A1 column spelling used by the generated corpus."""
    value = column + 1
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def remove_direct_xlsx_calc_pr(root: ET.Element) -> tuple[ET.Element, int]:
    """Copy a workbook and remove only direct spreadsheet ``calcPr`` nodes."""
    result = copy.deepcopy(root)
    removed = 0
    for child in list(result):
        if child.tag == f"{{{S_NS}}}calcPr":
            result.remove(child)
            removed += 1
    return result, removed


def archive_inventory(archive: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        {
            "name": name,
            "compressed_sha256": value["compressed_sha256"],
            "decompressed_sha256": sha256(value["data"]),
            "canonical_sha256": value["canonical_sha256"],
            "namespace_census": value["namespace_census"],
            "metadata": value["metadata"],
        }
        for name, value in sorted(archive["members"].items())
    ]


class Audit:
    def __init__(self, repo_root: Path):
        self.repo_root = repo_root
        self.errors: list[str] = []

    def error(self, message: str) -> None:
        self.errors.append(message)

    def resolve_manifest_path(self, path: str, root: Path) -> Path:
        candidate = Path(path)
        if not candidate.is_absolute():
            candidate = root / candidate
        resolved = candidate.resolve()
        if root.resolve() not in resolved.parents and resolved != root.resolve():
            raise ValueError(f"artifact path escapes output directory: {path!r}")
        return resolved

    def resolve_input_path(self, path: str) -> Path:
        candidate = Path(path)
        if not candidate.is_absolute():
            candidate = self.repo_root / candidate
        return candidate.resolve()

    def validate_manifest_identity(
        self, manifest: dict[str, Any], case: dict[str, Any], source: dict[str, Any], case_id: str
    ) -> None:
        fmt = case.get("format")
        expected_main = {
            "DOCX": "word/document.xml",
            "XLSX": "xl/workbook.xml",
            "PPTX": "ppt/presentation.xml",
        }.get(fmt)
        corpus = case.get("corpus") or {}
        actual_bytes = source["bytes"]
        actual_members = source["members"]
        if case.get("source_archive_bytes") != len(actual_bytes):
            self.error(f"{case_id}: source_archive_bytes disagrees with source file")
        if case.get("source_archive_sha256") != sha256(actual_bytes):
            self.error(f"{case_id}: source_archive_sha256 disagrees with source file")
        source_file = case.get("source_archive") or {}
        if source_file.get("bytes") != len(actual_bytes) or source_file.get("sha256") != sha256(actual_bytes):
            self.error(f"{case_id}: source_archive identity disagrees with source file")
        if corpus.get("archive_member_count") != len(actual_members):
            self.error(f"{case_id}: corpus archive_member_count disagrees with ZIP inventory")
        if corpus.get("archive_bytes") != len(actual_bytes) or corpus.get("archive_sha256") != sha256(actual_bytes):
            self.error(f"{case_id}: corpus archive identity disagrees with ZIP inventory")
        edit_target = case.get("edit_target") or {}
        if expected_main and edit_target.get("main_part") != expected_main:
            self.error(f"{case_id}: edit_target.main_part is not {expected_main!r}")

        # The generated builders expose logical workload counters in these
        # fields.  They are deliberately kept separate from ZIP inventory:
        # entry_count and uncompressed_payload_bytes count semantic records,
        # while archive_member_count and archive_bytes describe the package.
        if case.get("origin") == "generated-harness-corpus":
            self.validate_generated_manifest(case, source, corpus, case_id)
        elif case.get("origin") == "caller-named-real-file":
            if corpus.get("entry_count") != len(actual_members):
                self.error(f"{case_id}: real corpus entry_count disagrees with ZIP inventory")
            if expected_main:
                target = actual_members.get(expected_main)
                if target is None:
                    self.error(f"{case_id}: target member {expected_main!r} is absent")
                else:
                    if corpus.get("target_entry") != expected_main:
                        self.error(f"{case_id}: real corpus target_entry is not {expected_main!r}")
                    if corpus.get("target_payload_bytes") != len(target["data"]):
                        self.error(f"{case_id}: target_payload_bytes disagrees with decompressed member")
                    if corpus.get("target_payload_sha256") != sha256(target["data"]):
                        self.error(f"{case_id}: target_payload_sha256 disagrees with decompressed member")
            if corpus.get("uncompressed_payload_bytes") != sum(
                len(member["data"]) for member in actual_members.values()
            ):
                self.error(f"{case_id}: real uncompressed_payload_bytes disagrees with member inventory")
        else:
            self.error(f"{case_id}: unsupported corpus origin {case.get('origin')!r}")
        if case.get("format") == "XLSX":
            if edit_target.get("xlsx_address") != "A1":
                self.error(f"{case_id}: XLSX edit target is not A1")
            if not edit_target.get("xlsx_sheet"):
                self.error(f"{case_id}: XLSX edit target has no worksheet name")
        if case.get("format") == "PPTX" and case.get("edit_admitted"):
            if not isinstance(edit_target.get("pptx_slide"), int) or not isinstance(edit_target.get("pptx_shape"), int):
                self.error(f"{case_id}: admitted PPTX edit has no numeric target")
        admitted = case.get("edit_admitted")
        outcome = case.get("edit_outcome")
        if admitted and outcome != "admitted":
            self.error(f"{case_id}: edit_admitted=true but edit_outcome is {outcome!r}")
        if not admitted and (not isinstance(outcome, str) or not outcome.startswith("refused:")):
            self.error(f"{case_id}: edit_admitted=false but edit_outcome is {outcome!r}")

    def validate_generated_manifest(
        self,
        case: dict[str, Any],
        source: dict[str, Any],
        corpus: dict[str, Any],
        case_id: str,
    ) -> None:
        """Recompute the generated logical counters from bounded source XML.

        These checks mirror the three frozen source constructors as semantic
        projections.  They never substitute physical ZIP member counts for
        the declared logical workload, and they do not use the Rust project as
        a runtime dependency.
        """
        fmt = case.get("format")
        members = source["members"]

        def expect(field: str, value: Any) -> None:
            if corpus.get(field) != value:
                self.error(f"{case_id}: generated {field} disagrees with semantic source ({value!r})")

        if fmt == "DOCX":
            member = members.get("word/document.xml")
            if member is None:
                self.error(f"{case_id}: generated DOCX main part is absent")
                return
            paragraphs = docx_paragraphs(member["data"])
            expected_texts = [
                f"litchi-perf-baseline-docx-semantic-v1-source-{index:05}"
                for index in range(200)
            ]
            actual_texts = [text for text, _node in paragraphs]
            if actual_texts != expected_texts:
                self.error(f"{case_id}: generated DOCX semantic paragraph projection changed")
            payloads = [text.encode("utf-8") for text in actual_texts]
            target = payloads[0] if payloads else b""
            expect("entry_count", len(payloads))
            expect("entry_bytes", len(target))
            expect("uncompressed_payload_bytes", sum(map(len, payloads)))
            expect("target_entry", "paragraph:0")
            expect("target_payload_bytes", len(target))
            expect("target_payload_sha256", sha256(target))
            return

        if fmt == "PPTX":
            slides = presentation_slides(source)
            actual_texts: list[str] = []
            expected_texts: list[str] = []
            for slide_index, slide_name in enumerate(slides):
                slide_member = members.get(slide_name)
                if slide_member is None:
                    self.error(f"{case_id}: generated PPTX slide member {slide_name!r} is absent")
                    continue
                texts = pptx_shape_texts(slide_member["data"])
                actual_texts.extend(texts)
                expected_texts.extend(
                    f"litchi-perf-baseline-pptx-semantic-v1-source-{slide_index:03}-{shape_index:03}"
                    for shape_index in range(8)
                )
                if len(texts) != 8:
                    self.error(f"{case_id}: generated PPTX slide {slide_index} shape count is not 8")
            if actual_texts != expected_texts or len(slides) != 12:
                self.error(f"{case_id}: generated PPTX semantic shape projection changed")
            payloads = [text.encode("utf-8") for text in actual_texts]
            target = payloads[0] if payloads else b""
            expect("entry_count", len(payloads))
            expect("entry_bytes", len(target))
            expect("uncompressed_payload_bytes", sum(map(len, payloads)))
            expect("target_entry", "slide:0/shape:0")
            expect("target_payload_bytes", len(target))
            expect("target_payload_sha256", sha256(target))
            return

        if fmt == "XLSX":
            sheets = workbook_sheets(source)
            expected_sheet_names = ["Sheet1", "Bench01", "Bench02", "Bench03"]
            if list(sheets) != expected_sheet_names:
                self.error(f"{case_id}: generated XLSX sheet catalog differs from medium constructor")
            cells = worksheet_cell_semantics(source)
            expected_cells: dict[str, dict[str, tuple[str, str, str | None, str | None]]] = {}
            for sheet_index, sheet_name in enumerate(expected_sheet_names):
                expected_sheet: dict[str, tuple[str, str, str | None, str | None]] = {}
                for row in range(48):
                    for column in range(48):
                        address = f"{excel_column_name(column)}{row + 1}"
                        value = str(sheet_index * 1_000_000 + row * 1_000 + column)
                        expected_sheet[address] = ("number", value, None, None)
                expected_cells[sheet_name] = expected_sheet
            if cells != expected_cells:
                self.error(f"{case_id}: generated XLSX semantic cell projection changed")
            cell_count = sum(len(sheet) for sheet in cells.values())
            target = cells.get("Sheet1", {}).get("A1")
            target_payload = (target[1].encode("utf-8") if target is not None else b"")
            media_names = sorted(
                name for name in members if name.startswith(GENERATED_XLSX_MEDIA_PREFIX)
            )
            expected_media_names = [
                f"{GENERATED_XLSX_MEDIA_PREFIX}{index:02}.png"
                for index in range(GENERATED_XLSX_MEDIA_COUNT)
            ]
            if media_names != expected_media_names:
                self.error(f"{case_id}: generated XLSX media inventory differs from constructor")
            for name in expected_media_names:
                member = members.get(name)
                if member is None or len(member["data"]) != GENERATED_XLSX_MEDIA_BYTES:
                    self.error(f"{case_id}: generated XLSX media payload size is not bounded constructor size for {name}")
            logical_bytes = cell_count * GENERATED_XLSX_ENTRY_BYTES + (
                GENERATED_XLSX_MEDIA_COUNT * GENERATED_XLSX_MEDIA_BYTES
            )
            expect("entry_count", cell_count)
            expect("entry_bytes", GENERATED_XLSX_ENTRY_BYTES)
            expect("uncompressed_payload_bytes", logical_bytes)
            expect("target_entry", "Sheet1!A1")
            expect("target_payload_bytes", len(target_payload))
            expect("target_payload_sha256", sha256(target_payload))
            xlsx = corpus.get("xlsx") or {}
            expected_xlsx = {
                "sheet_count": 4,
                "rows_per_sheet": 48,
                "columns_per_sheet": 48,
                "one_percent_update_count": 93,
                "source_members": {
                    "workbook": "xl/workbook.xml",
                    "worksheets": [
                        "xl/worksheets/sheet1.xml",
                        "xl/worksheets/sheet2.xml",
                        "xl/worksheets/sheet3.xml",
                        "xl/worksheets/sheet4.xml",
                    ],
                    "shared_strings": None,
                    "styles": "xl/styles.xml",
                },
            }
            if xlsx != expected_xlsx:
                self.error(f"{case_id}: generated XLSX constructor metadata disagrees with manifest")
            return

        self.error(f"{case_id}: generated semantic accounting has unsupported format {fmt!r}")

    def validate_external_input(
        self, case: dict[str, Any], source: dict[str, Any], case_id: str
    ) -> None:
        input_path = case.get("input_path")
        if case.get("origin") == "caller-named-real-file":
            if not input_path:
                self.error(f"{case_id}: real case has no input_path")
                return
            path = self.resolve_input_path(input_path)
            if not path.is_file():
                self.error(f"{case_id}: real input path is absent: {path}")
                return
            expected = path.read_bytes()
            if expected != source["bytes"]:
                self.error(f"{case_id}: staged source differs byte-for-byte from input_path")
            if sha256(expected) != case.get("source_archive_sha256"):
                self.error(f"{case_id}: input_path hash disagrees with manifest source identity")
        elif input_path is not None:
            self.error(f"{case_id}: generated case unexpectedly carries input_path")

    def compare_archives(
        self, source: dict[str, Any], output: dict[str, Any], case_id: str
    ) -> dict[str, Any]:
        source_names = set(source["members"])
        output_names = set(output["members"])
        missing = sorted(source_names - output_names)
        extra = sorted(output_names - source_names)
        if missing:
            self.error(f"{case_id}: output is missing ZIP members: {missing}")
        if extra:
            self.error(f"{case_id}: output has unexpected ZIP members: {extra}")
        rows: list[dict[str, Any]] = []
        for name in sorted(source_names | output_names):
            left = source["members"].get(name)
            right = output["members"].get(name)
            if left is None or right is None:
                continue
            xml = bool(left["xml"] or right["xml"])
            decompressed_equal = left["data"] == right["data"]
            canonical_equal = (
                left["canonical_sha256"] is not None
                and left["canonical_sha256"] == right["canonical_sha256"]
            ) if xml else None
            compressed_equal = left["compressed"] == right["compressed"]
            metadata_equal = left["metadata"] == right["metadata"]
            row = {
                "name": name,
                "xml": xml,
                "decompressed_equal": decompressed_equal,
                "canonical_equal": canonical_equal,
                "namespace_census_equal": left["namespace_census"] == right["namespace_census"],
                "source_namespace_census": left["namespace_census"],
                "output_namespace_census": right["namespace_census"],
                "compressed_equal": compressed_equal,
                "metadata_equal": metadata_equal,
                "source_decompressed_sha256": sha256(left["data"]),
                "output_decompressed_sha256": sha256(right["data"]),
                "source_compressed_sha256": left["compressed_sha256"],
                "output_compressed_sha256": right["compressed_sha256"],
            }
            if xml and not canonical_equal:
                row["canonical_diff"] = limited_diff(canonical_xml(left["data"]), canonical_xml(right["data"]))
            rows.append(row)
        return {"missing": missing, "extra": extra, "members": rows}

    def check_docx_target(
        self, source: dict[str, Any], output: dict[str, Any], admitted: bool, case_id: str
    ) -> set[str]:
        name = "word/document.xml"
        if not admitted:
            return set()
        left = source["members"].get(name)
        right = output["members"].get(name)
        if not left or not right:
            return {name}
        source_paras = docx_paragraphs(left["data"])
        output_paras = docx_paragraphs(right["data"])
        if len(output_paras) != len(source_paras) + 1 or output_paras[-1][0] != MARKER:
            self.error(f"{case_id}: DOCX output does not append exactly one marker paragraph")
            return {name}
        if [text for text, _ in output_paras[:-1]] != [text for text, _ in source_paras]:
            self.error(f"{case_id}: DOCX non-target paragraph semantics changed")
        root = parse_xml(right["data"])
        paragraphs = [node for node in root.iter() if node.tag == f"{{{W_NS}}}p"]
        if not paragraphs or not remove_target_paragraph(root, paragraphs[-1]):
            self.error(f"{case_id}: DOCX marker paragraph could not be removed for comparison")
        elif canonical_xml(parse_xml(left["data"])) != canonical_xml(root):
            self.error(f"{case_id}: DOCX XML changed outside the appended paragraph")
        return {name}

    def check_xlsx_target(
        self, source: dict[str, Any], output: dict[str, Any], case: dict[str, Any], case_id: str
    ) -> set[str]:
        target = case.get("edit_target") or {}
        sheet_name = target.get("xlsx_sheet")
        address = target.get("xlsx_address")
        source_sheets = workbook_sheets(source)
        output_sheets = workbook_sheets(output)
        if sheet_name not in source_sheets or sheet_name not in output_sheets:
            self.error(f"{case_id}: target worksheet {sheet_name!r} is absent")
            return {"xl/sharedStrings.xml"}
        source_path = source_sheets[sheet_name][0]
        output_path = output_sheets[sheet_name][0]
        if source_path != output_path:
            self.error(f"{case_id}: target worksheet relationship moved from {source_path} to {output_path}")
        source_cells = worksheet_cell_semantics(source)
        output_cells = worksheet_cell_semantics(output)
        source_target = source_cells.get(sheet_name, {}).get(address)
        output_target = output_cells.get(sheet_name, {}).get(address)
        admitted = bool(case.get("edit_admitted"))
        if admitted:
            if output_target is None or output_target[0] != "text" or output_target[1] != MARKER:
                self.error(f"{case_id}: XLSX target {sheet_name}!{address} is not the marker text")
            if source_target and output_target and source_target[3] != output_target[3]:
                self.error(f"{case_id}: XLSX target cell style changed")
        elif output_target != source_target:
            self.error(f"{case_id}: refused XLSX edit changed target cell semantics")
        for name in sorted(set(source_cells) | set(output_cells)):
            left = dict(source_cells.get(name, {}))
            right = dict(output_cells.get(name, {}))
            if name == sheet_name:
                left.pop(address, None)
                right.pop(address, None)
            if left != right:
                self.error(f"{case_id}: XLSX non-target cell semantics changed on {name}")
        # Compare sheet structure with all cells removed; this catches changed
        # dimensions, row/column metadata, formulas, and extension structure
        # that the cell-value map cannot represent.
        for name in sorted(set(source_sheets) & set(output_sheets)):
            left = parse_xml(source["members"][source_sheets[name][0]]["data"])
            right = parse_xml(output["members"][output_sheets[name][0]]["data"])
            remove_all_cells(left)
            remove_all_cells(right)
            if canonical_xml(left) != canonical_xml(right):
                self.error(f"{case_id}: XLSX worksheet structure changed outside target cells on {name}")
        if admitted:
            self.check_xlsx_calc_pr_closure(source, output, case_id)
            return {source_path, output_path, "xl/sharedStrings.xml", "xl/workbook.xml"}
        return {source_path, output_path, "xl/sharedStrings.xml"}

    def check_xlsx_calc_pr_closure(
        self, source: dict[str, Any], output: dict[str, Any], case_id: str
    ) -> None:
        """Allow exactly the workbook dirty-calculation node owned by a cell edit."""
        source_member = source["members"].get("xl/workbook.xml")
        output_member = output["members"].get("xl/workbook.xml")
        if source_member is None or output_member is None:
            self.error(f"{case_id}: XLSX workbook part is absent for calcPr closure")
            return
        source_root = parse_xml(source_member["data"])
        output_root = parse_xml(output_member["data"])
        if source_root.tag != output_root.tag:
            self.error(f"{case_id}: XLSX workbook root tag changed")
        if source_root.attrib != output_root.attrib:
            self.error(f"{case_id}: XLSX workbook root attributes changed")
        if source_member["namespace_census"] != output_member["namespace_census"]:
            self.error(f"{case_id}: XLSX workbook namespace declarations changed")

        source_without_calc, source_calc_count = remove_direct_xlsx_calc_pr(source_root)
        output_without_calc, output_calc_count = remove_direct_xlsx_calc_pr(output_root)
        if source_calc_count > 1:
            self.error(f"{case_id}: XLSX source workbook has multiple direct calcPr nodes")
        if output_calc_count != 1:
            self.error(f"{case_id}: XLSX edited workbook does not have exactly one direct calcPr node")
        if canonical_xml(source_without_calc) != canonical_xml(output_without_calc):
            self.error(f"{case_id}: XLSX workbook XML changed outside direct calcPr")

        source_calc = next(
            (child for child in source_root if child.tag == f"{{{S_NS}}}calcPr"), None
        )
        if source_calc is not None:
            unknown_source_attributes = set(source_calc.attrib) - XLSX_KNOWN_CALC_PR_ATTRIBUTES
            if unknown_source_attributes:
                self.error(
                    f"{case_id}: XLSX source calcPr has unsupported attributes: "
                    f"{sorted(unknown_source_attributes)}"
                )
            if (
                source_calc.text not in (None, "")
                or source_calc.tail not in (None, "")
                or len(source_calc)
            ):
                self.error(f"{case_id}: XLSX source calcPr contains unsupported child or text content")

        output_calc = next(
            (child for child in output_root if child.tag == f"{{{S_NS}}}calcPr"), None
        )
        if output_calc is None:
            return
        if output_calc.attrib != XLSX_DIRTY_CALC_PR:
            self.error(f"{case_id}: XLSX calcPr attributes are not the exact dirty-edit flags")
        if (
            output_calc.text not in (None, "")
            or output_calc.tail not in (None, "")
            or len(output_calc)
        ):
            self.error(f"{case_id}: XLSX calcPr contains unexpected child or text content")

    def check_pptx_target(
        self, source: dict[str, Any], output: dict[str, Any], case: dict[str, Any], case_id: str
    ) -> set[str]:
        target = case.get("edit_target") or {}
        slide_index = target.get("pptx_slide")
        shape_index = target.get("pptx_shape")
        slides_source = presentation_slides(source)
        slides_output = presentation_slides(output)
        if not isinstance(slide_index, int) or not isinstance(shape_index, int):
            if case.get("edit_admitted"):
                self.error(f"{case_id}: admitted PPTX edit has no target coordinates")
            return set()
        if slide_index < 0 or slide_index >= len(slides_source) or slide_index >= len(slides_output):
            self.error(f"{case_id}: PPTX target slide is absent")
            return set()
        source_path = slides_source[slide_index]
        output_path = slides_output[slide_index]
        source_member = source["members"].get(source_path)
        output_member = output["members"].get(output_path)
        if not source_member or not output_member:
            self.error(f"{case_id}: PPTX target slide member is absent")
            return {source_path, output_path}
        source_root = parse_xml(source_member["data"])
        output_root = parse_xml(output_member["data"])
        source_shapes = pptx_shape_nodes(source_root)
        output_shapes = pptx_shape_nodes(output_root)
        if len(source_shapes) != len(output_shapes):
            self.error(f"{case_id}: PPTX shape inventory changed from {len(source_shapes)} to {len(output_shapes)}")
        if shape_index >= len(source_shapes) or shape_index >= len(output_shapes):
            self.error(f"{case_id}: PPTX target shape is absent")
            return {source_path, output_path}
        source_texts = [text_content(node, A_NS) for node in source_shapes]
        output_texts = [text_content(node, A_NS) for node in output_shapes]
        if case.get("edit_admitted"):
            if MARKER not in output_texts[shape_index]:
                self.error(f"{case_id}: PPTX target shape does not contain the marker")
        elif output_texts[shape_index] != source_texts[shape_index]:
            self.error(f"{case_id}: refused PPTX edit changed target shape text")
        for index, (left, right) in enumerate(zip(source_texts, output_texts)):
            if index != shape_index and left != right:
                self.error(f"{case_id}: PPTX non-target shape {index} text changed")
        if case.get("edit_admitted"):
            left_norm = copy.deepcopy(source_root)
            right_norm = copy.deepcopy(output_root)
            left_nodes = pptx_shape_nodes(left_norm)
            right_nodes = pptx_shape_nodes(right_norm)
            if not replace_target_texts(left_norm, left_nodes[shape_index], A_NS) or not replace_target_texts(
                right_norm, right_nodes[shape_index], A_NS
            ):
                self.error(f"{case_id}: PPTX target shape has no DrawingML text runs")
            elif canonical_xml(left_norm) != canonical_xml(right_norm):
                self.error(f"{case_id}: PPTX slide XML changed outside target text")
        return {source_path, output_path}

    def audit_case(
        self,
        manifest: dict[str, Any],
        case: dict[str, Any],
        artifacts_root: Path,
    ) -> dict[str, Any]:
        case_id = str(case.get("case_id", "<missing-case-id>"))
        result: dict[str, Any] = {"case_id": case_id, "format": case.get("format"), "origin": case.get("origin")}
        source_spec = case.get("source_archive") or {}
        try:
            source_path = self.resolve_manifest_path(str(source_spec["path"]), artifacts_root)
            source = read_archive(source_path)
        except Exception as exc:  # Keep auditing the other cases.
            self.error(f"{case_id}: cannot read source archive: {exc}")
            return {**result, "ok": False, "errors": [str(exc)]}
        self.validate_manifest_identity(manifest, case, source, case_id)
        self.validate_external_input(case, source, case_id)
        source_content_types = content_type_graph(source)
        source_relationships = relationship_graph(source)
        if not source_content_types["valid"]:
            self.error(f"{case_id}: source content-type graph is invalid: {source_content_types['errors']}")
        if not source_relationships["valid"]:
            self.error(f"{case_id}: source relationship graph is invalid: {source_relationships['errors']}")
        result["source_inventory"] = archive_inventory(source)
        result["source_content_types"] = source_content_types
        result["source_relationships"] = source_relationships
        published_sha = case.get("published_sha256")
        published_bytes = case.get("published_bytes")
        default_specs = [
            spec for spec in case.get("policy_outputs", []) if spec.get("policy") == "default"
        ]
        if len(default_specs) != 1 or published_sha != default_specs[0].get("sha256"):
            self.error(f"{case_id}: published_sha256 does not match default policy metadata")
        if len(case.get("policy_outputs", [])) != 4:
            self.error(f"{case_id}: exporter did not provide exactly four filesystem policies")
        if (case.get("stream_output") or {}).get("policy") != "stream":
            self.error(f"{case_id}: stream output does not identify policy=stream")
        policy_results: list[dict[str, Any]] = []
        default_output: dict[str, Any] | None = None
        for policy in POLICIES:
            if policy == "stream":
                spec = case.get("stream_output") or {}
            else:
                matches = [x for x in case.get("policy_outputs", []) if x.get("policy") == policy]
                if len(matches) != 1:
                    self.error(f"{case_id}: policy {policy!r} does not occur exactly once")
                    continue
                spec = matches[0]
            try:
                output_path = self.resolve_manifest_path(str((spec.get("output") or {})["path"]), artifacts_root)
                output = read_archive(output_path)
            except Exception as exc:
                self.error(f"{case_id}: cannot read {policy} output: {exc}")
                continue
            output_bytes = output["bytes"]
            output_sha = sha256(output_bytes)
            output_content_types = content_type_graph(output)
            output_relationships = relationship_graph(output)
            if not output_content_types["valid"]:
                self.error(f"{case_id}/{policy}: output content-type graph is invalid: {output_content_types['errors']}")
            if not output_relationships["valid"]:
                self.error(f"{case_id}/{policy}: output relationship graph is invalid: {output_relationships['errors']}")
            if output_sha != spec.get("sha256") or len(output_bytes) != spec.get("bytes"):
                self.error(f"{case_id}/{policy}: output metadata disagrees with bytes")
            if output_sha != published_sha or len(output_bytes) != published_bytes:
                self.error(f"{case_id}/{policy}: output does not match published reference identity")
            if not spec.get("matches_reference") or not spec.get("source_unchanged") or not spec.get("reopen_admitted"):
                self.error(f"{case_id}/{policy}: exporter gate flags are not all true")
            if bool(spec.get("matches_source")) != (output_bytes == source["bytes"]):
                self.error(f"{case_id}/{policy}: matches_source flag disagrees with archive bytes")
            if policy == "default":
                default_output = output
            elif default_output is not None and output_bytes != default_output["bytes"]:
                self.error(f"{case_id}: {policy} output is not byte-identical to default output")
            policy_results.append(
                {
                    "policy": policy,
                    "path": str(output_path),
                    "bytes": len(output_bytes),
                    "sha256": output_sha,
                    "matches_default": default_output is not None and output_bytes == default_output["bytes"],
                    "matches_source": output_bytes == source["bytes"],
                    "inventory": archive_inventory(output),
                    "content_types": output_content_types,
                    "relationships": output_relationships,
                }
            )
        if default_output is None:
            return {**result, "ok": False, "policy_outputs": policy_results}
        default_content_types = content_type_graph(default_output)
        default_relationships = relationship_graph(default_output)
        if source_content_types["resolved"] != default_content_types.get("resolved"):
            self.error(f"{case_id}: content-type assignments changed outside the edit closure")
        if source_relationships["edges"] != default_relationships.get("edges"):
            self.error(f"{case_id}: OPC relationship graph changed outside the edit closure")
        comparison = self.compare_archives(source, default_output, case_id)
        admitted = bool(case.get("edit_admitted"))
        if case.get("format") == "DOCX" and not admitted:
            if not case.get("refused_output_source_exact"):
                self.error(f"{case_id}: refused DOCX does not claim exact source preservation")
            if default_output["bytes"] != source["bytes"]:
                self.error(f"{case_id}: refused DOCX output is not byte-identical to source")
        if not admitted and default_output["bytes"] == source["bytes"]:
            pass
        elif not admitted:
            # A refusal may reserialize, but must retain all package semantics.
            for row in comparison["members"]:
                if row["xml"] and not row["canonical_equal"]:
                    self.error(f"{case_id}: refused edit changed XML member {row['name']}")
                if not row["xml"] and not row["decompressed_equal"]:
                    self.error(f"{case_id}: refused edit changed binary member {row['name']}")
        if case.get("format") == "DOCX":
            allowed = self.check_docx_target(source, default_output, admitted, case_id)
        elif case.get("format") == "XLSX":
            allowed = self.check_xlsx_target(source, default_output, case, case_id)
        elif case.get("format") == "PPTX":
            allowed = self.check_pptx_target(source, default_output, case, case_id)
        else:
            allowed = set()
            self.error(f"{case_id}: unsupported format {case.get('format')!r}")
        changed: list[dict[str, Any]] = []
        for row in comparison["members"]:
            if row["decompressed_equal"]:
                continue
            name = row["name"]
            if name not in allowed:
                self.error(f"{case_id}: decompressed member changed outside target: {name}")
            changed.append(row)
        for row in comparison["members"]:
            if row["xml"] and not row["canonical_equal"] and row["name"] not in allowed:
                self.error(f"{case_id}: canonical XML changed outside target: {row['name']}")
        result.update(
            {
                "source_equals_output": default_output["bytes"] == source["bytes"],
                "edit_admitted": admitted,
                "edit_outcome": case.get("edit_outcome"),
                "allowed_changed_members": sorted(allowed),
                "changed_members": changed,
                "comparison": comparison,
                "policy_outputs": policy_results,
            }
        )
        return result

    def run(self, artifacts_root: Path) -> dict[str, Any]:
        manifest_path = artifacts_root / "manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        if manifest.get("schema_version") != 1:
            self.error("manifest schema_version is not 1")
        if manifest.get("kind") != EXPECTED_KIND:
            self.error(f"manifest kind is not {EXPECTED_KIND!r}")
        if manifest.get("generator") != EXPECTED_GENERATOR:
            self.error(f"manifest generator is not {EXPECTED_GENERATOR!r}")
        cases = manifest.get("cases")
        if not isinstance(cases, list):
            self.error("manifest cases is not a list")
            cases = []
        seen: set[tuple[str, str]] = set()
        for case in cases:
            key = (str(case.get("origin")), str(case.get("format")))
            if key in seen:
                self.error(f"duplicate manifest case for {key}")
            seen.add(key)
            if key[0] not in {"generated-harness-corpus", "caller-named-real-file"} or key[1] not in FORMATS:
                self.error(f"unexpected case identity {key}")
        expected = {
            ("generated-harness-corpus", fmt) for fmt in FORMATS
        } | {("caller-named-real-file", fmt) for fmt in FORMATS}
        if seen != expected:
            self.error(f"case scope mismatch: expected {sorted(expected)}, got {sorted(seen)}")
        case_results = []
        for case in cases:
            error_start = len(self.errors)
            result = self.audit_case(manifest, case, artifacts_root)
            case_errors = self.errors[error_start:]
            result["errors"] = list(case_errors)
            result["ok"] = not case_errors
            case_results.append(result)
        return {
            "schema": SCHEMA,
            "ok": not self.errors,
            "artifact_directory": str(artifacts_root),
            "manifest": str(manifest_path),
            "cases": case_results,
            "errors": self.errors,
        }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, help="ordinary-save export directory")
    parser.add_argument("--report", type=Path, help="write the JSON audit report here")
    parser.add_argument(
        "--check",
        type=Path,
        help="replay an existing immutable audit report and require an exact JSON match",
    )
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path(__file__).resolve().parents[4],
        help="repository root used to resolve relative real input paths",
    )
    args = parser.parse_args(argv)
    if bool(args.report) and bool(args.check):
        parser.error("--report and --check are mutually exclusive")
    if not args.artifacts and not args.check:
        parser.error("one of --artifacts or --check is required")
    if args.check:
        try:
            retained = json.loads(args.check.read_text(encoding="utf-8"))
            artifacts_value = retained.get("artifact_directory")
            if not artifacts_value:
                raise ValueError("retained report has no artifact_directory")
            artifacts_root = Path(artifacts_value).resolve()
        except Exception as exc:
            print(json.dumps({"ok": False, "errors": [f"cannot read retained report: {exc}"]}, indent=2))
            return 1
    else:
        artifacts_root = args.artifacts.resolve()
    audit = Audit(args.repo_root.resolve())
    try:
        report = audit.run(artifacts_root)
    except Exception as exc:
        report = {
            "schema": SCHEMA,
            "ok": False,
            "artifact_directory": str(artifacts_root),
            "errors": [f"fatal: {exc}"],
        }
    if args.check:
        if report != retained:
            print(json.dumps({"ok": False, "errors": ["retained audit does not exactly replay"]}, indent=2))
            return 1
        print(json.dumps({"ok": True, "errors": []}, indent=2))
        return 0
    if args.report:
        if args.report.exists():
            print(json.dumps({"ok": False, "errors": [f"refusing to overwrite existing report {args.report}"]}, indent=2))
            return 1
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"ok": report.get("ok", False), "errors": report.get("errors", [])}, indent=2))
    return 0 if report.get("ok", False) else 1


if __name__ == "__main__":
    raise SystemExit(main())
