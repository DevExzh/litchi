#!/usr/bin/env python3
"""Independent semantic and preservation oracle for change 0778.

The exporter writes one manifest and four output archives for every corpus:
``default``, ``full``, ``file-only``, and ``no-sync``.  This checker reads the
archives with only :mod:`zipfile` and :mod:`xml.etree.ElementTree`; it does not
open them through Litchi or trust a semantic projection emitted by the Rust
runner.

The frozen Rust exporter emits an ``ordinary-save-artifact-export`` manifest
with numeric ``schema_version`` and a ``cases`` array.  Each case has a
``source_archive`` record, ``policy_outputs`` records, a separate
``stream_output`` record, and an ``edit_target``.  This checker consumes that
manifest directly.  The report it writes uses a ``corpora`` array so the
capture/analyze admission reader can bind the result without treating the
Rust manifest as a semantic oracle.

The preservation closure is defined here from the package format and the
selected edit target.  It is never read from the exporter.  A DOCX edit owns
``word/document.xml``; a PPTX edit owns the selected slide; and an XLSX edit
owns the selected worksheet, workbook, workbook relationships, and
content-type table.  ``xl/sharedStrings.xml`` is included only when the
selected source cell is a shared string.  An adjacent ``.rels`` part is
included only as a scoped lexical allowance and its relationship graph must
remain exact, except for the exact workbook ``calcChain`` relationship
removal paired with removal of the known ``calcChain.xml`` member and content
type override.  Workbook structure outside the exact dirty-calculation flags
is independently compared.  The only additional content-type topology
rewrite admitted here replaces a source printer-settings ``.bin`` Default
with one Override for each retained ``.bin`` member; effective member types
and every other declaration must remain equal.  Every other decoded ZIP
member must remain byte-identical.
This deliberately scopes the claim to members the ordinary editor can own and
records the actual changed-member set.

For DOCX, the independent projection checks that the original paragraph text
is an unchanged prefix and that the edit marker occurs exactly once.  For
XLSX, it resolves workbook sheet relationships itself, decodes shared,
inline, boolean, numeric, formula, and ordinary string cell values, and
compares every cell except the manifest-selected target.  For PPTX, it
resolves slide XML directly, identifies the selected shape by its frozen
zero-based index, and compares all other shape text while requiring the
target text to equal the marker.

This file intentionally contains no generated fixtures or permissive fallback
oracle.  Running it before the exporter has produced real artifacts is an
error.  The report also binds its manifest and this runner by SHA-256, and
checks the frozen edit description/target tuple for every corpus before any
archive semantic check runs.
"""

from __future__ import annotations

import argparse
from collections import OrderedDict
from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
import posixpath
import re
import sys
from typing import Any, Iterable, Mapping, Sequence
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parent
PLAN_PATH = ROOT / "plan.json"
SCHEMA = "litchi-0778-export-v1"
EXPORT_SCHEMA_VERSION = 1
EXPORT_KIND = "ordinary-save-artifact-export"
EDIT_MARKER = "litchi-perf-0638-ordinary-save"
POLICIES = ("default", "full", "file-only", "no-sync")
EXPECTED_CASES = OrderedDict(
    (
        ("generated-docx", ("docx", True)),
        ("generated-xlsx", ("xlsx", True)),
        ("generated-pptx", ("pptx", True)),
        ("numbered-list", ("docx", True)),
        ("alt-chunk-header", ("docx", False)),
        ("conditional-formatting", ("xlsx", True)),
        ("slide-section", ("pptx", True)),
    )
)
EXPECTED_EDIT_BINDINGS = {
    "generated-docx": {
        "description": 'Package::document_mut().add_paragraph_with_text("litchi-perf-0638-ordinary-save")',
        "target": {"main_part": "word/document.xml"},
    },
    "generated-xlsx": {
        "description": 'Workbook::edit().sheet("Sheet1").set("A1", "litchi-perf-0638-ordinary-save") and commit',
        "target": {"main_part": "xl/workbook.xml", "xlsx_sheet": "Sheet1", "xlsx_address": "A1"},
    },
    "generated-pptx": {
        "description": 'Package::opened_presentation_transaction().set_shape_text(0, 0, "litchi-perf-0638-ordinary-save") and apply',
        "target": {"main_part": "ppt/presentation.xml", "pptx_slide": 0, "pptx_shape": 0},
    },
    "numbered-list": {
        "description": 'Package::document_mut().add_paragraph_with_text("litchi-perf-0638-ordinary-save")',
        "target": {"main_part": "word/document.xml"},
    },
    "alt-chunk-header": {
        "description": 'Package::document_mut().add_paragraph_with_text("litchi-perf-0638-ordinary-save")',
        "target": {"main_part": "word/document.xml"},
    },
    "conditional-formatting": {
        "description": 'Workbook::edit().sheet("Home").set("A1", "litchi-perf-0638-ordinary-save") and commit',
        "target": {"main_part": "xl/workbook.xml", "xlsx_sheet": "Home", "xlsx_address": "A1"},
    },
    "slide-section": {
        "description": 'Package::opened_presentation_transaction().set_shape_text(1, 1, "litchi-perf-0638-ordinary-save") and apply',
        "target": {"main_part": "ppt/presentation.xml", "pptx_slide": 1, "pptx_shape": 1},
    },
}
SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")
SHEET_REF_RE = re.compile(r"^\$?([A-Za-z_][^!$]*)!?\$?([A-Za-z]{1,3})\$?([0-9]+)$")
CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_REL_STRICT = "http://purl.oclc.org/ooxml/package/relationships"
NS_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_DOC_REL_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships"
NS_WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_SHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P_STRICT = "http://purl.oclc.org/ooxml/presentationml/main"
NS_A_STRICT = "http://purl.oclc.org/ooxml/drawingml/main"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
CALC_CHAIN_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
)
PRINTER_SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
)
CALC_CHAIN_RELATION_TYPES = {
    f"{NS_DOC_REL}/calcChain",
    f"{NS_DOC_REL_STRICT}/calcChain",
}
EXPECTED_CALC_PR = {
    "calcId": "0",
    "fullCalcOnLoad": "true",
    "calcCompleted": "false",
    "calcOnSave": "true",
    "forceFullCalc": "true",
}


class OracleError(RuntimeError):
    """A fail-closed oracle violation."""


def fail(message: str) -> "NoReturn":
    raise OracleError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def require_sha(value: Any, where: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{where} must be a 64-character SHA-256 digest")
    return value.lower()


def text(value: Any, where: str) -> str:
    require(isinstance(value, str) and value != "", f"{where} must be non-empty text")
    return value


def mapping(value: Any, where: str) -> Mapping[str, Any]:
    require(isinstance(value, Mapping), f"{where} must be an object")
    return value


def list_value(value: Any, where: str) -> list[Any]:
    require(isinstance(value, list), f"{where} must be an array")
    return value


def first(mapping_value: Mapping[str, Any], keys: Sequence[str], where: str,
          *, allow_none: bool = False) -> Any:
    for key in keys:
        if key in mapping_value:
            value = mapping_value[key]
            if value is None and allow_none:
                return None
            return value
    fail(f"{where} is missing (accepted keys: {', '.join(keys)})")


def optional(mapping_value: Mapping[str, Any], keys: Sequence[str]) -> Any:
    for key in keys:
        if key in mapping_value:
            return mapping_value[key]
    return None


def resolve_path(raw: Any, base: Path, where: str, *, require_under_base: bool = False) -> Path:
    require(isinstance(raw, str) and raw != "", f"{where} must be a path string")
    candidate = Path(raw)
    if not candidate.is_absolute():
        candidate = base / candidate
    require(candidate.exists(), f"{where} does not exist: {candidate}")
    require(not candidate.is_symlink(), f"{where} must not be a symbolic link: {candidate}")
    try:
        path = candidate.resolve(strict=True)
    except FileNotFoundError:
        fail(f"{where} does not exist: {candidate}")
    require(path.is_file(), f"{where} is not a regular file: {path}")
    if require_under_base:
        base_resolved = base.resolve()
        try:
            path.relative_to(base_resolved)
        except ValueError:
            fail(f"{where} escapes the current artifact directory: {path}")
    return path


def safe_member_name(name: str) -> str:
    require(name != "" and "\\" not in name and not name.startswith("/"),
            f"unsafe ZIP member name {name!r}")
    parts = name.split("/")
    require(".." not in parts and "." not in parts,
            f"unsafe ZIP member path {name!r}")
    return name


@dataclass(frozen=True)
class ZipArchive:
    path: Path
    raw: bytes
    members: Mapping[str, bytes]

    @classmethod
    def read(cls, path: Path, where: str) -> "ZipArchive":
        raw = path.read_bytes()
        try:
            archive = zipfile.ZipFile(path)
        except (OSError, zipfile.BadZipFile) as exc:
            fail(f"{where} is not a readable ZIP archive: {exc}")
        try:
            infos = archive.infolist()
            names: list[str] = []
            members: dict[str, bytes] = {}
            for info in infos:
                name = safe_member_name(info.filename)
                require(name not in members, f"{where} contains duplicate ZIP member {name!r}")
                require(info.is_dir() is False, f"{where} contains a directory member {name!r}")
                # Reject Unix symlink entries rather than following an archive link.
                mode = (info.external_attr >> 16) & 0o170000
                require(mode != 0o120000,
                        f"{where} contains a symbolic-link ZIP member {name!r}")
                try:
                    members[name] = archive.read(info)
                except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
                    fail(f"cannot read {where} member {name!r}: {exc}")
                names.append(name)
            require(names, f"{where} contains no ZIP members")
        finally:
            archive.close()
        return cls(path=path, raw=raw, members=members)

    def member(self, name: str, where: str) -> bytes:
        safe_member_name(name)
        require(name in self.members, f"{where} is missing ZIP member {name!r}")
        return self.members[name]


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def namespace(tag: str) -> str:
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


def xml_root(data: bytes, where: str) -> ET.Element:
    try:
        # XMLParser does not resolve external entities.  ElementTree rejects
        # malformed declarations rather than silently projecting partial XML.
        return ET.fromstring(data)
    except (ET.ParseError, ValueError) as exc:
        fail(f"{where} is not independently parseable XML: {exc}")


def node_projection(node: ET.Element) -> Any:
    attrs = tuple(sorted((key, value) for key, value in node.attrib.items()))
    children = tuple(node_projection(child) for child in list(node))
    return (
        node.tag,
        attrs,
        node.text if node.text is not None else "",
        children,
    )


def xml_projection(data: bytes, where: str) -> Any:
    return node_projection(xml_root(data, where))


def direct_text(node: ET.Element) -> str:
    return "".join(part for part in node.itertext())


def count_occurrences(values: Iterable[str], marker: str) -> int:
    return sum(value.count(marker) for value in values)


def changed_member_names(source: ZipArchive, output: ZipArchive,
                         allowed_member_set_changes: frozenset[str]) -> set[str]:
    """Return changed members, allowing only the frozen closure to add/remove.

    The ordinary-save serializer may remove a known calculation-chain part
    while rewriting an XLSX.  That one scoped removal is recorded as a
    changed member; every other archive member must remain present.  A
    closure entry may also be newly emitted, but arbitrary archive growth or
    loss remains a hard failure.
    """
    source_names = set(source.members)
    output_names = set(output.members)
    removed = source_names - output_names
    added = output_names - source_names
    member_set_changes = removed | added
    require(member_set_changes <= allowed_member_set_changes,
            "output ZIP member set differs outside the fixed mutation closure: "
            f"{sorted(member_set_changes - allowed_member_set_changes)}")
    common = source_names & output_names
    return (removed | added | {
        name for name in common
        if source.members[name] != output.members[name]
    })


def manifest_source(case: Mapping[str, Any], base: Path, where: str) -> tuple[Path, str]:
    raw = first(case, ("source", "input", "source_archive"), where)
    if isinstance(raw, Mapping):
        source = raw
        path_value = first(source, ("path", "archive", "file"), f"{where}.source")
        digest_value = first(source, ("sha256", "source_sha256", "archive_sha256"),
                             f"{where}.source")
    else:
        path_value = raw
        digest_value = first(case, ("source_sha256", "input_sha256", "archive_sha256"), where)
    path = resolve_path(path_value, base, f"{where}.source.path", require_under_base=True)
    digest = require_sha(digest_value, f"{where}.source.sha256")
    actual = sha256_file(path)
    require(actual == digest,
            f"{where} source SHA-256 changed: manifest {digest}, actual {actual}")
    if isinstance(raw, Mapping):
        declared_bytes = optional(raw, ("bytes", "source_archive_bytes", "archive_bytes"))
        if declared_bytes is not None:
            require(type(declared_bytes) is int and declared_bytes == path.stat().st_size,
                    f"{where} source byte count is not bound to source bytes")
    return path, digest


def manifest_edit(case: Mapping[str, Any], where: str) -> tuple[bool, str, Mapping[str, Any]]:
    raw = optional(case, ("edit", "operation"))
    edit = mapping(raw, f"{where}.edit") if raw is not None else case
    admitted_value = optional(edit, ("admitted", "expected_edit_admitted", "accepted"))
    if admitted_value is None:
        admitted_value = optional(case, ("expected_edit_admitted", "edit_admitted"))
    require(type(admitted_value) is bool, f"{where}.edit.admitted must be boolean")
    marker_value = optional(edit, ("marker", "append_marker", "expected_marker"))
    if marker_value is None:
        marker_value = optional(case, ("marker", "append_marker", "expected_marker"))
    if marker_value is None:
        description = optional(edit, ("edit_description", "description"))
        if description is None:
            description = optional(case, ("edit_description", "description"))
        if isinstance(description, str):
            # The artifact exporter records the literal marker in its edit
            # description.  Only the one frozen marker is accepted; a
            # free-form description cannot silently choose the edit marker.
            candidates = re.findall(r"['\"](litchi-perf-[^'\"]+)['\"]", description)
            if candidates == [EDIT_MARKER]:
                marker_value = candidates[0]
    if admitted_value:
        marker = text(marker_value, f"{where}.edit.marker")
        require(marker == EDIT_MARKER,
                f"{where}.edit.marker must equal the frozen ordinary-save marker")
    else:
        # Refused saves are checked byte-for-byte, so no mutation closure or
        # marker is needed.  If the exporter records one, it must still be the
        # frozen marker rather than a caller-selected edit.
        marker = marker_value if isinstance(marker_value, str) else ""
        if marker:
            require(marker == EDIT_MARKER,
                    f"{where}.edit.marker must equal the frozen ordinary-save marker")
    return admitted_value, marker, edit


def output_records(case: Mapping[str, Any], base: Path, where: str) -> dict[str, tuple[Path, str | None, str | None, int | None]]:
    raw = first(case, ("outputs", "artifacts", "published", "policy_outputs"), where)
    records: dict[str, Mapping[str, Any] | str] = {}
    if isinstance(raw, Mapping):
        for policy, record in raw.items():
            records[str(policy)] = record
    elif isinstance(raw, list):
        for index, record in enumerate(raw):
            if isinstance(record, str):
                fail(f"{where}.outputs[{index}] lacks policy")
            obj = mapping(record, f"{where}.outputs[{index}]")
            policy = first(obj, ("policy", "durability", "level"),
                           f"{where}.outputs[{index}]")
            policy = text(policy, f"{where}.outputs[{index}].policy")
            require(policy not in records, f"{where} has duplicate output policy {policy!r}")
            records[policy] = obj
    else:
        fail(f"{where}.outputs must be an object or array")

    # The frozen ordinary-save exporter keeps the sequential sink outside the
    # four durability records.  Normalize it into the same internal shape so
    # the exact-parity check cannot accidentally omit it.
    stream = optional(case, ("stream_output", "stream"))
    if stream is not None:
        require("stream" not in records, f"{where} declares stream output twice")
        records["stream"] = stream

    # The Rust artifact exporter may retain a fifth sequential-sink artifact.
    # It is checked with the same byte/parity contract, but it is not one of
    # the four durability policies named by the frozen plan.
    extra = set(records) - set(POLICIES) - {"stream"}
    require(not extra,
            f"{where}.outputs contains unsupported policies {sorted(extra)}")
    require(set(POLICIES) <= set(records),
            f"{where}.outputs must contain all policies {POLICIES}; got {tuple(records)}")
    require("stream" in records,
            f"{where}.outputs must retain the frozen sequential-output control")
    result: dict[str, tuple[Path, str | None, str | None, int | None]] = {}
    policy_names = POLICIES + (("stream",) if "stream" in records else ())
    for policy in policy_names:
        record = records[policy]
        expected_sha: str | None = None
        expected_bytes: int | None = None
        if isinstance(record, str):
            path_value = record
            status = None
        else:
            obj = mapping(record, f"{where}.outputs[{policy}]")
            path_value = first(obj, ("path", "archive", "file", "output"),
                               f"{where}.outputs[{policy}]")
            if isinstance(path_value, Mapping):
                path_value = first(path_value, ("path", "archive", "file"),
                                   f"{where}.outputs[{policy}].output")
            raw_status = optional(obj, ("status", "result", "outcome"))
            status = str(raw_status) if raw_status is not None else None
            expected_sha = optional(obj, ("sha256", "output_sha256", "archive_sha256"))
            if expected_sha is not None:
                expected_sha = require_sha(expected_sha, f"{where}.outputs[{policy}].sha256")
            raw_bytes = optional(obj, ("bytes", "output_bytes", "archive_bytes"))
            if raw_bytes is not None:
                require(type(raw_bytes) is int and raw_bytes >= 0,
                        f"{where}.outputs[{policy}].bytes must be a non-negative integer")
                expected_bytes = raw_bytes
        path = resolve_path(path_value, base, f"{where}.outputs[{policy}].path",
                           require_under_base=True)
        result[policy] = (path, status, expected_sha, expected_bytes)
    return result


def status_is_success(status: str) -> bool:
    return status.lower() in {"ok", "success", "accepted", "admitted", "published", "complete"}


def status_is_refusal(status: str) -> bool:
    return status.lower() in {"refused", "rejected", "no-edit", "no_edit", "ineligible", "error"}


def norm_cell_ref(value: Any, where: str) -> str:
    value = text(value, where).upper().replace("$", "")
    match = CELL_RE.fullmatch(value)
    require(match is not None, f"{where} is not an A1 cell reference: {value!r}")
    return f"{match.group(1)}{int(match.group(2))}"


def target_value(target: Mapping[str, Any], keys: Sequence[str], where: str) -> Any:
    return first(target, keys, where)


def docx_paragraph_projection(data: bytes, where: str) -> list[str]:
    root = xml_root(data, where)
    paragraphs = [node for node in root.iter() if node.tag == f"{{{NS_WORD}}}p"]
    return [direct_text(paragraph) for paragraph in paragraphs]


def marker_append_projection(source: Sequence[str], output: Sequence[str], marker: str,
                            where: str) -> dict[str, Any]:
    require(marker != "", f"{where} marker must be non-empty")
    source_values = list(source)
    output_values = list(output)
    require(count_occurrences(source_values, marker) == 0,
            f"{where} source already contains the edit marker")
    require(count_occurrences(output_values, marker) == 1,
            f"{where} output must contain the edit marker exactly once")

    # An append may be represented by a new paragraph or by the existing
    # final paragraph receiving the marker.  Both preserve the original
    # paragraph sequence and leave only the explicitly selected text suffix.
    if output_values[: len(source_values)] == source_values and len(output_values) == len(source_values) + 1:
        changed = {len(source_values)}
        require(output_values[-1] == marker,
                f"{where} appended paragraph must equal the marker")
    elif len(output_values) == len(source_values):
        differences = [
            index for index, (before, after) in enumerate(zip(source_values, output_values))
            if before != after
        ]
        require(len(differences) == 1, f"{where} must change exactly one paragraph")
        index = differences[0]
        require(output_values[index] == source_values[index] + marker,
                f"{where} changed paragraph must preserve its source text and append marker")
        require(output_values[:index] == source_values[:index],
                f"{where} paragraph prefix changed before marker")
        require(output_values[index + 1 :] == source_values[index + 1 :],
                f"{where} paragraph suffix changed after marker")
        changed = {index}
    else:
        fail(f"{where} paragraph count changed in an unsupported way")
    return {
        "source_paragraphs": source_values,
        "output_paragraphs": output_values,
        "changed_paragraph_indexes": sorted(changed),
        "marker": marker,
    }


def check_docx(source: ZipArchive, output: ZipArchive, marker: str,
               target: Mapping[str, Any], where: str) -> dict[str, Any]:
    member_value = optional(target, ("member", "part", "xml_member"))
    member = member_value if isinstance(member_value, str) else "word/document.xml"
    safe_member_name(member)
    source_xml = source.member(member, f"{where}.source")
    output_xml = output.member(member, f"{where}.output")
    source_projection = docx_paragraph_projection(source_xml, f"{where}.source.{member}")
    output_projection = docx_paragraph_projection(output_xml, f"{where}.output.{member}")
    return {
        "format": "docx",
        "target_member": member,
        "paragraphs": marker_append_projection(
            source_projection, output_projection, marker, f"{where}.paragraphs"
        ),
        "target_xml_projection": xml_projection(output_xml, f"{where}.output.{member}"),
    }


def rels_target(source_part: str, target: str) -> str:
    directory = posixpath.dirname(source_part)
    joined = posixpath.normpath(posixpath.join(directory, target))
    require(not joined.startswith("../") and joined != ".." and not joined.startswith("/"),
            f"relationship target escapes package: {source_part!r} -> {target!r}")
    return joined


def relationship_map(data: bytes, where: str) -> dict[str, tuple[str, str, str]]:
    root = xml_root(data, where)
    require(local_name(root.tag) == "Relationships" and namespace(root.tag) in {NS_REL, NS_REL_STRICT},
            f"{where} has an invalid relationships root")
    result: dict[str, tuple[str, str, str]] = {}
    for rel in root:
        require(
            local_name(rel.tag) == "Relationship"
            and namespace(rel.tag) in {NS_REL, NS_REL_STRICT},
            f"{where} contains an unknown relationship element",
        )
        rid = rel.get("Id")
        target = rel.get("Target")
        kind = rel.get("Type")
        mode = rel.get("TargetMode", "Internal")
        require(mode in {"Internal", "External"},
                f"{where} relationship has invalid TargetMode {mode!r}")
        require(rid and target and kind, f"{where} relationship lacks Id, Target, or Type")
        require(rid not in result, f"{where} contains duplicate relationship id {rid!r}")
        result[rid] = (kind, target, mode)
    return result


def workbook_sheets(archive: ZipArchive, where: str) -> dict[str, str]:
    workbook_member = "xl/workbook.xml"
    rels_member = "xl/_rels/workbook.xml.rels"
    workbook = xml_root(archive.member(workbook_member, where), f"{where}.{workbook_member}")
    rels = relationship_map(archive.member(rels_member, where), f"{where}.{rels_member}")
    result: dict[str, str] = {}
    sheets = workbook.find(f"{{{NS_SHEET}}}sheets")
    require(sheets is not None, f"{where} workbook has no sheets element")
    for sheet in sheets:
        require(sheet.tag == f"{{{NS_SHEET}}}sheet", f"{where} has an unknown sheet child")
        name = sheet.get("name")
        rid = sheet.get(f"{{{NS_DOC_REL}}}id") or sheet.get(f"{{{NS_DOC_REL_STRICT}}}id")
        require(name and rid and rid in rels, f"{where} sheet has no resolvable relationship")
        kind, target, mode = rels[rid]
        require(mode == "Internal", f"{where} worksheet relationship must be internal")
        require(kind.endswith("/worksheet"), f"{where} sheet {name!r} does not target a worksheet")
        member = rels_target(workbook_member, target)
        require(member in archive.members, f"{where} sheet {name!r} target {member!r} is absent")
        require(name not in result, f"{where} has duplicate sheet name {name!r}")
        result[name] = member
    require(result, f"{where} workbook has no worksheets")
    return result


def workbook_without_calc_pr(data: bytes, where: str) -> tuple[Any, dict[str, str] | None]:
    """Project workbook structure while returning its direct calculation flag."""
    root = xml_root(data, where)
    require(root.tag == f"{{{NS_SHEET}}}workbook",
            f"{where} is not a SpreadsheetML workbook")
    calc_nodes = [child for child in root if local_name(child.tag) == "calcPr"]
    require(len(calc_nodes) <= 1, f"{where} contains duplicate calcPr elements")
    calc_attrs = dict(calc_nodes[0].attrib) if calc_nodes else None
    projection = (
        root.tag,
        tuple(sorted(root.attrib.items())),
        tuple(node_projection(child) for child in root
              if local_name(child.tag) != "calcPr"),
    )
    return projection, calc_attrs


def content_types_table(
    data: bytes, where: str
) -> tuple[dict[str, str], dict[str, str], list[tuple[str, tuple[tuple[str, str], ...]]]]:
    """Decode defaults, overrides, and an order-independent entry projection."""
    root = xml_root(data, where)
    require(root.tag == f"{{{NS_CT}}}Types", f"{where} has an invalid content-types root")
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    entries: list[tuple[str, tuple[tuple[str, str], ...]]] = []
    for child in root:
        require(namespace(child.tag) == NS_CT and local_name(child.tag) in {"Default", "Override"},
                f"{where} contains an unknown content-types child")
        if local_name(child.tag) == "Default":
            extension = child.get("Extension")
            content_type = child.get("ContentType")
            require(extension and content_type,
                    f"{where} has an incomplete content-type Default")
            extension = extension.lower()
            require(extension not in defaults,
                    f"{where} contains duplicate content-type Default {extension!r}")
            defaults[extension] = content_type
        else:
            part_name = child.get("PartName")
            content_type = child.get("ContentType")
            require(part_name and content_type and part_name.startswith("/"),
                    f"{where} has an incomplete content-type Override")
            require(part_name not in overrides,
                    f"{where} contains duplicate content-type Override {part_name!r}")
            overrides[part_name] = content_type
        entries.append((local_name(child.tag), tuple(sorted(child.attrib.items()))))
    return defaults, overrides, sorted(entries)


def content_types_projection(data: bytes, where: str) -> list[tuple[str, tuple[tuple[str, str], ...]]]:
    """Return the order-independent decoded content-type table."""
    return content_types_table(data, where)[2]


def effective_content_type(member: str, table: tuple[dict[str, str], dict[str, str], Any]) -> str | None:
    """Resolve one package member using OPC override-then-default precedence."""
    defaults, overrides, _ = table
    override = overrides.get("/" + member)
    if override is not None:
        return override
    extension = posixpath.splitext(member.rsplit("/", 1)[-1])[1].lstrip(".").lower()
    return defaults.get(extension) if extension else None


def check_xlsx_owned_parts(source: ZipArchive, output: ZipArchive, where: str) -> dict[str, Any]:
    """Check the finite workbook dirty-calculation mutation contract."""
    source_workbook, source_calc = workbook_without_calc_pr(
        source.member("xl/workbook.xml", where), f"{where}.source.xl/workbook.xml"
    )
    output_workbook, output_calc = workbook_without_calc_pr(
        output.member("xl/workbook.xml", where), f"{where}.output.xl/workbook.xml"
    )
    require(source_workbook == output_workbook,
            f"{where} workbook structure or attributes changed outside calcPr")
    require(output_calc == EXPECTED_CALC_PR,
            f"{where} output calcPr does not match the exact dirty-calculation contract")

    source_content_type_table = content_types_table(
        source.member("[Content_Types].xml", where), f"{where}.source.[Content_Types].xml"
    )
    output_content_type_table = content_types_table(
        output.member("[Content_Types].xml", where), f"{where}.output.[Content_Types].xml"
    )
    source_content_types = source_content_type_table[2]
    output_content_types = output_content_type_table[2]
    calc_override = ("Override", (
        ("ContentType", CALC_CHAIN_CONTENT_TYPE),
        ("PartName", "/xl/calcChain.xml"),
    ))
    printer_default = ("Default", (
        ("ContentType", PRINTER_SETTINGS_CONTENT_TYPE),
        ("Extension", "bin"),
    ))
    expected_content_types = list(source_content_types)
    topology_rewrites: list[str] = []
    if "xl/calcChain.xml" in source.members:
        require("xl/calcChain.xml" not in output.members,
                f"{where} output retained the removed calcChain member")
        require(source_content_types.count(calc_override) == 1,
                f"{where} source calcChain member lacks its exact content-type override")
        expected_content_types.remove(calc_override)
        topology_rewrites.append("calcChain-override-removal")
    if source_content_types.count(printer_default) == 1:
        retained_members = sorted(
            (set(source.members) | set(output.members)) - {"xl/calcChain.xml"}
        )
        retained_bin_members = [
            member for member in retained_members
            if posixpath.splitext(member.rsplit("/", 1)[-1])[1].lower() == ".bin"
        ]
        require(retained_bin_members,
                f"{where} source printer-settings Default has no retained .bin members")
        for member in retained_bin_members:
            require(effective_content_type(member, source_content_type_table) == PRINTER_SETTINGS_CONTENT_TYPE,
                    f"{where} source .bin member {member!r} has an unexpected effective content type")
            expected_content_types.append(("Override", (
                ("ContentType", PRINTER_SETTINGS_CONTENT_TYPE),
                ("PartName", "/" + member),
            )))
        expected_content_types.remove(printer_default)
        topology_rewrites.append("printer-settings-default-to-retained-overrides")
    require(output_content_types == sorted(expected_content_types),
            f"{where} content-type table changed outside the exact owned normalization")
    common_members = (set(source.members) & set(output.members)) - {"[Content_Types].xml"}
    common_members.discard("xl/calcChain.xml")
    for member in sorted(common_members):
        require(effective_content_type(member, source_content_type_table) ==
                effective_content_type(member, output_content_type_table),
                f"{where} effective content type changed for retained member {member!r}")
    return {
        "workbook_structure_equal": True,
        "source_calcPr": source_calc,
        "output_calcPr": output_calc,
        "calcPr_contract": dict(EXPECTED_CALC_PR),
        "content_types_effective_semantics_equal": True,
        "content_type_topology_rewrites": topology_rewrites,
        "printer_settings_default_removed": source_content_types.count(printer_default) == 1,
        "content_type_effective_members_checked": len(common_members),
        "calcChain_source_present": "xl/calcChain.xml" in source.members,
    }


def shared_strings(archive: ZipArchive, where: str) -> list[str]:
    member = "xl/sharedStrings.xml"
    if member not in archive.members:
        return []
    root = xml_root(archive.members[member], f"{where}.{member}")
    result: list[str] = []
    for item in root.findall(f"{{{NS_SHEET}}}si"):
        result.append(direct_text(item))
    return result


def cell_value(cell: ET.Element, strings: Sequence[str], where: str) -> tuple[str, str, str | None]:
    kind = cell.get("t", "")
    value_node = cell.find(f"{{{NS_SHEET}}}v")
    formula_node = cell.find(f"{{{NS_SHEET}}}f")
    formula = formula_node.text if formula_node is not None and formula_node.text is not None else None
    if kind == "s":
        require(value_node is not None and value_node.text is not None, f"{where} shared cell has no value")
        try:
            index = int(value_node.text)
        except ValueError:
            fail(f"{where} shared-string index is not an integer")
        require(0 <= index < len(strings), f"{where} shared-string index out of range")
        return ("text", strings[index], formula)
    if kind == "inlineStr":
        inline = cell.find(f"{{{NS_SHEET}}}is")
        require(inline is not None, f"{where} inline string has no is element")
        return ("text", direct_text(inline), formula)
    if kind == "b":
        require(value_node is not None and value_node.text in {"0", "1"},
                f"{where} boolean cell has invalid value")
        return ("boolean", "true" if value_node.text == "1" else "false", formula)
    if kind == "e":
        require(value_node is not None, f"{where} error cell has no value")
        return ("error", value_node.text or "", formula)
    if kind == "str":
        require(value_node is not None, f"{where} formula string cell has no value")
        return ("text", value_node.text or "", formula)
    if kind in {"", "n"}:
        if value_node is None:
            value = ""
        else:
            value = value_node.text or ""
        # Preserve the decoded lexical scalar: the oracle intentionally does
        # not reinterpret dates or formulas as a different value type.
        return ("number", value, formula)
    fail(f"{where} uses unsupported cell type {kind!r}")


def xlsx_cells(archive: ZipArchive, where: str) -> dict[tuple[str, str], tuple[str, str, str | None]]:
    sheets = workbook_sheets(archive, where)
    strings = shared_strings(archive, where)
    result: dict[tuple[str, str], tuple[str, str, str | None]] = {}
    for sheet_name, member in sheets.items():
        root = xml_root(archive.members[member], f"{where}.{member}")
        sheet_data = root.find(f"{{{NS_SHEET}}}sheetData")
        require(sheet_data is not None, f"{where}.{member} has no sheetData")
        for cell in sheet_data.iter(f"{{{NS_SHEET}}}c"):
            reference = cell.get("r")
            require(reference, f"{where}.{member} cell has no reference")
            reference = norm_cell_ref(reference, f"{where}.{member}.cell")
            value_node = cell.find(f"{{{NS_SHEET}}}v")
            inline_node = cell.find(f"{{{NS_SHEET}}}is")
            formula_node = cell.find(f"{{{NS_SHEET}}}f")
            if (value_node is None or value_node.text in {None, ""}) and inline_node is None and formula_node is None:
                # A styled empty cell has no decoded value and is outside the
                # cell-value projection required by this oracle.
                continue
            key = (sheet_name, reference)
            require(key not in result, f"{where} contains duplicate cell {sheet_name}!{reference}")
            result[key] = cell_value(cell, strings, f"{where}.{member}!{reference}")
    return result


def xlsx_source_cell_kind(archive: ZipArchive, sheet: str, reference: str,
                          where: str) -> str | None:
    member = workbook_sheets(archive, where)[sheet]
    root = xml_root(archive.members[member], f"{where}.{member}")
    for cell in root.iter(f"{{{NS_SHEET}}}c"):
        if norm_cell_ref(cell.get("r"), f"{where}.{member}.cell") == reference:
            return cell.get("t", "")
    return None


def target_sheet_and_cell(target: Mapping[str, Any], sheets: Mapping[str, str], where: str) -> tuple[str, str]:
    sheet_value = optional(target, ("sheet", "sheet_name", "worksheet", "xlsx_sheet"))
    cell_value_raw = optional(target, ("cell", "cell_ref", "reference", "target_cell", "xlsx_address"))
    if sheet_value is None and isinstance(optional(target, ("sheet_ref",)), str):
        match = SHEET_REF_RE.fullmatch(str(target["sheet_ref"]))
        if match:
            sheet_value = match.group(1)
            cell_value_raw = f"{match.group(2)}{match.group(3)}"
    require(isinstance(sheet_value, str) and sheet_value != "", f"{where}.target.sheet is required")
    require(sheet_value in sheets, f"{where}.target.sheet {sheet_value!r} is not present")
    cell = norm_cell_ref(cell_value_raw, f"{where}.target.cell")
    return sheet_value, cell


def relationship_member(part: str) -> str:
    directory, basename = posixpath.split(part)
    return posixpath.join(directory, "_rels", basename + ".rels")


def fixed_mutation_closure(case_format: str, source: ZipArchive,
                           target: Mapping[str, Any], where: str) -> frozenset[str]:
    """Return the source-derived member closure owned by the frozen edit."""
    if case_format == "docx":
        main = optional(target, ("main_part", "member", "part", "xml_member"))
        require(main == "word/document.xml",
                f"{where}.target.main_part must be word/document.xml")
        owned = {"word/document.xml"}
    elif case_format == "xlsx":
        main = optional(target, ("main_part", "member", "part", "xml_member"))
        require(main == "xl/workbook.xml",
                f"{where}.target.main_part must be xl/workbook.xml")
        sheets = workbook_sheets(source, f"{where}.source")
        sheet_name, address = target_sheet_and_cell(target, sheets, where)
        # The Rust workbook writer rewrites the workbook relationship graph
        # and content-type table as part of a normal XLSX save.  Keep this
        # explicit and finite: these well-known parts are admitted only when
        # they are present in the source archive, and the adjacent rels graph
        # plus effective content-type map are checked semantically below.
        owned = {
            name for name in (
                "[Content_Types].xml",
                "xl/_rels/workbook.xml.rels",
                "xl/workbook.xml",
                sheets[sheet_name],
            ) if name in source.members
        }
        # Some workbooks carry a calculation chain which the serializer
        # intentionally drops when recalculating.  Its removal is the only
        # permitted source/output member-set difference.
        if "xl/calcChain.xml" in source.members:
            owned.add("xl/calcChain.xml")
        # A shared-string index is a decoded dependency of the selected cell.
        # It is authorized only when the source target actually uses one;
        # adding a broad sharedStrings exemption would hide unrelated edits.
        if xlsx_source_cell_kind(source, sheet_name, address, f"{where}.source") == "s":
            require("xl/sharedStrings.xml" in source.members,
                    f"{where}.source shared-string target lacks sharedStrings.xml")
            owned.add("xl/sharedStrings.xml")
    elif case_format == "pptx":
        main = optional(target, ("main_part", "member", "part", "xml_member"))
        require(main == "ppt/presentation.xml",
                f"{where}.target.main_part must be ppt/presentation.xml")
        owned = {slide_member_from_target(target, source, f"{where}.source")}
        shape_index = optional(target, ("pptx_shape", "shape_index", "target_shape_index"))
        require(isinstance(shape_index, int) and not isinstance(shape_index, bool) and shape_index >= 0,
                f"{where}.target.pptx_shape must be a non-negative zero-based index")
    else:  # pragma: no cover - frozen cases exhaust formats.
        fail(f"{where} unsupported format {case_format!r}")

    # Relationship parts adjacent to an owned part are an explicitly scoped
    # lexical-rewrite allowance.  Their decoded relationship graph must still
    # be identical; no relationship topology change is authorized here.
    for part in tuple(owned):
        rels = relationship_member(part)
        if rels in source.members:
            owned.add(rels)
    return frozenset(owned)


def check_xlsx(source: ZipArchive, output: ZipArchive, marker: str,
               target: Mapping[str, Any], where: str) -> dict[str, Any]:
    source_sheets = workbook_sheets(source, f"{where}.source")
    output_sheets = workbook_sheets(output, f"{where}.output")
    require(list(source_sheets) == list(output_sheets),
            f"{where} worksheet names/order changed")
    sheet_name, cell = target_sheet_and_cell(target, output_sheets, where)
    source_cells = xlsx_cells(source, f"{where}.source")
    output_cells = xlsx_cells(output, f"{where}.output")
    target_key = (sheet_name, cell)
    require(target_key in output_cells, f"{where} target cell is absent from output")
    source_value = source_cells.get(target_key)
    output_value = output_cells[target_key]
    if source_value is not None:
        require(source_value[1].find(marker) < 0,
                f"{where} source target cell already contains the marker")
    require(output_value[0] == "text" and output_value[1] == marker,
            f"{where} target cell must decode to the marker exactly")
    require(sum(value[1] == marker for value in output_cells.values()) == 1,
            f"{where} output must contain the marker in exactly one cell")
    source_without = {key: value for key, value in source_cells.items() if key != target_key}
    output_without = {key: value for key, value in output_cells.items() if key != target_key}
    require(source_without == output_without,
            f"{where} non-target decoded cell values changed")
    return {
        "format": "xlsx",
        "owned_parts": check_xlsx_owned_parts(source, output, where),
        "target_sheet": sheet_name,
        "target_cell": cell,
        "source_target": source_value,
        "output_target": output_value,
        "cell_count": len(output_cells),
        "cells_sha256": sha256_bytes(json.dumps(
            sorted((sheet, ref, value) for (sheet, ref), value in output_cells.items()),
            ensure_ascii=False, separators=(",", ":"), default=list
        ).encode("utf-8")),
    }


def slide_member_from_target(target: Mapping[str, Any], archive: ZipArchive, where: str) -> str:
    raw = optional(target, ("member", "slide_member", "part"))
    if isinstance(raw, str):
        member = raw
    else:
        number = optional(target, ("slide", "slide_number", "slide_index"))
        zero_based = False
        if number is None:
            number = optional(target, ("pptx_slide",))
            zero_based = number is not None
        require(
            isinstance(number, int)
            and not isinstance(number, bool)
            and number >= (0 if zero_based else 1),
            f"{where}.target.slide must be a slide member or integer",
        )
        member = f"ppt/slides/slide{number + 1 if zero_based else number}.xml"
    safe_member_name(member)
    require(member in archive.members, f"{where} target slide member {member!r} is absent")
    return member


def ppt_shapes(data: bytes, where: str) -> list[tuple[str, str | None, str | None, str]]:
    root = xml_root(data, where)
    result: list[tuple[str, str | None, str | None, str]] = []
    shape_names = {"sp", "graphicFrame", "pic", "cxnSp", "grpSp", "contentPart"}

    def is_shape(node: ET.Element) -> bool:
        return namespace(node.tag) in {NS_P, NS_P_STRICT} and local_name(node.tag) in shape_names

    def direct_text_fragments(shape: ET.Element) -> list[str]:
        fragments: list[str] = []
        stack = list(reversed(list(shape)))
        while stack:
            node = stack.pop()
            if is_shape(node):
                # Nested group members are separate scene records.  Their text
                # must not be counted again on the group record itself.
                continue
            if namespace(node.tag) in {NS_A, NS_A_STRICT} and local_name(node.tag) == "t":
                fragments.append(node.text or "")
            stack.extend(reversed(list(node)))
        return fragments

    for index, shape in enumerate(root.iter()):
        if not is_shape(shape):
            continue
        props = None
        for candidate in shape.iter():
            if namespace(candidate.tag) in {NS_P, NS_P_STRICT} and local_name(candidate.tag) == "cNvPr":
                props = candidate
                break
        shape_id = props.get("id") if props is not None else None
        shape_name = props.get("name") if props is not None else None
        result.append((str(index), shape_id, shape_name, "".join(direct_text_fragments(shape))))
    return result


def check_pptx(source: ZipArchive, output: ZipArchive, marker: str,
               target: Mapping[str, Any], where: str) -> dict[str, Any]:
    source_member = slide_member_from_target(target, source, f"{where}.source")
    output_member = slide_member_from_target(target, output, f"{where}.output")
    require(source_member == output_member, f"{where} target slide member changed")
    source_shapes = ppt_shapes(source.members[source_member], f"{where}.source.{source_member}")
    output_shapes = ppt_shapes(output.members[output_member], f"{where}.output.{output_member}")
    shape_id = optional(target, ("shape_id", "target_shape_id"))
    shape_name = optional(target, ("shape_name", "target_shape_name"))
    shape_index = optional(target, ("shape_index", "target_shape_index", "pptx_shape"))
    if shape_id is None and shape_name is None and shape_index is None:
        nested = optional(target, ("shape",))
        if isinstance(nested, Mapping):
            shape_id = optional(nested, ("id", "shape_id"))
            shape_name = optional(nested, ("name", "shape_name"))
            shape_index = optional(nested, ("index", "shape_index"))
    matches: list[int] = []
    for index, (ordinal, candidate_id, candidate_name, _) in enumerate(source_shapes):
        if shape_id is not None and str(candidate_id) == str(shape_id):
            matches.append(index)
        elif shape_name is not None and candidate_name == str(shape_name):
            matches.append(index)
        elif shape_index is not None and index == shape_index:
            matches.append(index)
    require(len(matches) == 1, f"{where} target shape does not identify exactly one source shape")
    selected = matches[0]
    require(len(source_shapes) == len(output_shapes), f"{where} shape count changed")
    require([(row[1], row[2]) for row in source_shapes] == [(row[1], row[2]) for row in output_shapes],
            f"{where} shape identity/order changed")
    source_text = [row[3] for row in source_shapes]
    output_text = [row[3] for row in output_shapes]
    require(count_occurrences(source_text, marker) == 0,
            f"{where} source slide already contains the marker")
    require(count_occurrences(output_text, marker) == 1,
            f"{where} output slide must contain the marker exactly once")
    require(output_text[selected] == marker,
            f"{where} target shape text must equal the marker")
    for index, (before, after) in enumerate(zip(source_text, output_text)):
        if index != selected:
            require(before == after, f"{where} non-target shape text changed at index {index}")
    return {
        "format": "pptx",
        "target_slide_member": source_member,
        "target_shape_index": selected,
        "target_shape_id": source_shapes[selected][1],
        "target_shape_name": source_shapes[selected][2],
        "source_text": source_text[selected],
        "output_text": output_text[selected],
        "shape_text_count": len(output_text),
    }


def check_preservation(source: ZipArchive, output: ZipArchive, closure: frozenset[str], where: str) -> dict[str, Any]:
    changed = changed_member_names(source, output, closure)
    unexpected = changed - closure
    require(not unexpected,
            f"{where} changed members outside the fixed mutation closure: {sorted(unexpected)}")

    source_names = set(source.members)
    output_names = set(output.members)
    removed = source_names - output_names
    added = output_names - source_names

    relationship_checks: dict[str, Any] = {}
    for name in sorted(changed & closure & source_names & output_names):
        if not name.endswith(".rels"):
            continue
        source_relationships = relationship_map(source.members[name], f"{where}.source.{name}")
        output_relationships = relationship_map(output.members[name], f"{where}.output.{name}")
        expected_relationships = source_relationships
        normalization = "equal"
        if name == "xl/_rels/workbook.xml.rels" and "xl/calcChain.xml" in source_names:
            calc_edges = {
                rid: value for rid, value in source_relationships.items()
                if value[0] in CALC_CHAIN_RELATION_TYPES
                and value[1] == "calcChain.xml" and value[2] == "Internal"
            }
            require(len(calc_edges) == 1,
                    f"{where} source calcChain member lacks its exact workbook relationship")
            expected_relationships = {
                rid: value for rid, value in source_relationships.items()
                if rid not in calc_edges
            }
            normalization = "exact-calcChain-relationship-removal"
        require(output_relationships == expected_relationships,
                f"{where} authorized relationship rewrite changed graph semantics in {name!r}")
        relationship_checks[name] = {
            "source": source_relationships,
            "output": output_relationships,
            "expected": expected_relationships,
            "semantics_equal": True,
            "normalization": normalization,
        }
    # Every non-mutated member is checked as decompressed bytes.  That catches
    # lexical and binary changes alike while avoiding any claim about the
    # serializer's representation of an authorized changed XML part.
    preserved = {
        name: sha256_bytes(output.members[name])
        for name in sorted(source_names - closure)
    }
    for name in preserved:
        require(source.members[name] == output.members[name],
                f"{where} non-mutated member {name!r} is not byte-preserved")
    return {
        "changed_members": sorted(changed),
        "changed_decoded_member_names": sorted(changed),
        "removed_members": sorted(removed),
        "added_members": sorted(added),
        "authorized_mutation_closure": sorted(closure),
        "authorized_changed_members": sorted(changed & closure),
        "relationship_semantics": relationship_checks,
        "preserved_member_count": len(preserved),
        "preserved_member_sha256": preserved,
    }


def current_repo_root() -> Path:
    return ROOT.parents[3].resolve()


def frozen_repo_root(plan: Mapping[str, Any]) -> Path:
    raw = plan.get("root")
    require(isinstance(raw, str) and raw, "plan.root is missing")
    root = Path(raw)
    require(root.is_absolute(), "plan.root must be absolute")
    return root.resolve(strict=False)


def relocated_repo_path(raw: Any, plan: Mapping[str, Any], where: str) -> Path:
    """Resolve a captured repository path against the current worktree."""
    require(isinstance(raw, str) and raw, f"{where} must be a path string")
    current = current_repo_root()
    frozen = frozen_repo_root(plan)
    candidate = Path(raw)
    if not candidate.is_absolute():
        return (current / candidate).resolve(strict=False)
    normalized = candidate.resolve(strict=False)
    for root in (frozen, current):
        try:
            relative = normalized.relative_to(root)
        except ValueError:
            continue
        return (current / relative).resolve(strict=False)
    fail(f"{where} is outside both the frozen and current repository roots: {raw}")


def canonical_case_id(case: Mapping[str, Any], plan_cases: Mapping[str, Mapping[str, Any]],
                      plan: Mapping[str, Any], where: str) -> str:
    """Map frozen Rust export identities to the seven plan corpus identities."""
    raw = text(first(case, ("id", "case", "case_id", "corpus"), where), f"{where}.id")
    if raw in EXPECTED_CASES:
        return raw
    generated = {
        "generated-docx-medium": "generated-docx",
        "generated-xlsx-medium": "generated-xlsx",
        "generated-pptx-medium": "generated-pptx",
    }
    if raw in generated:
        return generated[raw]

    # Real files are numbered by export order.  Bind them by the declared
    # input path to the frozen plan path; an index or basename alone is not a
    # sufficient identity when the exporter accepts multiple files of one
    # format.
    input_raw = optional(case, ("input_path", "original_path", "source_input_path"))
    input_path = None
    if input_raw is not None:
        input_path = relocated_repo_path(input_raw, plan, f"{where}.input_path")
    matches: list[str] = []
    if input_path is not None:
        normalized = input_path.resolve(strict=False)
        for case_id, planned in plan_cases.items():
            planned_raw = planned.get("path")
            if not isinstance(planned_raw, str):
                continue
            planned_path = (current_repo_root() / planned_raw).resolve()
            if normalized == planned_path:
                matches.append(case_id)
    require(len(matches) == 1,
            f"{where} export identity {raw!r} is not bound to exactly one planned real-file path")
    return matches[0]


def expected_origin(case_id: str) -> str:
    return "generated-harness-corpus" if case_id.startswith("generated-") else "caller-named-real-file"


def format_target(case_format: str, edit: Mapping[str, Any]) -> Mapping[str, Any]:
    raw = optional(edit, ("target", "selection", "edit_target"))
    if raw is None:
        return edit
    return mapping(raw, "edit.target")


def check_case(case: Mapping[str, Any], base: Path, index: int,
               plan: Mapping[str, Any], plan_case: Mapping[str, Any]) -> dict[str, Any]:
    where = f"cases[{index}]"
    raw_case_id = text(first(case, ("id", "case", "corpus", "case_id"), where), f"{where}.id")
    case_id = raw_case_id
    # The caller has already mapped generated-medium and numbered real-file
    # identities.  Keeping this local guard catches accidental direct calls
    # with an exporter identity rather than silently changing the report key.
    require(case_id in EXPECTED_CASES, f"unexpected 0778 corpus id {case_id!r}")
    case_format = text(first(case, ("format", "kind"), where), f"{where}.format").lower()
    expected_format, expected_admitted = EXPECTED_CASES[case_id]
    require(case_format == expected_format,
            f"{where} format {case_format!r} disagrees with plan ({expected_format!r})")
    source_path, source_digest = manifest_source(case, base, where)
    planned_digest = plan_case.get("sha256")
    if planned_digest is not None:
        require(source_digest == require_sha(planned_digest, f"plan.{case_id}.sha256"),
                f"{where} source SHA-256 does not match the frozen plan")
    raw_input = optional(case, ("input_path", "original_path", "source_input_path"))
    if case_id.startswith("generated-"):
        require(raw_input is None,
                f"{where} generated corpus unexpectedly declares an input_path")
    else:
        require(isinstance(raw_input, str) and raw_input,
                f"{where} real-file corpus must declare input_path")
        # Real-file paths in the Rust manifest are repository-relative, while
        # the exported source artifact is packet-relative.  Bind both bytes
        # when the original path is available; absence is an error for a
        # declared input rather than permission to skip the binding.
        input_path = relocated_repo_path(raw_input, plan, f"{where}.input_path")
        require(input_path.is_file(), f"{where}.input_path does not resolve to a file: {input_path}")
        input_path = resolve_path(str(input_path), base, f"{where}.input_path")
        input_digest = sha256_file(input_path)
        require(input_digest == source_digest,
                f"{where} staged source differs from declared input_path SHA-256")
        planned_path = plan_case.get("path")
        require(isinstance(planned_path, str) and planned_path,
                f"plan.{case_id}.path is missing for a real-file corpus")
        require(input_path == (current_repo_root() / planned_path).resolve(),
                f"{where} input_path is not the frozen plan fixture")
    origin = optional(case, ("origin",))
    if origin is not None:
        require(origin == expected_origin(case_id),
                f"{where} origin is not bound to corpus {case_id!r}")
    admitted, marker, edit = manifest_edit(case, where)
    expected_edit = EXPECTED_EDIT_BINDINGS[case_id]
    require(case.get("edit_description") == expected_edit["description"],
            f"{where} edit_description differs from the frozen independent edit binding")
    target = format_target(case_format, edit)
    require(isinstance(target, Mapping) and dict(target) == expected_edit["target"],
            f"{where} edit_target differs from the frozen independent target binding")
    require(admitted == expected_admitted,
            f"{where} admission {admitted} disagrees with frozen corpus expectation {expected_admitted}")
    outcome = optional(case, ("edit_outcome", "outcome"))
    if outcome is not None:
        require(isinstance(outcome, str) and
                ((admitted and outcome == "admitted") or
                 (not admitted and outcome.startswith("refused:"))),
                f"{where} edit outcome is inconsistent with its frozen admission")
    if not admitted:
        require(case.get("refused_output_source_exact") is True,
                f"{where} refused/no-edit export did not record source-byte exactness")
    outputs = output_records(case, base, where)
    require("stream" in outputs,
            f"{where} exporter manifest must include the separate stream_output record")
    source = ZipArchive.read(source_path, f"{where}.source")
    target = format_target(case_format, edit)
    closure = frozenset() if not admitted else fixed_mutation_closure(case_format, source, target, where)

    corpus_manifest = optional(case, ("corpus",))
    if isinstance(corpus_manifest, Mapping):
        archive_digest = optional(corpus_manifest, ("archive_sha256",))
        if archive_digest is not None:
            require(source_digest == require_sha(archive_digest, f"{where}.corpus.archive_sha256"),
                    f"{where} source SHA-256 is not bound to its corpus identity")

    policy_rows: dict[str, Any] = {}
    first_raw: bytes | None = None
    first_sha: str | None = None
    first_projection: Any = None
    output_paths: set[Path] = set()
    for policy in POLICIES:
        path, raw_status, expected_sha, expected_bytes = outputs[policy]
        require(path not in output_paths,
                f"{where} output paths must be distinct across policy records")
        output_paths.add(path)
        status = raw_status or ("ok" if admitted else "refused")
        output = ZipArchive.read(path, f"{where}.outputs[{policy}]")
        output_sha = sha256_bytes(output.raw)
        if expected_sha is not None:
            require(output_sha == expected_sha,
                    f"{where} output {policy!r} SHA-256 differs from manifest: "
                    f"manifest {expected_sha}, actual {output_sha}")
        if expected_bytes is not None:
            require(len(output.raw) == expected_bytes,
                    f"{where} output {policy!r} byte count differs from manifest")
        if first_raw is None:
            first_raw = output.raw
            first_sha = output_sha
        else:
            require(output.raw == first_raw,
                    f"{where} policy {policy!r} archive is not byte-identical to {POLICIES[0]!r}")
        if admitted:
            require(status_is_success(status),
                    f"{where} admitted edit has non-success {policy!r} status {status!r}")
            require(path != source_path,
                    f"{where} admitted output {policy!r} points at the source archive")
            preservation = check_preservation(source, output, closure, f"{where}.{policy}")
            if case_format == "docx":
                semantic = check_docx(source, output, marker, target, f"{where}.{policy}")
            elif case_format == "xlsx":
                semantic = check_xlsx(source, output, marker, target, f"{where}.{policy}")
            elif case_format == "pptx":
                semantic = check_pptx(source, output, marker, target, f"{where}.{policy}")
            else:  # pragma: no cover - frozen cases above exhaust formats.
                fail(f"{where} unsupported format {case_format!r}")
            projection = {"preservation": preservation, "semantic": semantic}
        else:
            require(path != source_path,
                    f"{where} refusal output {policy!r} points at the source archive")
            require(output.raw == source.raw,
                    f"{where} refusal output {policy!r} is not byte-identical to the source archive")
            projection = {
                "refused": True,
                "source_archive_sha256": source_digest,
                "member_names": sorted(source.members),
            }
        if first_projection is None:
            first_projection = projection
        else:
            require(projection == first_projection,
                    f"{where} policy {policy!r} semantic/preservation projection differs")
        policy_rows[policy] = {
            "path": str(path),
            "bytes": len(output.raw),
            "sha256": output_sha,
            "status": status,
            "projection": projection,
        }
    if "stream" in outputs:
        path, raw_status, expected_sha, expected_bytes = outputs["stream"]
        require(path not in output_paths,
                f"{where} stream output path aliases a durability output")
        output_paths.add(path)
        status = raw_status or ("ok" if admitted else "refused")
        stream = ZipArchive.read(path, f"{where}.outputs['stream']")
        stream_sha = sha256_bytes(stream.raw)
        if expected_sha is not None:
            require(stream_sha == expected_sha,
                    f"{where} stream output SHA-256 differs from manifest: "
                    f"manifest {expected_sha}, actual {stream_sha}")
        if expected_bytes is not None:
            require(len(stream.raw) == expected_bytes,
                    f"{where} stream output byte count differs from manifest")
        require(stream.raw == first_raw,
                f"{where} sequential stream archive is not byte-identical to durability outputs")
        if admitted:
            require(status_is_success(status),
                    f"{where} admitted edit has non-success stream status {status!r}")
        else:
            require(stream.raw == source.raw,
                    f"{where} refusal stream output is not byte-identical to source")
        policy_rows["stream"] = {
            "path": str(path),
            "bytes": len(stream.raw),
            "sha256": stream_sha,
            "status": status,
            "projection": first_projection,
        }
    require(first_sha is not None and first_raw is not None, f"{where} has no policy output")
    declared_published = optional(case, ("published_sha256", "output_sha256"))
    if declared_published is not None:
        require(first_sha == require_sha(declared_published, f"{where}.published_sha256"),
                f"{where} published archive SHA-256 is not bound to policy output")
    declared_published_bytes = optional(case, ("published_bytes", "output_bytes"))
    if declared_published_bytes is not None:
        require(type(declared_published_bytes) is int and declared_published_bytes == len(first_raw),
                f"{where} published byte count is not bound to policy output")
    declared_source = optional(case, ("source_archive_sha256", "source_sha256"))
    if declared_source is not None:
        require(source_digest == require_sha(declared_source, f"{where}.source_archive_sha256"),
                f"{where} source archive digest is not self-consistent")
    declared_source_bytes = optional(case, ("source_archive_bytes",))
    if declared_source_bytes is not None:
        require(type(declared_source_bytes) is int and declared_source_bytes == len(source.raw),
                f"{where} source archive byte count is not bound to source bytes")
    return {
        "id": case_id,
        "corpus_id": case_id,
        "export_case_id": raw_case_id,
        "format": case_format,
        "expected_edit_admitted": expected_admitted,
        "edit_admitted": admitted,
        "admitted": admitted,
        "edit_marker": marker,
        "edit_description": expected_edit["description"],
        "edit_target": dict(expected_edit["target"]),
        "source": {"path": str(source_path), "sha256": source_digest, "bytes": len(source.raw)},
        "source_archive_sha256": source_digest,
        "source_archive_bytes": len(source.raw),
        "output_sha256": first_sha,
        "output_bytes": len(first_raw),
        "published_sha256": first_sha,
        "published_bytes": len(first_raw),
        "policy_outputs": {
            policy: {
                "path": row["path"],
                "bytes": row["bytes"],
                "sha256": row["sha256"],
            }
            for policy, row in policy_rows.items()
            if policy in POLICIES
        },
        "policies": policy_rows,
    }


def frozen_plan() -> Mapping[str, Any]:
    require(PLAN_PATH.is_file(), f"missing frozen plan {PLAN_PATH}")
    try:
        plan = json.loads(PLAN_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"cannot read frozen plan: {exc}")
    plan = mapping(plan, "plan")
    plan_cases = {row.get("id"): row for row in list_value(plan.get("corpora"), "plan.corpora") if isinstance(row, Mapping)}
    require(set(plan_cases) == set(EXPECTED_CASES),
            "frozen plan corpus ids differ from oracle's seven-case scope")
    require(set(EXPECTED_EDIT_BINDINGS) == set(EXPECTED_CASES),
            "oracle edit-binding map does not cover the frozen seven-case scope")
    for case_id, (kind, admitted) in EXPECTED_CASES.items():
        row = plan_cases[case_id]
        require(row.get("format") == kind and row.get("expected_edit_admitted") is admitted,
                f"frozen plan expectation for {case_id!r} changed")
    require(list_value(plan.get("policies"), "plan.policies") == list(POLICIES),
            "frozen plan policy order differs from oracle scope")
    return plan


def qualification_path() -> Path | None:
    for candidate in (ROOT / "qualification-identities.json",
                      ROOT / "capture-0" / "qualification-identities.json"):
        if candidate.is_file():
            return candidate
    return None


def check_qualification_binding(rows: Sequence[Mapping[str, Any]], *, required: bool) -> dict[str, Any]:
    path = qualification_path()
    if path is None:
        require(not required,
                "qualification identities are required before export admission")
        return {"status": "pending"}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"cannot read qualification identities: {exc}")
    value = mapping(value, "qualification-identities")
    require(value.get("schema") == "litchi-0778-durability-qualification-v1",
            "qualification identity schema is not the frozen capture schema")
    if "plan_sha256" in value:
        require(value.get("plan_sha256") == sha256_file(PLAN_PATH),
                "qualification identities are bound to a different plan")
    source_manifest = ROOT / "source.json"
    if "source_sha256" in value and source_manifest.is_file():
        require(value.get("source_sha256") == sha256_file(source_manifest),
                "qualification identities are bound to a different source census")
    corpora = mapping(value.get("corpora"), "qualification-identities.corpora")
    by_id = {row["id"]: row for row in rows}
    require(set(corpora) == set(by_id),
            "qualification identities do not cover exactly the seven oracle corpora")
    bound: dict[str, Any] = {}
    for case_id, row in by_id.items():
        identity = mapping(corpora[case_id], f"qualification-identities.corpora.{case_id}")
        ordinary = mapping(identity.get("ordinary_corpus"),
                           f"qualification-identities.corpora.{case_id}.ordinary_corpus")
        source_sha = require_sha(ordinary.get("source_archive_sha256"),
                                 f"qualification-identities.corpora.{case_id}.source_archive_sha256")
        published_sha = require_sha(ordinary.get("published_sha256"),
                                    f"qualification-identities.corpora.{case_id}.published_sha256")
        require(source_sha == row["source_archive_sha256"],
                f"qualification source identity differs for {case_id}")
        require(published_sha == row["published_sha256"],
                f"qualification publication identity differs for {case_id}")
        require(ordinary.get("edit_admitted") is row["expected_edit_admitted"],
                f"qualification edit admission differs for {case_id}")
        expected_edit = EXPECTED_EDIT_BINDINGS[case_id]
        require(ordinary.get("edit_description") == expected_edit["description"],
                f"qualification edit description differs for {case_id}")
        require(row.get("edit_description") == expected_edit["description"] and
                row.get("edit_target") == expected_edit["target"],
                f"oracle edit binding differs for {case_id}")
        bound[case_id] = {
            "source_archive_sha256": source_sha,
            "published_sha256": published_sha,
            "edit_admitted": ordinary["edit_admitted"],
            "edit_description": expected_edit["description"],
            "edit_target": dict(expected_edit["target"]),
        }
    return {
        "status": "bound",
        "path": str(path),
        "sha256": sha256_file(path),
        "corpora": bound,
    }


def check_manifest_output_directory(raw: Any, base: Path, plan: Mapping[str, Any]) -> None:
    require(isinstance(raw, str) and raw, "manifest.output_directory is missing")
    current = current_repo_root()
    frozen = frozen_repo_root(plan)
    candidate = Path(raw)
    if not candidate.is_absolute():
        candidate = frozen / candidate
    candidate = candidate.resolve(strict=False)
    base_resolved = base.resolve()
    for root in (frozen, current):
        try:
            relative = candidate.relative_to(root)
        except ValueError:
            continue
        relocated = (current / relative).resolve(strict=False)
        if relocated == base_resolved:
            return
    fail("manifest.output_directory is not the captured or current artifact directory")


def check_manifest(path: Path, *, require_qualification: bool = False) -> dict[str, Any]:
    plan = frozen_plan()
    plan_cases = {
        row["id"]: row
        for row in list_value(plan.get("corpora"), "plan.corpora")
        if isinstance(row, Mapping) and isinstance(row.get("id"), str)
    }
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"cannot read export manifest {path}: {exc}")
    manifest = mapping(value, "manifest")
    # The frozen exporter uses a numeric version plus a kind.  Accept it only
    # as this exact contract, never as a generic JSON report.
    require(manifest.get("schema_version") == EXPORT_SCHEMA_VERSION and
            manifest.get("kind") == EXPORT_KIND and
            manifest.get("generator") == "litchi-perf-ordinary-save-artifacts-v1",
            "manifest must identify the 0778 export schema")
    schema = SCHEMA
    raw_cases = list_value(manifest.get("cases"), "manifest.cases")
    require(len(raw_cases) == len(EXPECTED_CASES),
            f"manifest must contain exactly {len(EXPECTED_CASES)} cases")
    seen: set[str] = set()
    base = path.parent
    check_manifest_output_directory(manifest.get("output_directory"), base, plan)
    rows: list[dict[str, Any]] = []
    for index, raw_case in enumerate(raw_cases):
        case = mapping(raw_case, f"manifest.cases[{index}]")
        where = f"manifest.cases[{index}]"
        case_id = canonical_case_id(case, plan_cases, plan, where)
        require(case_id not in seen, f"manifest contains duplicate case {case_id!r}")
        seen.add(case_id)
        require(case_id in plan_cases, f"manifest case {case_id!r} is absent from the frozen plan")
        plan_case = plan_cases[case_id]
        normalized = dict(case)
        normalized["id"] = case_id
        rows.append(check_case(normalized, base, index, plan, plan_case))
    require(seen == set(EXPECTED_CASES),
            f"manifest case ids differ from frozen scope: {sorted(seen)}")
    qualification = check_qualification_binding(rows, required=require_qualification)
    return {
        "schema": SCHEMA,
        "status": "pass",
        "manifest": str(path),
        "manifest_sha256": sha256_file(path),
        "oracle_runner_sha256": sha256_file(Path(__file__).resolve()),
        "case_count": len(rows),
        "policy_count": len(POLICIES),
        "corpora": rows,
        "qualification_binding": qualification,
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--manifest",
        type=Path,
        default=ROOT / "artifacts-0" / "manifest.json",
        help="real exporter manifest (default: artifacts-0/manifest.json)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="write the passing oracle record to this path",
    )
    parser.add_argument(
        "--require-qualification",
        action="store_true",
        help="fail unless capture qualification-identities.json binds every corpus",
    )
    args = parser.parse_args(argv)
    try:
        result = check_manifest(
            args.manifest.resolve(strict=True),
            require_qualification=args.require_qualification,
        )
    except (OSError, OracleError) as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    encoded = json.dumps(result, indent=2, ensure_ascii=False, sort_keys=True)
    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded + "\n", encoding="utf-8")
    print(encoded)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
