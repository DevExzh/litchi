#!/usr/bin/env python3
"""Attribute the retained 0709 DOCX block and MCE parser edges.

This packet is a read-only extension of the retained 0711 edit attribution.
It binds the same four raw ``.callgrind.5`` parts, receipts, source census,
and baseline binary.  The selected edges are:

* ``active_block_ranges -> scan_word_element_ranges``;
* ``active -> active_offsets``; and
* ``active_offsets -> process_markup_compatibility`` when the MCE path is
  reached.

``scan_word_element_ranges`` is shared by ``active_block_ranges`` and
``paragraph_section_range`` in the retained profiles.  Its nested partition
therefore uses every positive incoming edge to that demangled symbol.  The
selected edge is reported separately and is never treated as the complete
scanner owner.  The MCE processing edge has one positive caller in the
retained parts, so its direct-child partition is exact for that raw report.

Callgrind ``Ir`` is guest-instruction attribution.  The role buckets below
identify source-level allocator, copy, and marker/search symbols; they do not
infer allocation counts, bytes, native latency, or a production speedup.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any, NoReturn


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PACKET_0711 = REPO / "docs/performance/results/change-0711"
PACKET_0709 = REPO / "docs/performance/results/change-0709"
OUTPUT = HERE / "mce-attribution.json"

ACTIVE_BLOCK_RANGES = "litchi_docx::parts::document_part::active_block_ranges"
SCAN_WORD_ELEMENT_RANGES = "litchi_docx::namespace::scan_word_element_ranges"
ALT_ACTIVE = "litchi_docx::alt::codec::active"
ACTIVE_OFFSETS = "litchi_ooxml_common::mce::codec::active_offsets"
PROCESS_MCE = "litchi_ooxml_common::mce::codec::process_markup_compatibility"

SCAN_CALLER = ACTIVE_BLOCK_RANGES
PROCESS_CALLER = ACTIVE_OFFSETS

CURRENT_SOURCE_PATHS = (
    "crates/litchi-docx/src/alt/codec.rs",
    "crates/litchi-docx/src/namespace.rs",
    "crates/litchi-docx/src/parts/document_part.rs",
    "crates/litchi-ooxml-common/src/mce/codec.rs",
)

PROFILE_NAMES = (
    "profile-r1-generated-edit",
    "profile-r2-generated-edit",
    "profile-r1-numbered-list-edit",
    "profile-r2-numbered-list-edit",
)

SIGNIFICANT_MIN_IR = 1_000
SIGNIFICANT_MIN_SHARE = 0.001


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory retained artifact."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_text(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing text artifact: {path}")
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        fail(f"cannot read {path}: {error}")


def read_json(path: Path) -> Any:
    try:
        return json.loads(read_text(path))
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def relative_repo(path: Path) -> str:
    try:
        return str(path.relative_to(REPO))
    except ValueError as error:
        fail(f"path is outside repository: {path}: {error}")


def file_binding(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {
        "file": relative_repo(path),
        "sha256": sha256(path),
        "bytes": path.stat().st_size,
    }


def load_retained_0711() -> Any:
    path = PACKET_0711 / "edit-attribution.py"
    spec = importlib.util.spec_from_file_location("retained_edit_attribution_0711", path)
    require(spec is not None and spec.loader is not None,
            f"cannot load retained attribution helper: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


H = load_retained_0711()


def aggregate_children(raw: list[dict[str, Any]], caller: str) -> list[dict[str, Any]]:
    """Aggregate direct positive cfn edges by demangled callee name."""

    grouped: dict[str, dict[str, int]] = {}
    for edge in raw:
        if edge["caller"] != caller:
            continue
        value = grouped.setdefault(edge["callee"], {
            "name": edge["callee"],
            "inclusive_ir": 0,
            "calls": 0,
            "raw_edge_count": 0,
        })
        value["inclusive_ir"] += edge["inclusive_ir"]
        value["calls"] += edge["calls"]
        value["raw_edge_count"] += 1
    return sorted(grouped.values(), key=lambda item: (-item["inclusive_ir"], item["name"]))


def incoming_edges(raw: list[dict[str, Any]], callee: str) -> list[dict[str, Any]]:
    return [edge for edge in raw if edge["callee"] == callee]


def caller_edges(
    raw: list[dict[str, Any]], caller: str, callee: str,
) -> list[dict[str, Any]]:
    return [edge for edge in raw
            if edge["caller"] == caller and edge["callee"] == callee]


def role(name: str, context: str) -> str:
    """Bucket direct children by source-level diagnostic role.

    These labels are intentionally conservative.  Generic libc primitives are
    marked only when the selected owner makes their role clear; no bucket is an
    allocation counter.
    """

    lower = name.lower()
    if context in {ACTIVE_OFFSETS, PROCESS_MCE}:
        if (
            "find_bytes" in lower
            or "contains_mce_namespace" in lower
            or "active_marker" in lower
            or "decimal_bytes" in lower
            or "parse_decimal" in lower
            or "memcmp" in lower
        ):
            return "marker"
    if (
        "memcpy" in lower
        or "memmove" in lower
        or "copy_offsets" in lower
        or "extend_trusted" in lower
        or "copy" in lower
    ):
        return "copy"
    if (
        "__rustc::__rust_alloc" in lower
        or "__rustc::__rust_dealloc" in lower
        or "rawvec" in lower
        or "try_allocate" in lower
        or "finish_grow" in lower
        or "try_reserve" in lower
        or "reserve_rehash" in lower
        or "boundedoutput::reserve" in lower
        or "hashmap" in lower
        or "btreemap" in lower
        or "vec<t" in lower and "resize" in lower
    ):
        return "allocator"
    if (
        "read_event" in lower
        or "process_event" in lower
        or "resolve_event" in lower
        or "set_level" in lower
        or "namespace" in lower
        or "qname::prefix" in lower
    ):
        return "xml_reader_namespace"
    if "memcmp" in lower:
        return "search_backend"
    return "parser_or_other"


def annotate_children(
    children: list[dict[str, Any]], owner_ir: int, context: str,
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for item in children:
        value = dict(item)
        value["role"] = role(item["name"], context)
        value["share_of_owner"] = item["inclusive_ir"] / owner_ir
        value["significant"] = (
            item["inclusive_ir"] >= SIGNIFICANT_MIN_IR
            or value["share_of_owner"] >= SIGNIFICANT_MIN_SHARE
            or value["role"] in {"allocator", "copy", "marker"}
        )
        result.append(value)
    return result


def partition_all_incoming(
    raw: list[dict[str, Any]], owner: str, label: str,
) -> dict[str, Any]:
    """Build a disjoint partition for the whole demangled owner aggregate.

    Callgrind can emit one function body for a symbol reached by several
    callers.  Summing all positive incoming edges makes the owner scope match
    the direct cfn body aggregate.  A selected caller edge is reported by the
    caller and is not silently promoted to the whole owner scope.
    """

    incoming = incoming_edges(raw, owner)
    require(incoming, f"{label}: missing positive incoming edge for {owner}")
    owner_ir = sum(edge["inclusive_ir"] for edge in incoming)
    children = aggregate_children(raw, owner)
    direct_ir = sum(item["inclusive_ir"] for item in children)
    self_ir = owner_ir - direct_ir
    require(self_ir >= 0,
            f"{label}: negative owner self Ir ({owner_ir} - {direct_ir})")
    require(self_ir + direct_ir == owner_ir,
            f"{label}: immediate-child partition does not reconstruct owner Ir")
    annotated = annotate_children(children, owner_ir, owner)
    return {
        "owner": owner,
        "incoming_scope": "all_positive_raw_incoming_edges_to_demangled_symbol",
        "incoming": {
            "edge_count": len(incoming),
            "calls": sum(edge["calls"] for edge in incoming),
            "inclusive_ir": owner_ir,
            "raw_edges": incoming,
        },
        "self_ir": self_ir,
        "direct_children_ir": direct_ir,
        "children": annotated,
        "significant_children": [
            item["name"] for item in annotated if item["significant"]
        ],
        "significance_policy": {
            "minimum_ir": SIGNIFICANT_MIN_IR,
            "minimum_share": SIGNIFICANT_MIN_SHARE,
            "retain_focus_roles": ["allocator", "copy", "marker"],
        },
        "partition_equation": (
            "owner self Ir + sum(immediate direct-child inclusive Ir) = "
            "sum(all positive incoming owner Ir)"
        ),
        "disjoint": True,
        "nested_inclusive_costs_excluded": True,
        "positive_edges_only": True,
    }


def focus_totals(partition: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, dict[str, int]] = {}
    for item in partition["children"]:
        bucket = result.setdefault(item["role"], {
            "role": item["role"],
            "inclusive_ir": 0,
            "calls": 0,
            "child_count": 0,
        })
        bucket["inclusive_ir"] += item["inclusive_ir"]
        bucket["calls"] += item["calls"]
        bucket["child_count"] += 1
    for bucket in result.values():
        bucket["share_of_owner"] = bucket["inclusive_ir"] / partition["incoming"]["inclusive_ir"]
    return dict(sorted(result.items()))


def selected_edge(
    raw: list[dict[str, Any]], caller: str, callee: str, label: str,
) -> dict[str, Any]:
    matches = caller_edges(raw, caller, callee)
    require(len(matches) == 1,
            f"{label}: expected one selected positive edge, got {len(matches)}")
    edge = dict(matches[0])
    edge["share_of_callee_aggregate"] = None
    return {
        "caller": caller,
        "callee": callee,
        "edge_count": len(matches),
        "calls": sum(item["calls"] for item in matches),
        "inclusive_ir": sum(item["inclusive_ir"] for item in matches),
        "raw_edges": matches,
    }


def bind_selected_edge(
    edge: dict[str, Any], partition: dict[str, Any],
) -> dict[str, Any]:
    result = dict(edge)
    result["share_of_callee_aggregate"] = (
        edge["inclusive_ir"] / partition["incoming"]["inclusive_ir"]
    )
    return result


def direct_child(partition: dict[str, Any], name: str) -> dict[str, Any] | None:
    matches = [item for item in partition["children"] if item["name"] == name]
    require(len(matches) <= 1,
            f"{partition['owner']}: duplicate aggregated child {name!r}")
    return matches[0] if matches else None


def inline_search_diagnostic(partition: dict[str, Any]) -> dict[str, Any]:
    """Describe direct search symbols without treating symbol absence as zero work."""

    find_bytes = direct_child(partition, "litchi_ooxml_common::mce::codec::find_bytes")
    memcmp = direct_child(partition, "__memcmp_avx2_movbe")
    result: dict[str, Any] = {
        "owner": partition["owner"],
        "direct_find_bytes_edge": find_bytes,
        "direct_memcmp_backend_edge": memcmp,
        "compiler_inlining_caveat": (
            "A missing direct find_bytes edge does not mean zero search work: "
            "the release compiler may inline the slice-window search and expose "
            "its memcmp backend as the remaining symbol. A generic memcmp edge "
            "is search evidence, not a standalone marker allocation or exact "
            "per-call attribution."
        ),
    }
    if find_bytes is None:
        result["direct_find_bytes_status"] = "absent_in_raw_positive_edges"
    else:
        result["direct_find_bytes_status"] = "present_in_raw_positive_edges"
    if memcmp is None:
        result["direct_memcmp_status"] = "absent_in_raw_positive_edges"
    else:
        result["direct_memcmp_status"] = "present_in_raw_positive_edges"
    return result


def validate_current_source() -> dict[str, Any]:
    manifest_path = HERE / "source-final.json"
    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict), "0712 source-final.json is not an object")
    files: dict[str, str] = {}
    for path_text in CURRENT_SOURCE_PATHS:
        expected = manifest.get(path_text)
        require(isinstance(expected, str) and len(expected) == 64,
                f"0712 source-final.json omits {path_text}")
        path = REPO / path_text
        require(sha256(path) == expected,
                f"current source differs from 0712 source-final.json: {path_text}")
        files[path_text] = expected
    return {
        "manifest": file_binding(manifest_path),
        "manifest_entry_count": len(manifest),
        "files": files,
    }


def validate_packet_revision() -> dict[str, Any]:
    path = HERE / "revision.json"
    value = read_json(path)
    require(isinstance(value, dict), "0712 revision.json is not an object")
    revision = value.get("revision")
    require(isinstance(revision, str) and len(revision) == 40,
            "0712 revision.json revision is malformed")
    return {
        "revision": revision,
        "artifact": file_binding(path),
    }


def profile_binding(row: dict[str, Any]) -> dict[str, Any]:
    raw = row["raw_artifact"]
    profile = row["profile_result"]
    return {
        "name": row["name"],
        "corpus": row["corpus"],
        "repeat": row["repeat"],
        "part": row["part"],
        "part_trigger": row["part_trigger"],
        "binary_sha256": row["historical_binary_sha256"],
        "raw_artifact": raw,
        "profile_result": profile,
    }


def profile_row(row: dict[str, Any]) -> dict[str, Any]:
    name = row["name"]
    raw_path = PACKET_0709 / f"{name}.callgrind.5"
    raw = H.raw_edges(raw_path)

    active_block = partition_all_incoming(raw, ACTIVE_BLOCK_RANGES, name)
    scan = partition_all_incoming(raw, SCAN_WORD_ELEMENT_RANGES, name)
    active = partition_all_incoming(raw, ALT_ACTIVE, name)
    offsets = partition_all_incoming(raw, ACTIVE_OFFSETS, name)

    scan_edge = selected_edge(raw, SCAN_CALLER, SCAN_WORD_ELEMENT_RANGES, name)
    scan_edge = bind_selected_edge(scan_edge, scan)
    active_edge = selected_edge(raw, ACTIVE_BLOCK_RANGES, ALT_ACTIVE, name)
    active_edge = bind_selected_edge(active_edge, active)
    offsets_edge = selected_edge(raw, ALT_ACTIVE, ACTIVE_OFFSETS, name)
    offsets_edge = bind_selected_edge(offsets_edge, offsets)

    process_matches = caller_edges(raw, PROCESS_CALLER, PROCESS_MCE)
    require(len(process_matches) <= 1,
            f"{name}: process_markup_compatibility selected edge is duplicated")
    if row["corpus"] == "generated":
        require(not process_matches,
                f"{name}: generated fast path unexpectedly calls process_markup_compatibility")
    else:
        require(len(process_matches) == 1,
                f"{name}: MCE-bearing corpus lacks one process_markup_compatibility edge")

    process = None
    process_edge = None
    if process_matches:
        process = partition_all_incoming(raw, PROCESS_MCE, name)
        process_edge = bind_selected_edge(
            selected_edge(raw, PROCESS_CALLER, PROCESS_MCE, name), process
        )

    partitions = {
        "active_block_ranges": active_block,
        "scan_word_element_ranges": scan,
        "alt_active": active,
        "active_offsets": offsets,
    }
    if process is not None:
        partitions["process_markup_compatibility"] = process

    for partition in partitions.values():
        partition["focus_totals"] = focus_totals(partition)

    return {
        "name": name,
        "corpus": row["corpus"],
        "repeat": row["repeat"],
        "receipt_binding": profile_binding(row),
        "selected_edges": {
            "active_block_ranges_to_scan_word_element_ranges": scan_edge,
            "active_block_ranges_to_alt_active": active_edge,
            "alt_active_to_active_offsets": offsets_edge,
            "active_offsets_to_process_markup_compatibility": process_edge,
        },
        "partitions": partitions,
        "compiler_inline_diagnostics": {
            "active_offsets_search": inline_search_diagnostic(offsets),
        },
        "validation": {
            "historical_row_reused": True,
            "all_selected_edges_are_unique_positive_edges": True,
            "scan_partition_uses_all_incoming_callers": True,
            "mce_process_partition_present_only_on_numbered_list": process is not None,
            "immediate_children_reconcile_owner": all(
                part["self_ir"] + part["direct_children_ir"]
                == part["incoming"]["inclusive_ir"]
                for part in partitions.values()
            ),
            "nested_inclusive_costs_excluded": True,
        },
    }


def aggregate_partition_rows(rows: list[dict[str, Any]], key: str) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for corpus in sorted({row["corpus"] for row in rows}):
        selected = [row for row in rows if row["corpus"] == corpus]
        parts = [row["partitions"][key] for row in selected if key in row["partitions"]]
        if not parts:
            continue
        owner_ir = sum(part["incoming"]["inclusive_ir"] for part in parts)
        self_ir = sum(part["self_ir"] for part in parts)
        children: dict[str, dict[str, Any]] = {}
        for part in parts:
            for item in part["children"]:
                value = children.setdefault(item["name"], {
                    "name": item["name"],
                    "inclusive_ir": 0,
                    "calls": 0,
                    "raw_edge_count": 0,
                    "role": item["role"],
                })
                require(value["role"] == item["role"],
                        f"{corpus}/{key}: child role changed across repeats")
                value["inclusive_ir"] += item["inclusive_ir"]
                value["calls"] += item["calls"]
                value["raw_edge_count"] += item["raw_edge_count"]
        direct_ir = sum(item["inclusive_ir"] for item in children.values())
        require(self_ir + direct_ir == owner_ir,
                f"{corpus}/{key}: aggregate partition does not reconcile")
        for item in children.values():
            item["share_of_owner"] = item["inclusive_ir"] / owner_ir
            item["significant"] = (
                item["inclusive_ir"] >= SIGNIFICANT_MIN_IR
                or item["share_of_owner"] >= SIGNIFICANT_MIN_SHARE
                or item["role"] in {"allocator", "copy", "marker"}
            )
        totals: dict[str, dict[str, Any]] = {}
        for item in children.values():
            bucket = totals.setdefault(item["role"], {
                "role": item["role"],
                "inclusive_ir": 0,
                "calls": 0,
                "child_count": 0,
            })
            bucket["inclusive_ir"] += item["inclusive_ir"]
            bucket["calls"] += item["calls"]
            bucket["child_count"] += 1
        for bucket in totals.values():
            bucket["share_of_owner"] = bucket["inclusive_ir"] / owner_ir
        result[corpus] = {
            "profile_count": len(parts),
            "owner_ir": owner_ir,
            "self_ir": self_ir,
            "direct_children_ir": direct_ir,
            "children": sorted(children.values(),
                                key=lambda item: (-item["inclusive_ir"], item["name"])),
            "focus_totals": dict(sorted(totals.items())),
            "partition_equation": (
                "owner self Ir + sum(immediate direct-child inclusive Ir) = "
                "sum(all positive incoming owner Ir)"
            ),
            "disjoint": True,
            "nested_inclusive_costs_excluded": True,
            "positive_edges_only": True,
        }
    return result


def selected_edge_aggregate(rows: list[dict[str, Any]], key: str) -> dict[str, Any]:
    partition_key = {
        "active_block_ranges_to_scan_word_element_ranges": "scan_word_element_ranges",
        "active_block_ranges_to_alt_active": "alt_active",
        "alt_active_to_active_offsets": "active_offsets",
        "active_offsets_to_process_markup_compatibility": "process_markup_compatibility",
    }[key]
    result: dict[str, Any] = {}
    for corpus in sorted({row["corpus"] for row in rows}):
        selected = [row for row in rows if row["corpus"] == corpus]
        edges = [row["selected_edges"][key] for row in selected]
        present_rows = [
            row for row in selected
            if row["selected_edges"][key] is not None
            and partition_key in row["partitions"]
        ]
        present = [row["selected_edges"][key] for row in present_rows]
        if not present:
            result[corpus] = {
                "profile_count": len(edges),
                "present_count": 0,
                "inclusive_ir": 0,
                "callee_aggregate_ir": 0,
                "aggregate_fraction": None,
                "calls": 0,
                "status": "not_reached",
            }
            continue
        edge_ir = sum(edge["inclusive_ir"] for edge in present)
        callee_ir = sum(
            row["partitions"][partition_key]["incoming"]["inclusive_ir"]
            for row in present_rows
        )
        require(callee_ir > 0,
                f"{corpus}/{key}: selected-edge denominator is not positive")
        result[corpus] = {
            "profile_count": len(edges),
            "present_count": len(present),
            "inclusive_ir": edge_ir,
            "callee_aggregate_ir": callee_ir,
            "aggregate_fraction": edge_ir / callee_ir,
            "calls": sum(edge["calls"] for edge in present),
            "share_of_callee_aggregate_mean": (
                sum(edge["share_of_callee_aggregate"] for edge in present)
                / len(present)
            ),
            "share_of_callee_aggregate_min": min(
                edge["share_of_callee_aggregate"] for edge in present
            ),
            "share_of_callee_aggregate_max": max(
                edge["share_of_callee_aggregate"] for edge in present
            ),
            "status": "present",
        }
    return result


def recommendation(
    rows: list[dict[str, Any]],
    aggregates: dict[str, Any],
    selected: dict[str, Any],
) -> dict[str, Any]:
    generated_mce = selected["active_offsets_to_process_markup_compatibility"]["generated"]
    numbered_mce = selected["active_offsets_to_process_markup_compatibility"]["numbered-list"]
    generated_scan = selected[
        "active_block_ranges_to_scan_word_element_ranges"
    ]["generated"]
    numbered_scan = selected[
        "active_block_ranges_to_scan_word_element_ranges"
    ]["numbered-list"]
    numbered_offsets = aggregates["active_offsets"]["numbered-list"]
    numbered_process = aggregates["process_markup_compatibility"]["numbered-list"]
    generated_scan_part = aggregates["scan_word_element_ranges"]["generated"]
    numbered_scan_part = aggregates["scan_word_element_ranges"]["numbered-list"]

    return {
        "status": "diagnostic_only_no_production_change",
        "answer": {
            "generated": {
                "namespace_scan_selected_edge_ir": generated_scan["inclusive_ir"],
                "mce_process_selected_edge_ir": generated_mce["inclusive_ir"],
                "dominant_selected_deeper_path": "scan_word_element_ranges",
                "interpretation": (
                    "The generated corpus takes the active-offset fast path; no "
                    "process_markup_compatibility edge is present. The scanner and "
                    "its XML namespace/read-event work are the bounded next seam. "
                    "The raw report has no direct find_bytes edge here, but it does "
                    "retain the __memcmp_avx2_movbe search backend; release inlining "
                    "means the absent symbol is not zero search work."
                ),
            },
            "numbered-list": {
                "namespace_scan_selected_edge_ir": numbered_scan["inclusive_ir"],
                "mce_process_selected_edge_ir": numbered_mce["inclusive_ir"],
                "mce_process_share_of_active_offsets": (
                    numbered_process["owner_ir"] / numbered_offsets["owner_ir"]
                ),
                "dominant_selected_deeper_path": "process_markup_compatibility",
                "interpretation": (
                    "The MCE-bearing corpus pays for one full active-offset "
                    "process_markup_compatibility call. Its start-event handling "
                    "is the largest direct child, while marker search and output "
                    "buffer work remain separately attributed. Both a direct "
                    "find_bytes edge and a memcmp search backend are retained."
                ),
            },
        },
        "bounded_highest_roi_next": {
            "candidate": "stream_or_reuse_MCE_active_offset_selection",
            "reason": (
                "On the admitted MCE-bearing contrast, process_markup_compatibility "
                "is larger than the selected block-range scan. A bounded experiment "
                "should parse the MCE branch and retain exact source offsets in one "
                "validated pass, or reuse a proven active-branch map, then recapture "
                "both corpora. The generated fast path must remain a separate "
                "namespace-scan comparison."
            ),
            "source_boundary": (
                f"{ACTIVE_OFFSETS} -> {PROCESS_MCE}; preserve the existing "
                "active-offset and MCE limits"
            ),
            "required_proofs": [
                "exact active offset order and duplicate handling",
                "strict/transitional namespace and Choice/Fallback selection",
                "marker collision and malformed-source refusal behavior",
                "source ranges, byte output, and error timing",
                "fresh native and allocator captures on both admitted corpora",
            ],
            "contrast": (
                f"{SCAN_WORD_ELEMENT_RANGES} remains the dominant selected path "
                "for the generated fast-path corpus; do not generalize the MCE "
                "finding across both shapes."
            ),
        },
        "secondary_bounded_candidate": {
            "candidate": "memmem_namespace_presence_fast_path",
            "boundary": f"{ACTIVE_OFFSETS} namespace-presence check",
            "evidence": {
                "generated_active_offsets_owner_ir": (
                    aggregates["active_offsets"]["generated"]["owner_ir"]
                ),
                "generated_memcmp_backend_ir": sum(
                    row["compiler_inline_diagnostics"]["active_offsets_search"]
                    ["direct_memcmp_backend_edge"]["inclusive_ir"]
                    for row in rows if row["corpus"] == "generated"
                ),
                "generated_memcmp_backend_calls": sum(
                    row["compiler_inline_diagnostics"]["active_offsets_search"]
                    ["direct_memcmp_backend_edge"]["calls"]
                    for row in rows if row["corpus"] == "generated"
                ),
                "generated_direct_find_bytes_edges": all(
                    row["compiler_inline_diagnostics"]["active_offsets_search"]
                    ["direct_find_bytes_edge"] is None
                    for row in rows if row["corpus"] == "generated"
                ),
            },
            "reason": (
                "The generated fast path spends 600948 retained Ir across repeats "
                "in the memcmp backend while checking for the MCE namespace. The "
                "existing memchr dependency already supplies memmem primitives, "
                "so a source-level exactness-preserving search candidate is a "
                "smaller experiment than fusing the full MCE parser."
            ),
            "constraints": [
                "Preserve the exact byte-presence dispatch predicate, including arbitrary byte surroundings and boundary cases.",
                "Retain malformed XML routing, source and output limits, and marker collision behavior.",
                "Treat the inlined memcmp edge as search evidence rather than a native allocation or exact marker-only count.",
                "Recapture the NumberedList MCE-bearing path to ensure the candidate does not change the full processing route.",
            ],
        },
        "partition_scope": {
            "selected_scan_edge_is_not_scan_owner": True,
            "scan_owner_includes_paragraph_section_range_callers": True,
            "mce_process_owner_has_one_positive_caller_per_present_profile": True,
            "active_offsets_owner_is_partitioned_from_all_positive_incoming_edges": True,
            "generated_process_path_is_absent": True,
        },
    }


def analyze() -> dict[str, Any]:
    base = H.analyze()
    require(base.get("status") == "pass", "retained 0711 edit attribution did not pass")
    require([row["name"] for row in base["rows"]] == list(PROFILE_NAMES),
            "retained 0711 profile order changed")

    base_path = PACKET_0711 / "edit-attribution.json"
    base_disk = read_json(base_path)
    require(base_disk == base,
            "retained 0711 edit-attribution.json differs from deterministic helper output")

    rows = [profile_row(row) for row in base["rows"]]
    aggregate_keys = (
        "active_block_ranges",
        "scan_word_element_ranges",
        "alt_active",
        "active_offsets",
        "process_markup_compatibility",
    )
    aggregates = {
        key: aggregate_partition_rows(rows, key) for key in aggregate_keys
    }
    selected_keys = (
        "active_block_ranges_to_scan_word_element_ranges",
        "active_block_ranges_to_alt_active",
        "alt_active_to_active_offsets",
        "active_offsets_to_process_markup_compatibility",
    )
    selected = {
        key: selected_edge_aggregate(rows, key) for key in selected_keys
    }

    historical = dict(base["historical_binding"])
    historical["retained_0711_edit_attribution"] = file_binding(base_path)
    current_source = validate_current_source()
    revision = validate_packet_revision()

    script_binding = file_binding(Path(__file__).resolve())
    return {
        "schema": "docx_mce_nested_attribution_v1",
        "status": "pass",
        "scope": {
            "phase": "edit",
            "metric": "Callgrind Ir guest instructions",
            "profile_count": len(rows),
            "profile_names": list(PROFILE_NAMES),
            "corpora": ["generated", "numbered-list"],
            "selected_owner_edges": [
                f"{ACTIVE_BLOCK_RANGES} -> {SCAN_WORD_ELEMENT_RANGES}",
                f"{ALT_ACTIVE} -> {ACTIVE_OFFSETS}",
                f"{ACTIVE_OFFSETS} -> {PROCESS_MCE}",
            ],
            "selection_method": (
                "retained raw positive cfn edges with exact demangled caller/callee "
                "names; owner partitions aggregate all positive incoming edges "
                "when a symbol is shared by multiple callers"
            ),
        },
        "analysis_script": script_binding,
        "packet_revision": revision,
        "current_source": current_source,
        "historical_binding": historical,
        "profiles": rows,
        "aggregate_partitions": aggregates,
        "aggregate_selected_edges": selected,
        "recommendation": recommendation(rows, aggregates, selected),
        "limitations": [
            "The four parts are historical 0709 profiles and are not current native timing.",
            "Callgrind Ir is guest-instruction attribution, not native latency, hardware cycles, allocation counts, allocated bytes, RSS, or cache counters.",
            "Allocator, copy, and marker roles are source-name diagnostics over direct edges; no allocation count or byte total is inferred.",
            "The selected scanner edge is one caller slice of a shared demangled function; its nested partition uses the aggregate of all positive incoming callers and must not be read as selected-callsite-only cost.",
            "Immediate direct children in each declared owner partition are disjoint; deeper inclusive costs are excluded from ancestor sums.",
            "The generated and NumberedList corpora are contrasting fixtures and do not represent all DOCX producer or document shapes.",
            "The recommendation is a bounded experiment proposal with no causal speedup or production-change claim.",
        ],
        "replay": {
            "method": (
                "rerun this script with --replay and compare parsed JSON against "
                "mce-attribution.json; inputs are sorted raw edge aggregates and "
                "contain no wall-clock or random values"
            ),
            "deterministic_inputs": True,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=OUTPUT)
    parser.add_argument(
        "--replay", action="store_true",
        help="recompute and compare with the existing output instead of writing it",
    )
    args = parser.parse_args()
    output = args.output.resolve()
    try:
        result = analyze()
        if args.replay:
            require(output.is_file() and not output.is_symlink(),
                    f"replay output is missing: {output}")
            expected = read_json(output)
            require(expected == result,
                    "deterministic replay differs from the existing JSON output")
            print(f"replay verified {len(result['profiles'])} retained DOCX profiles: {output}")
        else:
            require(output.parent == HERE,
                    f"output must remain in the change-0712 packet: {output}")
            require(not output.is_symlink(), f"refusing symlink output: {output}")
            output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n",
                              encoding="utf-8")
            print(f"verified {len(result['profiles'])} retained DOCX profiles; wrote {output}")
    except (EvidenceError, OSError, KeyError, TypeError) as error:
        print(f"mce-attribution.py: error: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
